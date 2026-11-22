"""One-click research, word lookup, and offline grammar companion for Word."""

from __future__ import annotations

import csv
import html
import json
import os
import re
import shutil
import subprocess
import threading
import time
import urllib.error
import urllib.parse
import urllib.request
import webbrowser
import xml.etree.ElementTree as ET
from datetime import datetime
from pathlib import Path
from tkinter import END, BOTH, LEFT, RIGHT, VERTICAL, W, X, Y, BooleanVar, StringVar, Tk, filedialog, messagebox, ttk
from tkinter.scrolledtext import ScrolledText


APP_DIR = Path(__file__).resolve().parent
DEFAULT_TEMPLATE = "Senior Project template_Final Version.docx"
ROOT = APP_DIR.parent if (APP_DIR.parent / DEFAULT_TEMPLATE).is_file() else APP_DIR
SOURCE_DOC = ROOT / "Senior Project template_Final Version.docx"
WORK_DOC = ROOT / "Senior Project Working Draft.docx"
PROPOSAL_DOC = ROOT / "02_Two_Port_EV_Charging_Proposal.pdf"
VENDOR = APP_DIR / "vendor"
SETTINGS_FILE = APP_DIR / "local_settings.json"
RESULTS_DIR = ROOT / "Research Results"
LT_URL = "http://127.0.0.1:8081/v2/check"
HEADERS = {"User-Agent": "KirasPawn/1.0 (personal academic research)", "Accept": "application/json"}
BG = "#0A0F0B"
PANEL = "#101A13"
FIELD = "#0D160F"
GREEN = "#A5F5B6"
BRIGHT = "#70F398"
MUTED = "#8CA896"
BORDER = "#315B3C"
SELECT = "#235C36"
FONT = "Consolas"


def get_json(url: str, timeout: int = 15):
    request = urllib.request.Request(url, headers=HEADERS)
    with urllib.request.urlopen(request, timeout=timeout) as response:
        return json.load(response)


def crossref_search(query: str) -> list[dict]:
    url = "https://api.crossref.org/works?" + urllib.parse.urlencode(
        {"query.bibliographic": query, "rows": 8, "select": "DOI,title,author,published,URL,abstract,container-title"}
    )
    items = get_json(url)["message"]["items"]
    out = []
    for item in items:
        title = (item.get("title") or [""])[0]
        if not title:
            continue
        authors = ", ".join(
            " ".join(filter(None, (person.get("given"), person.get("family"))))
            for person in item.get("author", [])[:4]
        )
        date_parts = item.get("published", {}).get("date-parts", [[]])[0]
        year = str(date_parts[0]) if date_parts else ""
        abstract = re.sub("<[^>]+>", " ", item.get("abstract", ""))
        out.append({"source": "Crossref", "year": year, "title": html.unescape(title),
                    "authors": authors, "url": item.get("URL", ""), "doi": item.get("DOI", ""),
                    "details": html.unescape(abstract).strip(), "alphaxiv": ""})
    return out


def arxiv_search(query: str) -> list[dict]:
    terms = re.findall(r"[\w-]+", query)[:4]
    search = " AND ".join(f"all:{term}" for term in terms)
    url = "https://export.arxiv.org/api/query?" + urllib.parse.urlencode(
        {"search_query": search, "start": 0, "max_results": 8, "sortBy": "relevance", "sortOrder": "descending"}
    )
    request = urllib.request.Request(url, headers={"User-Agent": HEADERS["User-Agent"]})
    with urllib.request.urlopen(request, timeout=20) as response:
        root = ET.fromstring(response.read())
    ns = {"a": "http://www.w3.org/2005/Atom"}
    out = []
    for entry in root.findall("a:entry", ns):
        link = (entry.findtext("a:id", "", ns) or "").strip()
        paper_id = link.rstrip("/").split("/")[-1]
        out.append({"source": "arXiv", "year": entry.findtext("a:published", "", ns)[:4],
                    "title": " ".join(entry.findtext("a:title", "", ns).split()),
                    "authors": ", ".join(a.findtext("a:name", "", ns) for a in entry.findall("a:author", ns)[:4]),
                    "url": link, "doi": "", "details": " ".join(entry.findtext("a:summary", "", ns).split()),
                    "alphaxiv": f"https://www.alphaxiv.org/abs/{paper_id}" if paper_id else ""})
    return out


def github_search(query: str) -> list[dict]:
    url = "https://api.github.com/search/repositories?" + urllib.parse.urlencode(
        {"q": query, "sort": "best-match", "per_page": 6}
    )
    items = get_json(url)["items"]
    return [{"source": "GitHub", "year": item.get("created_at", "")[:4],
             "title": item.get("full_name", ""), "authors": item.get("owner", {}).get("login", ""),
             "url": item.get("html_url", ""), "doi": "", "details": item.get("description") or "",
             "alphaxiv": ""} for item in items]


def word_lookup(word: str) -> str:
    result = []
    try:
        data = get_json("https://api.dictionaryapi.dev/api/v2/entries/en/" + urllib.parse.quote(word))
        for entry in data[:2]:
            for meaning in entry.get("meanings", [])[:3]:
                result.append(meaning.get("partOfSpeech", ""))
                for definition in meaning.get("definitions", [])[:3]:
                    result.append("  • " + definition.get("definition", ""))
                    if definition.get("example"):
                        result.append("    Example: " + definition["example"])
    except Exception as exc:
        result.append(f"Definition unavailable: {exc}")
    try:
        synonyms = get_json("https://api.datamuse.com/words?" + urllib.parse.urlencode(
            {"rel_syn": word, "max": 12}))
        if synonyms:
            result.append("\nRelated words: " + ", ".join(x["word"] for x in synonyms))
    except Exception as exc:
        result.append(f"Related words unavailable: {exc}")
    return "\n".join(result)


def grammar_check(text: str) -> list[dict]:
    data = urllib.parse.urlencode({"text": text, "language": "en-US"}).encode()
    request = urllib.request.Request(LT_URL, data=data, headers={"Content-Type": "application/x-www-form-urlencoded"})
    with urllib.request.urlopen(request, timeout=25) as response:
        return json.load(response).get("matches", [])


class WritingMode:
    def __init__(self):
        self.root = Tk()
        self.root.title("Kira's Pawn")
        self.root.geometry("1060x760")
        self.root.minsize(820, 600)
        self._theme()
        self.results: list[dict] = []
        self.matches: list[dict] = []
        self.checked_text = ""
        self.grammar_generation = 0
        self.lt_process = None
        self.lt_log = None
        self.lt_ready = False
        self.grammar_timer = None
        self.source_doc = SOURCE_DOC if SOURCE_DOC.is_file() else None
        self.work_doc = WORK_DOC if self.source_doc else None
        if self.source_doc is None and SETTINGS_FILE.is_file():
            try:
                remembered = Path(json.loads(SETTINGS_FILE.read_text(encoding="utf-8"))["source_doc"])
                if remembered.is_file():
                    self.source_doc = remembered
                    self.work_doc = remembered.with_name(remembered.stem + " Working Draft.docx")
            except (OSError, ValueError, KeyError):
                pass
        self.doc_label = StringVar(value=str(self.work_doc) if self.work_doc else "Choose a Word template to begin")
        self.status = StringVar(value="Ready. Search needs internet; grammar runs on this computer.")
        self.grammar_status = StringVar(value="Starting local LanguageTool…")
        self.auto_check = BooleanVar(value=True)
        self._build_ui()
        self.root.protocol("WM_DELETE_WINDOW", self.close)
        self.open_word()
        threading.Thread(target=self.start_languagetool, daemon=True).start()

    def _theme(self):
        self.root.configure(background=BG)
        style = ttk.Style(self.root)
        style.theme_use("clam")
        style.configure("TFrame", background=BG)
        style.configure("TLabel", background=BG, foreground=GREEN, font=(FONT, 10))
        style.configure("Banner.TLabel", background=BG, foreground=BRIGHT, font=(FONT, 17, "bold"))
        style.configure("Muted.TLabel", background=BG, foreground=MUTED, font=(FONT, 9))
        style.configure("TButton", background=PANEL, foreground=BRIGHT, bordercolor=BORDER,
                        lightcolor=BORDER, darkcolor=BORDER, font=(FONT, 9, "bold"), padding=(9, 6))
        style.map("TButton", background=[("active", SELECT), ("pressed", SELECT)],
                  foreground=[("disabled", MUTED), ("active", "#FFFFFF")])
        style.configure("TEntry", fieldbackground=FIELD, foreground=GREEN, insertcolor=BRIGHT,
                        bordercolor=BORDER, lightcolor=BORDER, darkcolor=BORDER, padding=5,
                        font=(FONT, 10))
        style.configure("TCheckbutton", background=BG, foreground=GREEN, font=(FONT, 9))
        style.map("TCheckbutton", background=[("active", BG)], foreground=[("active", BRIGHT)])
        style.configure("TNotebook", background=BG, bordercolor=BORDER)
        style.configure("TNotebook.Tab", background=PANEL, foreground=MUTED,
                        bordercolor=BORDER, font=(FONT, 10, "bold"), padding=(15, 8))
        style.map("TNotebook.Tab", background=[("selected", SELECT), ("active", PANEL)],
                  foreground=[("selected", "#FFFFFF"), ("active", BRIGHT)])
        style.configure("Treeview", background=FIELD, fieldbackground=FIELD, foreground=GREEN,
                        bordercolor=BORDER, rowheight=26, font=(FONT, 9))
        style.map("Treeview", background=[("selected", SELECT)], foreground=[("selected", "#FFFFFF")])
        style.configure("Treeview.Heading", background=PANEL, foreground=BRIGHT,
                        bordercolor=BORDER, font=(FONT, 9, "bold"), padding=5)
        style.map("Treeview.Heading", background=[("active", SELECT)])
        style.configure("Vertical.TScrollbar", background=PANEL, troughcolor=BG,
                        bordercolor=BORDER, arrowcolor=BRIGHT)

    def _style_text(self, widget):
        widget.configure(background=FIELD, foreground=GREEN, insertbackground=BRIGHT,
                         selectbackground=SELECT, selectforeground="#FFFFFF", relief="flat",
                         borderwidth=1, highlightthickness=1, highlightbackground=BORDER,
                         highlightcolor=BRIGHT)

    def _build_ui(self):
        ttk.Label(self.root, text="KIRA'S PAWN  //  WRITING MODE",
                  style="Banner.TLabel").pack(anchor=W, padx=16, pady=(12, 3))
        ttk.Label(self.root, text="[ WORD ]  [ RESEARCH ]  [ DICTIONARY ]  [ OFFLINE GRAMMAR ]",
                  style="Muted.TLabel").pack(anchor=W, padx=16, pady=(0, 10))
        row = ttk.Frame(self.root)
        row.pack(fill=X, padx=16, pady=(0, 8))
        ttk.Button(row, text="OPEN WORD DRAFT", command=self.open_word).pack(side=LEFT)
        if PROPOSAL_DOC.is_file():
            ttk.Button(row, text="OPEN PROPOSAL", command=lambda: os.startfile(str(PROPOSAL_DOC))).pack(side=LEFT, padx=8)
        ttk.Label(row, textvariable=self.doc_label, style="Muted.TLabel").pack(side=LEFT, padx=8)
        tabs = ttk.Notebook(self.root)
        tabs.pack(fill=BOTH, expand=True, padx=16, pady=(0, 8))
        research = ttk.Frame(tabs, padding=10)
        dictionary = ttk.Frame(tabs, padding=10)
        grammar = ttk.Frame(tabs, padding=10)
        tabs.add(research, text="01  RESEARCH")
        tabs.add(dictionary, text="02  DICTIONARY")
        tabs.add(grammar, text="03  LOCAL GRAMMAR")
        self._research_tab(research)
        self._dictionary_tab(dictionary)
        self._grammar_tab(grammar)
        ttk.Label(self.root, textvariable=self.status, style="Muted.TLabel").pack(anchor=W, padx=16, pady=(0, 8))

    def _research_tab(self, frame):
        row = ttk.Frame(frame)
        row.pack(fill=X)
        ttk.Label(row, text="Keywords").pack(side=LEFT)
        self.query = ttk.Entry(row)
        self.query.pack(side=LEFT, fill=X, expand=True, padx=8)
        self.query.insert(0, "electric vehicle charging scheduling building power limit")
        self.query.bind("<Return>", lambda event: self.search())
        ttk.Button(row, text="Search", command=self.search).pack(side=LEFT)
        links = ttk.Frame(frame)
        links.pack(fill=X, pady=8)
        ttk.Button(links, text="Google Scholar", command=self.open_scholar).pack(side=LEFT)
        ttk.Button(links, text="Search alphaXiv", command=self.open_alphaxiv_search).pack(side=LEFT, padx=6)
        ttk.Button(links, text="Open selected source", command=self.open_selected).pack(side=LEFT)
        ttk.Button(links, text="Open selected on alphaXiv", command=self.open_selected_alphaxiv).pack(side=LEFT, padx=6)
        ttk.Button(links, text="Save results as CSV", command=self.save_results).pack(side=RIGHT)
        ttk.Label(frame, text="Results from Crossref, arXiv, and GitHub. Open each source before citing it.").pack(anchor=W)
        holder = ttk.Frame(frame)
        holder.pack(fill=BOTH, expand=True, pady=(7, 7))
        self.tree = ttk.Treeview(holder, columns=("source", "year", "title"), show="headings", selectmode="browse")
        for key, label, width in (("source", "Source", 95), ("year", "Year", 60), ("title", "Title", 740)):
            self.tree.heading(key, text=label)
            self.tree.column(key, width=width, anchor=W)
        scroll = ttk.Scrollbar(holder, orient=VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scroll.set)
        self.tree.pack(side=LEFT, fill=BOTH, expand=True)
        scroll.pack(side=RIGHT, fill=Y)
        self.tree.bind("<<TreeviewSelect>>", self.show_result)
        self.tree.bind("<Double-1>", lambda event: self.open_selected())
        self.details = ScrolledText(frame, height=7, wrap="word", state="disabled", font=(FONT, 10))
        self._style_text(self.details)
        self.details.pack(fill=X)

    def _dictionary_tab(self, frame):
        row = ttk.Frame(frame)
        row.pack(fill=X)
        ttk.Label(row, text="English word").pack(side=LEFT)
        self.word = ttk.Entry(row)
        self.word.pack(side=LEFT, fill=X, expand=True, padx=8)
        self.word.bind("<Return>", lambda event: self.lookup())
        ttk.Button(row, text="Look up", command=self.lookup).pack(side=LEFT)
        ttk.Button(row, text="Use clipboard", command=self.word_from_clipboard).pack(side=LEFT, padx=6)
        ttk.Button(row, text="Cambridge Dictionary", command=self.open_cambridge).pack(side=LEFT)
        ttk.Label(frame, text="Definitions and related words need internet. The grammar checker does not.").pack(anchor=W, pady=8)
        self.definition = ScrolledText(frame, wrap="word", font=(FONT, 11))
        self._style_text(self.definition)
        self.definition.pack(fill=BOTH, expand=True)

    def _grammar_tab(self, frame):
        ttk.Label(frame, textvariable=self.grammar_status).pack(anchor=W)
        row = ttk.Frame(frame)
        row.pack(fill=X, pady=8)
        ttk.Button(row, text="Paste from Word / clipboard", command=self.paste_grammar).pack(side=LEFT)
        ttk.Button(row, text="Check now", command=self.check_grammar).pack(side=LEFT, padx=6)
        ttk.Button(row, text="Apply selected suggestion", command=self.apply_suggestion).pack(side=LEFT)
        ttk.Button(row, text="Copy corrected text", command=self.copy_grammar).pack(side=LEFT, padx=6)
        ttk.Checkbutton(row, text="Check as I type", variable=self.auto_check).pack(side=RIGHT)
        self.grammar_text = ScrolledText(frame, wrap="word", undo=True, font=(FONT, 11), height=13)
        self._style_text(self.grammar_text)
        self.grammar_text.pack(fill=BOTH, expand=True)
        self.grammar_text.bind("<KeyRelease>", self.schedule_grammar)
        ttk.Label(frame, text="Suggestions appear below. Your text is sent only to 127.0.0.1 on this PC.").pack(anchor=W, pady=(8, 2))
        self.suggestions = ttk.Treeview(frame, columns=("issue", "suggestion"), show="headings", height=7)
        self.suggestions.heading("issue", text="Issue")
        self.suggestions.heading("suggestion", text="Suggested change")
        self.suggestions.column("issue", width=590)
        self.suggestions.column("suggestion", width=330)
        self.suggestions.pack(fill=X)

    def open_word(self):
        try:
            if self.source_doc is None:
                selected = filedialog.askopenfilename(
                    parent=self.root, title="Select a Word template or document",
                    filetypes=[("Word document", "*.docx")],
                )
                if not selected:
                    self.status.set("Select a Word document when you are ready to write.")
                    return
                self.source_doc = Path(selected)
                self.work_doc = self.source_doc.with_name(self.source_doc.stem + " Working Draft.docx")
                self.doc_label.set(str(self.work_doc))
                SETTINGS_FILE.write_text(json.dumps({"source_doc": str(self.source_doc)}, indent=2), encoding="utf-8")
            if not self.work_doc.exists():
                shutil.copy2(self.source_doc, self.work_doc)
            script = APP_DIR / "open_word.ps1"
            result = subprocess.run(
                ["powershell.exe", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", str(script), str(self.work_doc)],
                capture_output=True, text=True, timeout=30,
                creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
            )
            if result.returncode:
                os.startfile(str(self.work_doc))
            self.status.set("Working draft opened in Word. The original template is untouched.")
        except Exception as exc:
            messagebox.showerror("Could not open Word draft", str(exc))

    def _thread(self, work, done):
        def run():
            try:
                result = work()
                self.root.after(0, lambda: done(result, None))
            except Exception as exc:
                self.root.after(0, lambda exc=exc: done(None, exc))
        threading.Thread(target=run, daemon=True).start()

    def search(self):
        query = self.query.get().strip()
        if not query:
            return
        self.status.set("Searching Crossref, arXiv, and GitHub…")
        def work():
            results, errors = [], []
            for name, fn in (("Crossref", crossref_search), ("arXiv", arxiv_search), ("GitHub", github_search)):
                try:
                    results.extend(fn(query))
                except Exception as exc:
                    errors.append(f"{name}: {exc}")
            return results, errors
        self._thread(work, self.search_done)

    def search_done(self, outcome, error):
        if error:
            self.status.set(f"Search failed: {error}")
            return
        self.results, errors = outcome
        self.tree.delete(*self.tree.get_children())
        for i, item in enumerate(self.results):
            self.tree.insert("", END, iid=str(i), values=(item["source"], item["year"], item["title"]))
        self.status.set(f"Found {len(self.results)} results." + (" Unavailable: " + "; ".join(errors) if errors else ""))
        if self.results:
            self.tree.selection_set("0")
            self.show_result()

    def selected_result(self):
        selection = self.tree.selection()
        return self.results[int(selection[0])] if selection else None

    def show_result(self, event=None):
        item = self.selected_result()
        if not item:
            return
        content = f"{item['title']}\n{item['authors']} ({item['year']})\n{item['url']}\n"
        if item["doi"]:
            content += f"DOI: {item['doi']}\n"
        content += "\n" + (item["details"] or "No abstract or description supplied by this source.")
        self.details.configure(state="normal")
        self.details.delete("1.0", END)
        self.details.insert(END, content)
        self.details.configure(state="disabled")

    def open_selected(self):
        item = self.selected_result()
        if item and item["url"]:
            webbrowser.open(item["url"])

    def open_selected_alphaxiv(self):
        item = self.selected_result()
        if item and item["alphaxiv"]:
            webbrowser.open(item["alphaxiv"])
        else:
            self.status.set("alphaXiv link is available for arXiv results.")

    def open_scholar(self):
        query = self.query.get().strip()
        if query:
            webbrowser.open("https://scholar.google.com/scholar?" + urllib.parse.urlencode({"q": query}))

    def open_alphaxiv_search(self):
        query = self.query.get().strip()
        if query:
            webbrowser.open("https://www.google.com/search?" + urllib.parse.urlencode({"q": "site:alphaxiv.org/abs/ " + query}))

    def save_results(self):
        if not self.results:
            self.status.set("Search first, then save the results.")
            return
        RESULTS_DIR.mkdir(exist_ok=True)
        path = RESULTS_DIR / f"research_{datetime.now():%Y-%m-%d_%H%M%S}.csv"
        with path.open("w", encoding="utf-8-sig", newline="") as stream:
            writer = csv.DictWriter(stream, fieldnames=("source", "year", "title", "authors", "url", "doi", "details", "alphaxiv"))
            writer.writeheader()
            writer.writerows(self.results)
        self.status.set(f"Saved {len(self.results)} results to {path}")
        os.startfile(str(RESULTS_DIR))

    def word_from_clipboard(self):
        try:
            self.word.delete(0, END)
            self.word.insert(0, self.root.clipboard_get().strip())
        except Exception:
            self.status.set("Clipboard has no text.")

    def lookup(self):
        word = self.word.get().strip()
        if not word:
            return
        self.definition.delete("1.0", END)
        self.definition.insert(END, "Looking up word…")
        self._thread(lambda: word_lookup(word), self.lookup_done)

    def lookup_done(self, result, error):
        self.definition.delete("1.0", END)
        self.definition.insert(END, str(error) if error else result)

    def open_cambridge(self):
        word = self.word.get().strip()
        if word:
            webbrowser.open("https://dictionary.cambridge.org/dictionary/english/" + urllib.parse.quote(word))

    def start_languagetool(self):
        try:
            grammar_check("A simple sentence.")
            self.lt_ready = True
            self.root.after(0, lambda: self.grammar_status.set("Offline LanguageTool ready."))
            return
        except Exception:
            pass
        jars = list(VENDOR.rglob("languagetool-server.jar"))
        if not jars:
            self.root.after(0, lambda: self.grammar_status.set("LanguageTool is not installed; Word proofing is still available."))
            return
        jar = jars[0]
        try:
            java_bin = next((str(path) for path in Path("C:/Program Files/Java").glob("jdk-*/bin/java.exe")
                             if path.is_file()), "java")
            self.lt_log = (Path(__file__).resolve().parent / "languagetool.log").open("a", encoding="utf-8")
            self.lt_process = subprocess.Popen(
                [java_bin, "-cp", jar.name, "org.languagetool.server.HTTPServer", "--port", "8081"],
                cwd=jar.parent, stdout=self.lt_log, stderr=subprocess.STDOUT,
                creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
            )
            for _ in range(45):
                if self.lt_process.poll() is not None:
                    break
                try:
                    grammar_check("A simple sentence.")
                    self.lt_ready = True
                    self.root.after(0, lambda: self.grammar_status.set("Offline LanguageTool ready."))
                    return
                except Exception:
                    time.sleep(1)
            self.root.after(0, lambda: self.grammar_status.set("LanguageTool did not start. See writing_mode/languagetool.log."))
        except Exception as exc:
            self.root.after(0, lambda exc=exc: self.grammar_status.set(f"LanguageTool could not start: {exc}"))

    def paste_grammar(self):
        try:
            self.grammar_text.delete("1.0", END)
            self.grammar_text.insert("1.0", self.root.clipboard_get())
            self.check_grammar()
        except Exception:
            self.grammar_status.set("Clipboard has no text.")

    def schedule_grammar(self, event=None):
        if self.grammar_timer is not None:
            self.root.after_cancel(self.grammar_timer)
            self.grammar_timer = None
        if self.auto_check.get() and self.lt_ready:
            self.grammar_timer = self.root.after(1100, self.check_grammar)

    def check_grammar(self):
        content = self.grammar_text.get("1.0", "end-1c")
        if not content.strip():
            return
        if not self.lt_ready:
            self.grammar_status.set("Waiting for local LanguageTool to start…")
            return
        self.grammar_generation += 1
        generation = self.grammar_generation
        self.grammar_status.set("Checking on this computer…")
        self._thread(lambda: grammar_check(content), lambda result, error: self.grammar_done(content, generation, result, error))

    def grammar_done(self, content, generation, result, error):
        if generation != self.grammar_generation or content != self.grammar_text.get("1.0", "end-1c"):
            return
        if error:
            self.grammar_status.set(f"Local check failed: {error}")
            return
        self.checked_text = content
        self.matches = result
        self.suggestions.delete(*self.suggestions.get_children())
        for i, match in enumerate(result):
            choices = match.get("replacements", [])
            suggestion = choices[0].get("value", "") if choices else ""
            self.suggestions.insert("", END, iid=str(i), values=(match.get("message", ""), suggestion))
        self.grammar_status.set(f"Offline check found {len(result)} possible issue(s). Review suggestions before applying.")

    def apply_suggestion(self):
        selection = self.suggestions.selection()
        if not selection or self.grammar_text.get("1.0", "end-1c") != self.checked_text:
            self.grammar_status.set("Select a suggestion; check again if the text has changed.")
            return
        match = self.matches[int(selection[0])]
        choices = match.get("replacements", [])
        if not choices:
            return
        start = int(match["offset"])
        end = start + int(match["length"])
        updated = self.checked_text[:start] + choices[0]["value"] + self.checked_text[end:]
        self.grammar_text.delete("1.0", END)
        self.grammar_text.insert("1.0", updated)
        self.check_grammar()

    def copy_grammar(self):
        self.root.clipboard_clear()
        self.root.clipboard_append(self.grammar_text.get("1.0", "end-1c"))
        self.grammar_status.set("Text copied. Paste it into Word.")

    def close(self):
        if self.lt_process and self.lt_process.poll() is None:
            self.lt_process.terminate()
        if self.lt_log:
            self.lt_log.close()
        self.root.destroy()


if __name__ == "__main__":
    app = WritingMode()
    app.root.mainloop()
