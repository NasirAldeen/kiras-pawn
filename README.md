# Kira's Pawn

Kira's Pawn is a terminal-themed writing companion for Windows. One desktop shortcut opens a working copy of a Word document and a window for research, word lookup, and offline grammar checks. It never edits the original document.

## Setup

Install Python with Tkinter, Java, and Microsoft Word. Clone this repository, then run:

```powershell
powershell -ExecutionPolicy Bypass -File .\setup.ps1
```

Setup downloads the official LanguageTool package (about 250 MB) into `vendor/` and creates a **Kira's Pawn** desktop shortcut. To set up the shortcut without LanguageTool, add `-SkipLanguageTool`; Word's own proofing still works. The downloaded package and your documents stay on your computer and are excluded from Git.

On first launch, choose a `.docx` document. Kira's Pawn makes a separate `Working Draft.docx` beside it and remembers your selection in a local settings file. If the senior project template is in the parent folder, the app finds it automatically. Word is set to mark spelling and grammar errors as you type.

## While writing

1. Write in the working draft opened in Word.
2. Enter keywords in **Research**. The app retrieves publication metadata from Crossref and arXiv, plus related GitHub repositories. Double-click a result to read the source. Google Scholar opens a normal browser search. The alphaXiv button searches indexed alphaXiv papers, and arXiv results have direct alphaXiv links.
3. Look up a word in **Dictionary** or copy a word in Word and use **Use clipboard**. Definitions come from Free Dictionary API; related words come from Datamuse.
4. For a second grammar check, copy a paragraph from Word and use **Paste from Word / clipboard** under **Local Grammar**. LanguageTool checks text on `127.0.0.1` only. You can also type in that panel and review suggestions as you type. Nothing is changed automatically; select a suggestion to apply it, then copy the corrected text back to Word.
5. **Save results as CSV** exports the current research list to the `Research Results` folder. Check the paper itself and its publication status before citing it.

Research and dictionary searches need internet. Word proofing and the LanguageTool panel work locally. Google Scholar and alphaXiv results open in your browser; the app does not scrape either service.

## Sources and service documentation

- [Crossref REST API](https://www.crossref.org/documentation/retrieve-metadata/rest-api/)
- [arXiv API manual](https://info.arxiv.org/help/api/user-manual.html)
- [GitHub REST search API](https://docs.github.com/en/rest/search/search)
- [Google Scholar search help](https://scholar.google.com/intl/en-us/scholar/help.html)
- [alphaXiv](https://www.alphaxiv.org/)
- [Free Dictionary API](https://dictionaryapi.dev/)
- [Datamuse API](https://www.datamuse.com/api/)
- [LanguageTool local server](https://dev.languagetool.org/http-server)
- [Word spelling and grammar settings](https://support.microsoft.com/en-us/office/check-spelling-and-grammar-in-office)

The app uses only Python's standard library and the local LanguageTool package. If the shortcut is missing, run `pyw .\writing_mode.py` from this folder.
