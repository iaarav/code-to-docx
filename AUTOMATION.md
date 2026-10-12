# GitHub automation

The web page creates a document after a folder or ZIP is selected. For unattended use in GitHub Actions, `generate_docx.py` creates the same kind of document from a repository checkout. It includes readable text and code files, attaches screenshots found beside their code, and leaves an output placeholder when no screenshot is available.

The `semester01_C` repository's `.github/workflows/code-to-docx.yml` checks out this repository, runs the generator over `semester 01`, and commits `code-export.docx` to the repository root when its contents change. The workflow needs the repository's Actions token to have `contents: write` permission. GitHub suppresses another push-triggered run for commits made with `GITHUB_TOKEN`.

To run the generator locally:

```sh
python -m pip install -r requirements.txt
python generate_docx.py --root "/path/to/repository" --output "/path/to/repository/code-export.docx"
```
