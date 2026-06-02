# OOXML DOCX Skill

An open-source skill for generating `.docx` documents through OOXML with enhanced WPS compatibility.

## Why This Project

Many existing DOCX generation libraries work well in Microsoft Word but encounter compatibility issues in WPS Office, especially when handling:

* Tracked Changes (修订)
* Comments (批注)
* Complex Tables
* Headers and Footers
* Legal Document Formatting

This project aims to provide a reusable skill layer that enables AI agents and automation workflows to generate high-quality DOCX files that can be opened and edited reliably in both Microsoft Word and WPS Office.

## Features

* Generate OOXML-based `.docx` documents
* Paragraphs, headings, lists, and tables
* Images, headers, and footers
* Comments support
* Tracked revisions support
* WPS compatibility optimization
* AI Agent / Skill integration friendly

## Use Cases

### Legal Document Automation

* Contract review reports
* Risk warning reports
* Legal opinions
* Due diligence reports

### AI Agent Workflows

* Automated report generation
* Structured document output
* Enterprise document automation

## Roadmap

* [ ] Core OOXML document writer
* [ ] Comment insertion
* [ ] Tracked changes generation
* [ ] WPS compatibility test suite
* [ ] Agent Skill integration examples
* [ ] Legal document templates

## License

MIT License
