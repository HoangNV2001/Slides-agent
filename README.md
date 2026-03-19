# 🎯 AI Slide Builder

A multi-agent AI system that generates professional PowerPoint presentations from documents and templates.

**Pure Python — runs on macOS, Linux, Windows. No external scripts or system dependencies.**

## Setup

```bash
pip install -r requirements.txt
export OPENAI_API_KEY="sk-proj-..."     # or enter in the app sidebar
export OPENAI_MODEL="gpt-5.4"           # optional, defaults to GPT-5.4
streamlit run app.py
```

## Architecture

```
┌─────────────────────────────────────────────────────┐
│                  Streamlit App (UI)                  │
│  Upload → Draft → Review/Edit → Generate → Download │
└─────────┬───────────────┬───────────────┬───────────┘
          │               │               │
    ┌─────▼─────┐  ┌──────▼──────┐  ┌─────▼──────┐
    │  Content   │  │   Slide     │  │   Slide    │
    │  Drafter   │  │   Mapper    │  │  Generator │
    │  Agent     │  │   Agent     │  │   Agent    │
    └─────┬─────┘  └──────┬──────┘  └─────┬──────┘
          │               │               │
    ┌─────▼─────┐  ┌──────▼──────┐  ┌─────▼──────┐
    │ Document   │  │  Template   │  │   PPTX     │
    │ Parser     │  │  Analyzer   │  │  Builder   │
    └───────────┘  └─────────────┘  └────────────┘
```

### Agents

| Agent | Role | Input | Output |
|-------|------|-------|--------|
| **Content Drafter** | Analyzes document + template → drafts slide content | Document text, template summary, user instructions | Structured JSON with slide titles, body, notes, visual suggestions |
| **Slide Mapper** | Maps drafted content to template layouts | Draft content, template analysis | Slide plan with source-slide-index assignments and text replacements |
| **Slide Generator** | Orchestrates PPTX manipulation | Template, draft, slide plan | Final `.pptx` file |

### Utilities

| Utility | Purpose |
|---------|---------|
| `document_parser.py` | Extracts text from PDF, TXT, JSON files |
| `template_analyzer.py` | Analyzes PPTX templates using python-pptx: structure, text inventory, layouts |
| `pptx_builder.py` | Slide duplication, reordering, deletion, text replacement via python-pptx + lxml |

## Pipeline Flow

```
1. UPLOAD & PARSE
   ├── Parse document (PDF/TXT/JSON) → plain text
   └── Analyze template → structure, text inventory, layouts

2. DRAFT CONTENT (AI Agent)
   ├── Send document + template summary to OpenAI API
   └── Generate structured slide content (JSON): titles, body, bullets, notes

3. REVIEW & EDIT (Human-in-the-loop)
   ├── Display draft in editable UI
   ├── User modifies content per-slide
   └── Optional: AI refinement based on user feedback

4. GENERATE SLIDES (AI Agent + PPTX Builder)
   ├── Map content → template layouts (AI decides which template slide per content)
   ├── Duplicate slides as needed (python-pptx + lxml deep copy)
   ├── Replace text content (preserving formatting)
   ├── Reorder slides to match plan, remove unused
   ├── Save final PPTX
   └── Validate output

5. DOWNLOAD
   └── Present final PPTX for download
```

## File Structure

```
slide_agent/
├── app.py                          # Streamlit main app
├── requirements.txt
├── README.md
├── agents/
│   ├── __init__.py
│   ├── content_drafter.py          # AI content generation agent
│   ├── slide_mapper.py             # AI template mapping agent
│   └── slide_generator.py          # PPTX generation orchestrator
├── utils/
│   ├── __init__.py                 # Exports all utilities
│   ├── document_parser.py          # PDF/TXT/JSON parser
│   ├── template_analyzer.py        # PPTX template analysis (python-pptx)
│   └── pptx_builder.py             # Slide manipulation (python-pptx + lxml)
└── sample_data/
    └── sample_report.json          # Example document for testing
```

## Supported Formats

| Input Type | Extensions | Notes |
|------------|-----------|-------|
| Documents | `.pdf`, `.txt`, `.json` | PDF uses pdfplumber with pypdf fallback |
| Templates | `.pptx` | Any PowerPoint template with text placeholders |
