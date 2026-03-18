# 🎯 AI Slide Builder

A multi-agent AI system that generates professional PowerPoint presentations from documents and templates. Built with Streamlit for an interactive demo experience.

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
| **Slide Mapper** | Maps drafted content to template layouts | Draft content, template analysis | Slide plan with source-slide assignments and text replacements |
| **Slide Generator** | Orchestrates PPTX manipulation | Template, draft, slide plan | Final `.pptx` file |

### Utilities

| Utility | Purpose |
|---------|---------|
| `document_parser.py` | Extracts text from PDF, TXT, JSON files |
| `template_analyzer.py` | Analyzes PPTX templates: structure, text inventory, metadata |
| `pptx_builder.py` | Low-level PPTX manipulation: duplicate, reorder, replace text |

## Pipeline Flow

```
1. UPLOAD & PARSE
   ├── Parse document (PDF/TXT/JSON) → plain text
   └── Analyze template → structure, text inventory, slide count

2. DRAFT CONTENT (AI Agent)
   ├── Send document + template summary to Claude API
   ├── Generate structured slide content (JSON)
   └── Return: titles, body, bullets, visual suggestions, speaker notes

3. REVIEW & EDIT (Human-in-the-loop)
   ├── Display draft in editable UI
   ├── User modifies content per-slide
   └── Optional: AI refinement based on user feedback

4. GENERATE SLIDES (AI Agent + PPTX Builder)
   ├── Map content → template layouts (AI decides which template slide per content)
   ├── Unpack template PPTX
   ├── Duplicate slides as needed
   ├── Replace text content (preserving formatting)
   ├── Reorder slides to match plan
   ├── Clean orphaned files
   ├── Pack into final PPTX
   └── Validate output

5. DOWNLOAD
   └── Present final PPTX for download
```

## Setup

```bash
# 1. Install dependencies
pip install -r requirements.txt

# 2. Set API key (or enter in the app sidebar)
export ANTHROPIC_API_KEY="sk-ant-..."

# 3. Run the app
streamlit run app.py
```

## Usage

1. **Upload** a source document (`.pdf`, `.txt`, or `.json`) and a PPTX template
2. Set the number of slides and any special instructions
3. Click **Parse & Analyze** to process both files
4. Click **Generate Draft Content** — the AI analyzes your document and creates slide content
5. **Review and edit** each slide's title, body, type, and speaker notes
6. Optionally use **AI Refinement** to adjust the draft with natural language feedback
7. Click **Approve & Generate Slides** to build the final presentation
8. **Download** your finished `.pptx` file

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
│   ├── __init__.py
│   ├── document_parser.py          # PDF/TXT/JSON parser
│   ├── template_analyzer.py        # PPTX template analysis
│   └── pptx_builder.py             # Low-level PPTX manipulation
└── sample_data/
    └── sample_report.json          # Example document for testing
```

## Supported Formats

| Input Type | Extensions | Notes |
|------------|-----------|-------|
| Documents | `.pdf`, `.txt`, `.json` | PDF uses pdfplumber with pypdf fallback |
| Templates | `.pptx` | Any PowerPoint template with text placeholders |

## Technical Notes

- Uses Claude (claude-sonnet-4-20250514) for content generation and template mapping
- PPTX manipulation uses XML-level editing for maximum fidelity to template formatting
- Template analysis extracts all text shapes with their names for precise replacement
- Slide duplication preserves all formatting, images, and layout properties
- Validation checks for leftover placeholder text after generation
