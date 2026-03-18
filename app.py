"""
AI Slide Builder - Streamlit Demo App
A multi-agent system for generating presentations from documents + templates.
"""
import json
import os
import shutil
import sys
import tempfile
import time

import streamlit as st

# Add project to path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from utils.document_parser import parse_document
from utils.template_analyzer import analyze_template, get_template_summary
from agents.content_drafter import draft_slide_content, refine_draft
from agents.slide_mapper import map_content_to_template
from agents.slide_generator import generate_slides

# ─── Page Config ───
st.set_page_config(
    page_title="AI Slide Builder",
    page_icon="🎯",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─── Custom Styling ───
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;700&family=JetBrains+Mono:wght@400;500&display=swap');

:root {
    --primary: #6C5CE7;
    --primary-light: #A29BFE;
    --accent: #00CEC9;
    --bg-dark: #0F0F1A;
    --bg-card: #1A1A2E;
    --text: #E4E4F0;
    --text-muted: #8888AA;
    --success: #00B894;
    --warning: #FDCB6E;
    --danger: #E17055;
}

.stApp {
    font-family: 'DM Sans', sans-serif;
}

/* Header */
.main-header {
    background: linear-gradient(135deg, #6C5CE7 0%, #0984E3 50%, #00CEC9 100%);
    -webkit-background-clip: text;
    -webkit-text-fill-color: transparent;
    font-size: 2.5rem;
    font-weight: 700;
    margin-bottom: 0;
    letter-spacing: -0.02em;
}

.sub-header {
    color: #8888AA;
    font-size: 1.05rem;
    margin-top: -0.5rem;
    margin-bottom: 2rem;
}

/* Pipeline Steps */
.pipeline-step {
    padding: 1rem 1.2rem;
    border-radius: 12px;
    border: 1px solid rgba(108, 92, 231, 0.2);
    margin-bottom: 0.7rem;
    transition: all 0.3s ease;
}

.pipeline-step.active {
    border-color: #6C5CE7;
    background: rgba(108, 92, 231, 0.08);
    box-shadow: 0 0 20px rgba(108, 92, 231, 0.1);
}

.pipeline-step.done {
    border-color: #00B894;
    background: rgba(0, 184, 148, 0.06);
}

.step-number {
    display: inline-block;
    width: 28px;
    height: 28px;
    border-radius: 50%;
    text-align: center;
    line-height: 28px;
    font-weight: 700;
    font-size: 0.85rem;
    margin-right: 0.6rem;
}

.step-active .step-number { background: #6C5CE7; color: white; }
.step-done .step-number { background: #00B894; color: white; }
.step-pending .step-number { background: #2A2A3E; color: #8888AA; }

/* Slide Card */
.slide-card {
    background: rgba(26, 26, 46, 0.6);
    border: 1px solid rgba(108, 92, 231, 0.15);
    border-radius: 12px;
    padding: 1.2rem;
    margin-bottom: 1rem;
}

.slide-card h4 {
    color: #A29BFE;
    margin-bottom: 0.5rem;
}

/* Status Badge */
.status-badge {
    display: inline-block;
    padding: 0.2rem 0.7rem;
    border-radius: 20px;
    font-size: 0.8rem;
    font-weight: 500;
}

.status-success { background: rgba(0,184,148,0.15); color: #00B894; }
.status-warning { background: rgba(253,203,110,0.15); color: #FDCB6E; }
.status-error { background: rgba(225,112,85,0.15); color: #E17055; }
.status-info { background: rgba(108,92,231,0.15); color: #A29BFE; }

/* Log */
.log-entry {
    font-family: 'JetBrains Mono', monospace;
    font-size: 0.82rem;
    padding: 0.25rem 0;
    color: #8888AA;
}

.log-entry.success { color: #00B894; }
.log-entry.warning { color: #FDCB6E; }
.log-entry.error { color: #E17055; }
</style>
""", unsafe_allow_html=True)


# ─── Session State Initialization ───
def init_session():
    defaults = {
        "current_step": 0,  # 0=upload, 1=draft, 2=review, 3=generate, 4=done
        "document_text": None,
        "template_analysis": None,
        "template_summary": None,
        "draft_content": None,
        "slide_plan": None,
        "generation_result": None,
        "output_path": None,
        "api_key": None,
        "generation_log": [],
        "template_path_saved": None,
        "doc_filename": None,
        "template_filename": None,
    }
    for key, val in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = val


init_session()


# ─── Sidebar ───
with st.sidebar:
    st.markdown("### ⚙️ Configuration")

    api_key = st.text_input(
        "Anthropic API Key",
        type="password",
        value=st.session_state.api_key or "",
        help="Required for AI content generation",
    )
    if api_key:
        st.session_state.api_key = api_key

    st.divider()

    # Pipeline progress
    st.markdown("### 📋 Pipeline")
    steps = [
        ("Upload & Parse", "Upload document + template"),
        ("Draft Content", "AI drafts slide content"),
        ("Review & Edit", "Human review & refinement"),
        ("Generate Slides", "Build the PPTX file"),
        ("Download", "Get your presentation"),
    ]

    for i, (name, desc) in enumerate(steps):
        if i < st.session_state.current_step:
            st.markdown(f"✅ **{name}**")
        elif i == st.session_state.current_step:
            st.markdown(f"🔵 **{name}** ← _current_")
        else:
            st.markdown(f"⬜ {name}")

    st.divider()

    if st.button("🔄 Reset Pipeline", use_container_width=True):
        for key in list(st.session_state.keys()):
            if key != "api_key":
                del st.session_state[key]
        init_session()
        st.rerun()


# ─── Header ───
st.markdown('<h1 class="main-header">AI Slide Builder</h1>', unsafe_allow_html=True)
st.markdown(
    '<p class="sub-header">Upload a document & template → AI drafts content → Review → Generate professional slides</p>',
    unsafe_allow_html=True,
)


# ═══════════════════════════════════════════════
# STEP 0: Upload & Parse
# ═══════════════════════════════════════════════
if st.session_state.current_step == 0:
    st.markdown("## 📤 Step 1: Upload Files")

    col1, col2 = st.columns(2)

    with col1:
        st.markdown("### Source Document")
        doc_file = st.file_uploader(
            "Upload your document",
            type=["pdf", "txt", "json"],
            help="The content source for your presentation",
        )

    with col2:
        st.markdown("### Slide Template")
        template_file = st.file_uploader(
            "Upload PPTX template",
            type=["pptx"],
            help="The PowerPoint template to use as the visual base",
        )

    st.divider()

    num_slides = st.slider("Number of slides to generate", min_value=3, max_value=30, value=8)
    user_instructions = st.text_area(
        "Additional instructions (optional)",
        placeholder="e.g., Focus on the financial data, make it suitable for C-level audience, include a Q&A slide...",
        height=80,
    )

    st.session_state["num_slides"] = num_slides
    st.session_state["user_instructions"] = user_instructions

    if doc_file and template_file:
        if st.button("🚀 Parse & Analyze", type="primary", use_container_width=True):
            if not st.session_state.api_key:
                st.error("Please enter your Anthropic API Key in the sidebar.")
                st.stop()

            with st.spinner("Parsing document..."):
                # Save uploads
                tmp_dir = tempfile.mkdtemp()
                doc_path = os.path.join(tmp_dir, doc_file.name)
                with open(doc_path, "wb") as f:
                    f.write(doc_file.getvalue())

                tpl_path = os.path.join(tmp_dir, template_file.name)
                with open(tpl_path, "wb") as f:
                    f.write(template_file.getvalue())

                # Parse document
                doc_text = parse_document(doc_path)
                st.session_state.document_text = doc_text
                st.session_state.doc_filename = doc_file.name

            with st.spinner("Analyzing template structure..."):
                analysis = analyze_template(tpl_path)
                summary = get_template_summary(analysis)
                st.session_state.template_analysis = analysis
                st.session_state.template_summary = summary
                st.session_state.template_path_saved = tpl_path
                st.session_state.template_filename = template_file.name

            st.session_state.current_step = 1
            st.rerun()
    else:
        st.info("👆 Upload both a document and a PPTX template to continue.")


# ═══════════════════════════════════════════════
# STEP 1: Draft Content
# ═══════════════════════════════════════════════
elif st.session_state.current_step == 1:
    st.markdown("## ✍️ Step 2: AI Content Drafting")

    # Show parsed info
    with st.expander("📄 Parsed Document Preview", expanded=False):
        st.text(st.session_state.document_text[:3000] + ("..." if len(st.session_state.document_text) > 3000 else ""))

    with st.expander("🗂️ Template Structure", expanded=False):
        st.text(st.session_state.template_summary)

    st.divider()

    if st.button("🤖 Generate Draft Content", type="primary", use_container_width=True):
        with st.spinner("AI is drafting your slide content... This may take 15-30 seconds."):
            draft = draft_slide_content(
                document_text=st.session_state.document_text,
                template_summary=st.session_state.template_summary,
                num_slides=st.session_state.get("num_slides", 8),
                user_instructions=st.session_state.get("user_instructions", ""),
                api_key=st.session_state.api_key,
            )
            st.session_state.draft_content = draft

        if draft.get("error"):
            st.error(f"Drafting error: {draft['error']}")
        else:
            st.session_state.current_step = 2
            st.rerun()


# ═══════════════════════════════════════════════
# STEP 2: Review & Edit
# ═══════════════════════════════════════════════
elif st.session_state.current_step == 2:
    st.markdown("## 👀 Step 3: Review & Edit Draft")

    draft = st.session_state.draft_content

    if draft.get("outline"):
        st.info(f"**Strategy:** {draft['outline']}")

    slides = draft.get("slides", [])
    if not slides:
        st.warning("No slides generated. Please go back and try again.")
        if st.button("← Back to Draft"):
            st.session_state.current_step = 1
            st.rerun()
        st.stop()

    # Editable slides
    edited_slides = []
    for i, slide in enumerate(slides):
        with st.expander(
            f"Slide {slide.get('slide_number', i+1)}: {slide.get('title', 'Untitled')} "
            f"({slide.get('slide_type', 'content')})",
            expanded=(i < 3),
        ):
            col_a, col_b = st.columns([2, 1])

            with col_a:
                title = st.text_input(f"Title###{i}", value=slide.get("title", ""), key=f"title_{i}")
                subtitle = st.text_input(f"Subtitle###{i}", value=slide.get("subtitle", ""), key=f"sub_{i}")

                # Body / bullet points
                body = slide.get("body", "")
                bullets = slide.get("bullet_points", [])
                if bullets and not body:
                    body = "\n".join(f"• {b}" for b in bullets)
                elif bullets:
                    body = body + "\n" + "\n".join(f"• {b}" for b in bullets)

                body_edited = st.text_area(f"Body Content###{i}", value=body, height=150, key=f"body_{i}")

            with col_b:
                slide_type = st.selectbox(
                    f"Type###{i}",
                    ["title", "content", "comparison", "data", "quote", "section_divider", "closing"],
                    index=["title", "content", "comparison", "data", "quote", "section_divider", "closing"].index(
                        slide.get("slide_type", "content")
                    ) if slide.get("slide_type", "content") in ["title", "content", "comparison", "data", "quote", "section_divider", "closing"] else 1,
                    key=f"type_{i}",
                )
                visual = st.text_area(
                    f"Visual###{i}",
                    value=slide.get("visual_suggestion", "") or "",
                    height=80,
                    key=f"visual_{i}",
                )
                notes = st.text_area(
                    f"Speaker Notes###{i}",
                    value=slide.get("speaker_notes", "") or "",
                    height=80,
                    key=f"notes_{i}",
                )

            edited_slides.append({
                "slide_number": slide.get("slide_number", i + 1),
                "slide_type": slide_type,
                "title": title,
                "subtitle": subtitle,
                "body": body_edited,
                "bullet_points": [
                    line.lstrip("•-– ").strip()
                    for line in body_edited.split("\n")
                    if line.strip().startswith(("•", "-", "–"))
                ] or slide.get("bullet_points", []),
                "visual_suggestion": visual,
                "speaker_notes": notes,
                "template_slide_hint": slide.get("template_slide_hint", ""),
            })

    st.divider()

    # AI Refinement
    st.markdown("### 🔄 AI Refinement (Optional)")
    feedback = st.text_area(
        "Provide feedback to refine the draft",
        placeholder="e.g., Make slide 3 more concise, add more data points to slide 5, change the tone to be more casual...",
        height=80,
    )

    col_refine, col_proceed = st.columns(2)

    with col_refine:
        if st.button("🔄 Refine with AI", use_container_width=True, disabled=not feedback):
            if feedback:
                with st.spinner("AI is refining your draft..."):
                    updated_draft = {**draft, "slides": edited_slides}
                    refined = refine_draft(
                        current_draft=updated_draft,
                        user_feedback=feedback,
                        document_text=st.session_state.document_text,
                        api_key=st.session_state.api_key,
                    )
                    if not refined.get("error"):
                        st.session_state.draft_content = refined
                        st.rerun()
                    else:
                        st.error(f"Refinement error: {refined['error']}")

    with col_proceed:
        if st.button("✅ Approve & Generate Slides", type="primary", use_container_width=True):
            # Save edited slides
            st.session_state.draft_content = {**draft, "slides": edited_slides}
            st.session_state.current_step = 3
            st.rerun()


# ═══════════════════════════════════════════════
# STEP 3: Generate Slides
# ═══════════════════════════════════════════════
elif st.session_state.current_step == 3:
    st.markdown("## 🏗️ Step 4: Generating Slides")

    # Phase 1: Map content to template
    if not st.session_state.slide_plan:
        with st.spinner("🗺️ Mapping content to template layouts..."):
            plan = map_content_to_template(
                draft=st.session_state.draft_content,
                template_analysis=st.session_state.template_analysis,
                user_instructions=st.session_state.get("user_instructions", ""),
                api_key=st.session_state.api_key,
            )
            st.session_state.slide_plan = plan

        if plan.get("error"):
            st.error(f"Mapping error: {plan['error']}")
            if plan.get("raw_response"):
                with st.expander("Raw AI Response"):
                    st.code(plan["raw_response"])
            if st.button("← Back to Edit"):
                st.session_state.current_step = 2
                st.rerun()
            st.stop()

    # Show mapping plan
    plan = st.session_state.slide_plan
    with st.expander("🗺️ Slide Mapping Plan", expanded=False):
        if plan.get("strategy_notes"):
            st.info(plan["strategy_notes"])
        for item in plan.get("slide_plan", []):
            st.markdown(
                f"**Slide {item.get('draft_slide_number')}** → "
                f"`{item.get('source_template_slide')}` "
                f"({item.get('layout_reason', '')})"
            )

    # Phase 2: Generate PPTX
    st.divider()

    if not st.session_state.generation_result:
        log_placeholder = st.empty()
        progress_bar = st.progress(0, text="Starting generation...")

        with st.spinner("🔧 Building your presentation..."):
            work_dir = tempfile.mkdtemp()
            output_path = os.path.join(work_dir, "output_presentation.pptx")

            result = generate_slides(
                template_path=st.session_state.template_path_saved,
                draft=st.session_state.draft_content,
                slide_plan=st.session_state.slide_plan,
                output_path=output_path,
                work_dir=work_dir,
            )

            st.session_state.generation_result = result
            st.session_state.output_path = output_path if result["status"] == "success" else None

        # Show logs
        steps = result.get("steps", [])
        for i, step in enumerate(steps):
            progress_bar.progress((i + 1) / max(len(steps), 1), text=step)

        st.rerun()

    else:
        result = st.session_state.generation_result

        # Show generation log
        st.markdown("### 📋 Generation Log")
        for step in result.get("steps", []):
            if "✓" in step:
                st.markdown(f'<div class="log-entry success">{step}</div>', unsafe_allow_html=True)
            elif "Error" in step or "error" in step:
                st.markdown(f'<div class="log-entry error">{step}</div>', unsafe_allow_html=True)
            else:
                st.markdown(f'<div class="log-entry">{step}</div>', unsafe_allow_html=True)

        # Warnings
        warnings = result.get("warnings", [])
        if warnings:
            with st.expander(f"⚠️ {len(warnings)} Warnings", expanded=False):
                for w in warnings:
                    st.warning(w)

        if result["status"] == "success":
            st.success("🎉 Presentation generated successfully!")
            st.session_state.current_step = 4
            st.rerun()
        else:
            st.error(f"Generation failed: {result.get('error', 'Unknown error')}")
            if result.get("traceback"):
                with st.expander("Traceback"):
                    st.code(result["traceback"])

            col_retry, col_back = st.columns(2)
            with col_retry:
                if st.button("🔄 Retry Generation"):
                    st.session_state.generation_result = None
                    st.session_state.slide_plan = None
                    st.rerun()
            with col_back:
                if st.button("← Back to Edit"):
                    st.session_state.generation_result = None
                    st.session_state.slide_plan = None
                    st.session_state.current_step = 2
                    st.rerun()


# ═══════════════════════════════════════════════
# STEP 4: Download
# ═══════════════════════════════════════════════
elif st.session_state.current_step == 4:
    st.markdown("## 🎉 Step 5: Your Presentation is Ready!")

    result = st.session_state.generation_result
    output_path = st.session_state.output_path

    if output_path and os.path.exists(output_path):
        col1, col2 = st.columns([2, 1])

        with col1:
            # Download button
            with open(output_path, "rb") as f:
                pptx_data = f.read()

            filename = f"ai_slides_{st.session_state.get('doc_filename', 'output').rsplit('.', 1)[0]}.pptx"

            st.download_button(
                label="⬇️ Download Presentation",
                data=pptx_data,
                file_name=filename,
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                type="primary",
                use_container_width=True,
            )

        with col2:
            file_size = len(pptx_data) / 1024
            st.metric("File Size", f"{file_size:.1f} KB")

        # Show summary
        st.divider()
        st.markdown("### 📊 Generation Summary")

        col_a, col_b, col_c = st.columns(3)
        with col_a:
            st.metric("Source", st.session_state.get("doc_filename", "Unknown"))
        with col_b:
            st.metric("Template", st.session_state.get("template_filename", "Unknown"))
        with col_c:
            num = len(st.session_state.draft_content.get("slides", []))
            st.metric("Slides Generated", num)

        # Show final content
        with st.expander("📝 Final Slide Content"):
            for slide in st.session_state.draft_content.get("slides", []):
                st.markdown(f"**Slide {slide.get('slide_number', '?')}: {slide.get('title', 'Untitled')}**")
                st.markdown(f"_{slide.get('slide_type', 'content')}_")
                body = slide.get("body", "")
                bullets = slide.get("bullet_points", [])
                if body:
                    st.text(body)
                if bullets:
                    for b in bullets:
                        st.markdown(f"  • {b}")
                st.divider()

        # Validation text
        if result and result.get("validation_text"):
            with st.expander("🔍 Output Validation (markitdown)"):
                st.text(result["validation_text"])

        # Warnings
        warnings = result.get("warnings", []) if result else []
        if warnings:
            with st.expander(f"⚠️ {len(warnings)} Warnings"):
                for w in warnings:
                    st.warning(w)

        st.divider()
        col_new, col_redo = st.columns(2)
        with col_new:
            if st.button("🆕 Start New Project", use_container_width=True):
                for key in list(st.session_state.keys()):
                    if key != "api_key":
                        del st.session_state[key]
                init_session()
                st.rerun()
        with col_redo:
            if st.button("✏️ Edit & Regenerate", use_container_width=True):
                st.session_state.generation_result = None
                st.session_state.slide_plan = None
                st.session_state.current_step = 2
                st.rerun()

    else:
        st.error("Output file not found. Please regenerate.")
        if st.button("← Back"):
            st.session_state.current_step = 3
            st.session_state.generation_result = None
            st.rerun()
