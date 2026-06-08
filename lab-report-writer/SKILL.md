---
name: lab-report-writer
description: "Draft and complete Chinese laboratory or experiment reports from a Word template, PDF manuals or guides, screenshots, and result photos. Use when Codex needs to fill a structured report with sections such as objectives, equipment, principles and procedure, reflections, steps, and conclusion, especially when the user provides a .docx template plus reference PDFs and images and expects a finished report instead of raw notes."
---

# Lab Report Writer

## Overview

Write a complete Chinese experiment report from user materials instead of producing detached notes.
Preserve the user's template, extract only the relevant facts from manuals, and turn them into concise academic report prose that matches student lab-report style.

## Workflow

### 1. Confirm deliverables and inputs

Collect these items first:

- report template path, usually a `.docx`
- reference manuals, experiment guides, or product handbooks, often `.pdf`
- result photos, screenshots, or on-site images
- experiment topic, such as motor, torque, or kinematics

Assume the user wants a finished document unless they explicitly ask for draft text only.

If the output must be a Word file, pair this skill with a DOCX editing workflow. If another DOCX-focused skill is available, use it. Otherwise, use Word COM automation on Windows for final document editing.

### 2. Build the report frame

Default to this section order unless the template clearly differs:

1. objectives
2. equipment or tools
3. principles and operating process
4. reflections
5. content and steps
6. conclusion and takeaways

When the template already contains headings and placeholder text, inspect its paragraph structure before editing.
On repeated reports that reuse the same template, preserve the known heading positions and replace only the body paragraphs.

### 3. Extract only the useful source facts

Read the reference materials with a narrow goal:

- experiment objectives
- hardware or software list
- wiring, communication, control, or modeling steps
- formulas, coordinate systems, DH parameters, forward or inverse solve logic, control modes, or sensor and load relationships
- safety notes and shutdown procedure
- exact specs that are explicitly present in the manual

Prefer local text extraction and search over manual browsing of the whole file.
Use `pdftotext` plus `rg` when available.
Do not bulk-copy the manual into the report.

Separate facts into two buckets:

- directly supported facts from the manual
- observed results or reasonable inferences from result images

Do not invent measurements, screenshots, software outputs, or parameter values that are not visible in the sources.

### 4. Draft each section with the right level of detail

Use this writing pattern.

#### Objectives

Write 3 to 4 numbered items.
State what the student needs to understand, operate, verify, or compare.
Combine theory, platform familiarity, and practical validation.

#### Equipment Or Tools

List the main hardware, controller, cables, software, manuals, and reference images.
Name actual platform or model numbers when they are visible in the source.

#### Principles And Operating Process

Split this into two logical parts inside the prose:

- principles: explain the mechanism, model, control relationship, or mathematical basis
- operating process: summarize the practical sequence from wiring or configuration to execution and shutdown

This section should connect theory to the actual platform instead of reading like copied textbook text.

#### Reflections

Write 2 to 4 reflective points.
Focus on modeling accuracy, parameter consistency, safety, verification logic, common failure points, or the gap between simulation and hardware.

#### Content And Steps

Write 7 to 10 chronological steps.
Make them operational and concrete.
Use actual software names, buttons, or files when the source provides them.

#### Conclusion And Takeaways

Start with a summary of what was completed and what was verified.
Then connect theory, software output, and observed results.
End with a short reflection on engineering practice, safety, validation, or later experiments.

When the user provides result images, place them in this section unless the template clearly requires another location.

### 5. Produce the final DOCX reliably

For Windows `.docx` template filling, prefer Word COM automation through Python when formatting preservation matters.
This is the default path for repeated lab-report work.

Preferred implementation pattern:

1. copy or save the template as a new output file
2. inspect paragraph indices and heading locations
3. clear placeholder ranges section by section
4. insert drafted paragraphs with consistent body formatting
5. insert result images with centered captions
6. save the new file without overwriting the original template

For repeated reports built from the same template, reuse the previous working script shape:

- constants for template, output, and image paths
- paragraph arrays for each section
- helper functions for formatting, paragraph clearing, and image insertion

Keep the generated script in a per-report working directory so it can be reused and adjusted for the next experiment.

### 6. Insert images deliberately

Use local image files that correspond to the user-provided thread images.
Map each image to a meaningful caption based on what is actually shown.

Typical caption style:

- `Figure 6.1 Physical experiment platform`
- `Figure 6.2 Monitoring or control interface`
- `Figure 6.3 Combined validation result`

Do not claim hidden details that the image does not show.

### 7. Validate before finishing

After generation, verify at least:

- the output file exists
- the title and main section headings are still correct
- section order is intact
- conclusion text exists after image insertion
- image captions appear in the intended position

When needed, reopen the generated Word file and inspect key paragraph indices.
If the template is complex, optionally export a PDF preview or inspect the end pages.

## Style Rules

- Write the final report in concise Chinese academic prose suitable for an undergraduate experiment report.
- Prefer practical, factual wording over exaggerated conclusions.
- Use numbered paragraphs when the template style implies them.
- Keep the tone as a student's report, not a product brochure.
- Cite exact model names and specs only when the source provides them.
- When translating source material into prose, summarize and reorganize; do not dump raw manual text.

## Decision Rules

- If the user provides a `.docx` template and wants a final report, produce the document directly instead of only returning draft text.
- If the source material is incomplete, write only what can be supported and keep wording conservative.
- If previous reports using the same template already succeeded, reuse their paragraph-index strategy and script pattern.
- If terminal encoding causes problems on Windows, prefer Python plus `win32com.client` for the final document edit path.

## Output Convention

By default, save the completed report as a new file next to the template using a clear suffix such as:

- `report-finished.docx`
- `motor-report-finished.docx`
- `kinematics-report-finished.docx`

Never overwrite the original template unless the user explicitly asks for that.
