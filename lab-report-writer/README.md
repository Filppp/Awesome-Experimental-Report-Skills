# Lab Report Writer

![Codex Skill](https://img.shields.io/badge/Codex-Skill-black)
![Chinese Output](https://img.shields.io/badge/output-Chinese%20Lab%20Report-blue)
![Windows Friendly](https://img.shields.io/badge/Windows-Word%20COM%20Preferred-0078D4)
![MIT License](https://img.shields.io/badge/license-MIT-green)

A reusable Codex skill for generating complete Chinese laboratory reports from a Word template, reference manuals, and result images.

This repository is structured so that the repository root is also the skill root. If you place or clone this directory into your Codex skills directory, it can be used directly.

## Features

- Generate structured Chinese lab reports from existing source materials
- Reuse `.docx` templates instead of drafting from scratch
- Extract only the useful facts from manuals, guides, and screenshots
- Organize reports into a stable section layout suitable for student submissions
- Support result-image insertion and conclusion-oriented presentation
- Prefer reliable Windows Word COM editing when final `.docx` fidelity matters

## What It Does

`lab-report-writer` is designed for tasks where the user already has:

- a `.docx` report template
- reference PDFs such as manuals or experiment guides
- screenshots or on-site result images
- a clear experiment topic

The skill helps Codex turn those materials into a finished Chinese lab report with a stable section structure and practical student-report wording.

Typical sections include:

1. objectives
2. equipment or tools
3. principles and operating process
4. reflections
5. content and steps
6. conclusion and takeaways

## Quick Start

Use this skill when you want Codex to produce a complete Chinese experiment report rather than raw notes.

```text
Use $lab-report-writer to complete a Chinese lab report from a Word template, reference PDFs, and result images, and deliver a submission-ready DOCX.
```

## Installation

This repository is directly usable as a Codex skill.

### Option 1: Clone Into Your Skills Directory

Clone the repository so that the repository root is the skill root.

```bash
git clone https://github.com/<your-name>/lab-report-writer.git <your-skills-dir>/lab-report-writer
```

Examples of target skill directories depend on your Codex setup:
- `~/.codex/skills/lab-report-writer`

### Option 2: Copy The Folder Manually

Copy the full folder into your skills directory and keep the root structure unchanged.

The minimum required files are:

- `SKILL.md`
- `agents/openai.yaml`

### Option 3: Use The Repository As-Is

If your Codex installation already loads skills from the directory where this repository is placed, no extra packaging is needed.

## Repository Layout

```text
lab-report-writer/
├── SKILL.md
├── README.md
├── LICENSE
├── .gitignore
└── agents/
    └── openai.yaml
```

## Trigger Intent

Use this skill when the task is to complete a structured Chinese experiment report from source materials, especially when the user expects a finished `.docx` output rather than raw notes.

Common triggers include:

- completing a lab report from a template
- filling standard report sections from a PDF guide
- inserting result images into the conclusion section
- reusing the same template across multiple experiments

## Workflow Summary

The skill follows this high-level process:

1. confirm inputs and final deliverable
2. identify the report structure from the template
3. extract only useful facts from manuals and images
4. draft each section with controlled detail
5. generate the final `.docx`
6. insert images with meaningful captions
7. validate the final output

## Environment Notes

The skill is platform-agnostic at the instruction level, but on Windows it explicitly prefers Word COM automation through Python for reliable `.docx` template editing and image insertion.

If you already have another DOCX-focused skill, this skill is intended to pair with it.

## Why This Repository Is Safe To Publish

This public version is sanitized for open-source release:

- no personal file paths
- no personal identifiers
- no private reports
- no bundled images or manuals
- no local machine logs or outputs

It contains only reusable skill metadata and instructions.

## Scope

This skill is optimized for structured lab-report generation.

It is not intended for:

- open-ended academic paper writing
- fabricating missing experiment data
- publishing private manuals or images
- replacing a full document-processing toolchain by itself

## License

This repository is released under the MIT License. See [LICENSE](LICENSE).
