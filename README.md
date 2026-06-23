# Diagram Generator

AI-powered diagram generator with editable PPTX export. No login required.

[![Deploy to Render](https://render.com/images/deploy-to-render-button.svg)](https://render.com/deploy?repo=https://github.com/Gukki2021/diagram-generator)

## Features

- AI diagram generation (Anthropic Claude Opus 4.7, with sketch-to-diagram vision input)
- 11 diagram types: 2x2 Matrix, Process Flow, Pyramid, Venn, Timeline, Waterfall, Radar, Funnel, Porter's 5 Forces, Framework
- 6 color themes + custom color picker
- 13 fonts including Aptos
- Export to **SVG**, **PNG**, and **editable PPTX** (native shapes, not images)
- Image/sketch input with drag-and-drop
- Mobile-friendly (responsive web UI, works in any phone browser)
- No login required for end users

## One-Click Deploy

Click the **Deploy to Render** button above, then set:

- `ANTHROPIC_API_KEY` = your Anthropic API key ([create one](https://console.anthropic.com/settings/keys))

Optionally override the model:

- `CLAUDE_MODEL` = `claude-opus-4-7` (default), or `claude-sonnet-4-6` for lower cost

## Run Locally

```bash
git clone https://github.com/Gukki2021/diagram-generator.git
cd diagram-generator
python3 -m venv .venv && source .venv/bin/activate
pip install -r requirements.txt
echo "ANTHROPIC_API_KEY=sk-ant-..." > .env
python app.py
```

Open http://localhost:5555

## Docker

```bash
docker build -t diagram-generator .
docker run -p 5555:10000 -e ANTHROPIC_API_KEY=your_key diagram-generator
```

## Tech Stack

- Python / Flask
- Anthropic Claude API (Opus 4.7 with adaptive thinking + vision)
- python-pptx (editable PPTX export)
- Vanilla JS frontend
