# Diagram Generator

AI-powered diagram generator with editable PPTX export. No login required.

[![Deploy to Render](https://render.com/images/deploy-to-render-button.svg)](https://render.com/deploy?repo=https://github.com/Gukki2021/diagram-generator)

## Features

- AI diagram generation (Google Gemini 2.5 Flash)
- 11 diagram types: 2x2 Matrix, Process Flow, Pyramid, Venn, Timeline, Waterfall, Radar, Funnel, Porter's 5 Forces, Framework
- 6 color themes + custom color picker
- 13 fonts including Aptos
- Export to **SVG**, **PNG**, and **editable PPTX** (native shapes, not images)
- Image/sketch input with drag-and-drop
- No login required for end users

## One-Click Deploy

Click the **Deploy to Render** button above, then set:

- `GEMINI_API_KEY` = your Google Gemini API key ([get one free](https://aistudio.google.com/apikey))

## Run Locally

```bash
git clone https://github.com/Gukki2021/diagram-generator.git
cd diagram-generator
pip install -r requirements.txt
export GEMINI_API_KEY=your_key_here
python app.py
```

Open http://localhost:5555

## Docker

```bash
docker build -t diagram-generator .
docker run -p 5555:10000 -e GEMINI_API_KEY=your_key diagram-generator
```

## Tech Stack

- Python / Flask
- Google Gemini 2.5 Flash API
- python-pptx (editable PPTX export)
- Vanilla JS frontend
