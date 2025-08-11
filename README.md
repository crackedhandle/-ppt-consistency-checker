# PPTX Consistency Checker

This project provides a command-line tool to analyze PowerPoint presentations (`.pptx` files) for inconsistencies in content. It combines text extraction, OCR, and AI-powered analysis using Google’s Gemini API to detect numerical, textual, timeline, and logical inconsistencies across slides.

## Features

- Extracts structured text from PowerPoint slides including slide titles, content, and notes.
- Converts slides to PDF and then to images to perform OCR text extraction.
- Uses Tesseract OCR to read any textual content embedded as images.
- Integrates with Google Gemini API to analyze combined text for:
  - Numerical inconsistencies (conflicting stats, revenue numbers, percentages)
  - Textual contradictions
  - Timeline mismatches (dates, schedules)
  - Logical flaws or contradictory conclusions
- Outputs a JSON file listing detected inconsistencies with slide references and confidence scores.
- Supports debugging with detailed logs for troubleshooting.

## How It Works

1. **Text Extraction**: Extracts text directly from slide elements and notes using `python-pptx`.
2. **Image Conversion**: Converts the presentation to PDF using LibreOffice, then to images using Poppler.
3. **OCR**: Applies Tesseract OCR on slide images to capture any text embedded as images.
4. **Content Aggregation**: Combines extracted structured text and OCR text for each slide.
5. **AI Analysis**: Sends aggregated slide content to Google Gemini API for inconsistency detection.
6. **Output**: Saves the AI response as a JSON file and logs key findings.

## Limitations

- Requires LibreOffice installed and accessible via command line for PPTX to PDF conversion.
- Poppler utilities must be installed for PDF to image conversion.
- Tesseract OCR must be installed and configured in the system environment.
- Google Gemini API key is mandatory and usage might incur costs or require quota.
- Accuracy depends on the quality of slide content and OCR results.
- Large presentations might slow down due to PDF/image conversions and API calls.

## Usage

### Prerequisites

- Python 3.7+
- LibreOffice installed and added to system PATH
- Poppler utilities installed (and provide path with `--poppler` if needed)
- Tesseract OCR installed and configured (`TESSDATA_PREFIX` environment variable if on Windows)
- Google Gemini API key (set as environment variable `GOOGLE_API_KEY` or pass with `--api-key`)

### Clone the Repository

```bash
git clone https://github.com/crackedhandle/pptx-consistency-checker.git
cd pptx-consistency-checker

