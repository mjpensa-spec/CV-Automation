# Quick Start Guide

This guide will help you get started with the CV Automation script.

## Installation

1. Ensure you have Python 3.7+ installed:
```bash
python3 --version
```

2. Install dependencies:
```bash
pip install -r requirements.txt
```

## Running the Examples

The repository includes sample files in the `examples/` directory. Try running:

```bash
python cv_automation.py \
  --cv examples/sample_cv.txt \
  --instructions examples/sample_instructions.xlsx \
  --template examples/sample_template.pptx \
  --job-description examples/sample_job_description.txt \
  --output-dir ./output
```

This will create:
- `output/CV_Presentation_YYYYMMDD_HHMMSS.pptx` - The generated PowerPoint
- `output/Traceability_Report_YYYYMMDD_HHMMSS.json` - Processing log
- `output/cv_automation.log` - Detailed execution log

## Using Your Own Files

### 1. Prepare Your CV

Save your CV as a text file (`.txt`). The script will extract sections automatically.

Example structure:
```
YOUR NAME
Job Title

SUMMARY
Your professional summary here...

EXPERIENCE
Company | Role | Dates
- Responsibilities and achievements

SKILLS
List of your skills
```

### 2. Create Instructions File

Create an Excel file (`.xlsx`) with these columns:

| Section | Field | Rule | Value |
|---------|-------|------|-------|
| Summary | Length | max_words | 150 |
| Experience | Companies | include_all | Company Name |
| Skills | Highlight | top_skills | Python,AWS |

### 3. Prepare PowerPoint Template

Use any `.pptx` file as a template. The script will add slides to it.

### 4. Run the Script

```bash
python cv_automation.py \
  --cv your_cv.txt \
  --instructions your_rules.xlsx \
  --template your_template.pptx \
  --output-dir ./my_output
```

## Command Options

- `--cv` (required): Path to your CV file
- `--instructions` (required): Path to Excel instruction file
- `--template` (required): Path to PowerPoint template
- `--job-description` (optional): Path to job description file
- `--output-dir` (optional): Where to save output files (default: current directory)
- `--verbose` or `-v` (optional): Show detailed logging

## Troubleshooting

### Missing Dependencies
If you see "ModuleNotFoundError", install dependencies:
```bash
pip install -r requirements.txt
```

### File Not Found
Ensure all file paths are correct. Use absolute paths or paths relative to where you run the script.

### Permission Errors
Make sure you have write permissions in the output directory.

## Next Steps

- Customize the instruction file for your needs
- Create branded PowerPoint templates
- Add more CV content sections
- Review the traceability report to understand processing

## Getting Help

Run the script with `--help` for quick reference:
```bash
python cv_automation.py --help
```

Check the main README.md for detailed documentation.
