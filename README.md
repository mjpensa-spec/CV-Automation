# CV-Automation

A Python script for automating the creation of PowerPoint slides from CV data, instructions, and optionally a job description.

## Purpose

This script processes:
- **CV files** (PDF, DOCX, or text format)
- **Excel instruction files** for customization rules
- **PowerPoint templates** for consistent formatting
- **Job descriptions** (optional) for tailored output

And generates:
- **Client-facing PowerPoint slides** with formatted CV content
- **Traceability reports** documenting all processing steps

## Features

- 📄 Multi-format CV parsing (TXT, PDF, DOCX)
- 📊 Excel-based instruction system for flexible customization
- 🎨 PowerPoint template support for branded output
- 🎯 Job description integration for targeted presentations
- 📝 Comprehensive traceability reporting
- 🔍 Detailed logging for debugging and auditing
- ⚡ Command-line interface for easy automation

## Installation

1. Clone this repository:
```bash
git clone https://github.com/mjpensa-spec/CV-Automation.git
cd CV-Automation
```

2. Install Python dependencies:
```bash
pip install -r requirements.txt
```

## Usage

### Basic Usage

```bash
python cv_automation.py \
  --cv path/to/resume.pdf \
  --instructions path/to/rules.xlsx \
  --template path/to/template.pptx
```

### With Job Description

```bash
python cv_automation.py \
  --cv path/to/resume.pdf \
  --instructions path/to/rules.xlsx \
  --template path/to/template.pptx \
  --job-description path/to/job_description.txt
```

### With Custom Output Directory

```bash
python cv_automation.py \
  --cv path/to/resume.pdf \
  --instructions path/to/rules.xlsx \
  --template path/to/template.pptx \
  --output-dir ./output
```

### Command-Line Options

| Option | Required | Description |
|--------|----------|-------------|
| `--cv` | Yes | Path to CV file (PDF, DOCX, or TXT) |
| `--instructions` | Yes | Path to Excel instruction file |
| `--template` | Yes | Path to PowerPoint template file |
| `--job-description` | No | Path to job description file |
| `--output-dir` | No | Output directory (default: current directory) |
| `--verbose`, `-v` | No | Enable verbose logging |

## Excel Instruction File Format

The Excel instruction file should contain the following columns:

| Section | Field | Rule | Value |
|---------|-------|------|-------|
| Summary | Length | max_words | 100 |
| Experience | Include | company | ABC Corp |
| Skills | Highlight | skill | Python |

Example rules:
- **Section**: CV section to process (Summary, Experience, Education, Skills, etc.)
- **Field**: Specific field within the section
- **Rule**: Processing rule to apply
- **Value**: Value or parameter for the rule

## Output Files

The script generates two main outputs:

### 1. PowerPoint Presentation
- Filename: `CV_Presentation_YYYYMMDD_HHMMSS.pptx`
- Contains: Formatted CV content based on template and instructions

### 2. Traceability Report
- Filename: `Traceability_Report_YYYYMMDD_HHMMSS.json`
- Contains: Complete log of processing steps and decisions
- Format: JSON for easy parsing and auditing

## Project Structure

```
CV-Automation/
├── cv_automation.py      # Main script
├── requirements.txt      # Python dependencies
├── README.md            # This file
└── examples/            # Example files (optional)
    ├── sample_cv.txt
    ├── sample_instructions.xlsx
    └── sample_template.pptx
```

## Development

### Code Structure

The script is organized into a main `CVAutomation` class with the following key methods:

- `validate_inputs()`: Validate all input files
- `parse_cv()`: Extract data from CV file
- `parse_instructions()`: Load and process instruction rules
- `parse_job_description()`: Process optional job description
- `generate_powerpoint()`: Create output presentation
- `generate_traceability_report()`: Generate audit trail
- `run()`: Execute complete workflow

### Extending the Script

The modular design makes it easy to enhance:

1. **Add new CV parsers**: Extend `parse_cv()` method
2. **Add custom rules**: Extend instruction processing in `parse_instructions()`
3. **Modify slide layout**: Update `_populate_slides()` method
4. **Add output formats**: Create new generation methods

### Logging

Logs are written to:
- Console (stdout)
- File: `cv_automation.log` in the output directory

## Requirements

- Python 3.7+
- pandas
- openpyxl
- python-pptx

Optional (for enhanced CV parsing):
- PyPDF2 or pdfplumber (for PDF files)
- python-docx (for DOCX files)

## License

[Specify your license here]

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## Support

For issues and questions, please open an issue on GitHub.

## Roadmap

Future enhancements may include:
- [ ] Enhanced CV parsing with NLP
- [ ] AI-powered content summarization
- [ ] Multiple output templates
- [ ] Batch processing support
- [ ] Web interface
- [ ] API endpoints
- [ ] Database integration
- [ ] Multi-language support