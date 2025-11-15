# AI-Powered Document Intelligence

AutoPDFBinder now includes **optional AI-powered document analysis** using Claude AI to dramatically improve document processing quality.

## What It Does

Instead of using just filenames, Claude analyzes the actual content of your documents and extracts:

- **Real Document Titles** - Extracted from content, not filenames
- **Descriptions** - Concise summaries of each document
- **Document Type** - Contract, Invoice, Report, Letter, etc.
- **Key Metadata** - Dates, parties involved, reference numbers
- **Quality Checks** - Detects illegible or corrupted pages

## Before vs After

### WITHOUT AI:
```
Cover Page:
  DOCUMENT 1
  contract_v3_final_FINAL.pdf
  Document Type: PDF Document
  File Size: 1.2 MB

TOC Entry:
  1. contract_v3_final_FINAL.pdf ............. 003
```

### WITH AI:
```
Cover Page:
  DOCUMENT 1 | Legal
  Employment Agreement - John Doe
  Fixed-term employment contract for software engineer position

  Document Type: Contract
  Date: 2024-03-15
  Parties: TechCorp Inc., John Doe
  Reference: EMP-2024-001
  File Size: 1.2 MB

  Analyzed by Claude AI

TOC Entry:
  1. Employment Agreement - John Doe (2024-03-15) .... 003
```

**MUCH MORE PROFESSIONAL!** ✨

## Setup

### 1. Install the Anthropic SDK
```bash
pip install anthropic
```

### 2. Get an API Key
- Sign up at https://console.anthropic.com/
- Create an API key
- Set it as an environment variable:

**Windows:**
```cmd
set ANTHROPIC_API_KEY=sk-ant-api03-...
```

**Linux/Mac:**
```bash
export ANTHROPIC_API_KEY=sk-ant-api03-...
```

### 3. Enable AI Features

**Option A: Command Line (Quick)**
```bash
python autopdfbinder.py --use-ai
```

**Option B: Config File (Persistent)**
Edit `config.json`:
```json
{
  "ai_features": {
    "enabled": true,
    "max_pages_to_analyze": 3,
    "model": "claude-3-5-sonnet-20241022"
  }
}
```

## Usage Examples

### Basic AI Analysis
```bash
# Analyze first 3 pages of each document (default)
python autopdfbinder.py --use-ai
```

### Custom Page Count
```bash
# Analyze only first page (faster, cheaper)
python autopdfbinder.py --use-ai --ai-pages 1

# Analyze first 5 pages (more thorough)
python autopdfbinder.py --use-ai --ai-pages 5
```

### Combined with Other Features
```bash
# AI + custom output + starting Bates number
python autopdfbinder.py --use-ai --output case_bundle.pdf --bates-start 100
```

## Cost Estimation

Claude 3.5 Sonnet pricing:
- **Input**: $3 / million tokens
- **Output**: $15 / million tokens

**Per document** (3 pages analyzed):
- Images: ~2000 tokens
- Analysis: ~200 tokens
- **Cost: ~$0.006 per document**

**Real-world examples:**
- 10 documents: ~$0.06
- 100 documents: ~$0.60
- 1000 documents: ~$6.00

**Extremely affordable for the value added!**

## What Gets Analyzed

### Supported Documents:
- ✅ **PDF files** - Analyzed directly
- ⚠️ **DOCX files** - Currently skipped (would need Windows conversion first)

### Information Extracted:

1. **Title** - The actual document title from headers/content
2. **Description** - One-line summary (max 100 characters)
3. **Document Type** - Contract, Invoice, Report, Letter, Agreement, Memo, Other
4. **Date** - Document date in YYYY-MM-DD format
5. **Parties** - People/organizations mentioned
6. **Reference Number** - Document/case/invoice numbers
7. **Category** - Legal, Financial, Administrative, Technical, Other
8. **Quality Issues** - Warnings about illegible/corrupt pages

## Enhanced Cover Pages

Cover pages now include:
- ✓ Real title instead of filename
- ✓ Document category in header
- ✓ Concise description
- ✓ Extracted date and parties
- ✓ Reference/case numbers
- ✓ Quality warnings if detected
- ✓ "Analyzed by Claude AI" attribution

## Enhanced Table of Contents

TOC entries show meaningful information:
```
Instead of: "1. doc_final.pdf"
Shows:      "1. Q1 Financial Report (2024-03-31)"

Instead of: "2. contract_signed.pdf"
Shows:      "2. Service Agreement - Smith Corp"
```

## Performance

- **Speed**: ~2-3 seconds per document for AI analysis
- **Accuracy**: Claude 3.5 Sonnet has excellent document understanding
- **Caching**: Results can be cached to avoid re-analyzing same documents
- **Parallel**: Could be enhanced to analyze multiple documents simultaneously

## Privacy & Security

- **Your data**: Sent to Anthropic's API (see their privacy policy)
- **Retention**: Anthropic doesn't train on API data
- **API Key**: Keep your ANTHROPIC_API_KEY secure
- **Sensitive docs**: Disable AI for confidential documents if needed

## Troubleshooting

### "AI features requested but anthropic package not installed"
```bash
pip install anthropic
```

### "AI features enabled but ANTHROPIC_API_KEY not set"
```bash
export ANTHROPIC_API_KEY=your-key-here
```

### "AI analysis failed for [file]"
- Check if PDF is corrupted
- Verify API key is valid
- Check internet connection
- See script_log.txt for details

### "AI returned invalid JSON"
- Rare issue with Claude response format
- File will still process, just without AI metadata
- Logged in script_log.txt

## Limitations

**Current:**
- Only analyzes PDF files (DOCX requires conversion)
- Sequential processing (one at a time)
- No offline mode

**Future Enhancements:**
- Batch processing for speed
- Support for DOCX analysis
- Local caching of results
- Custom prompts per document type
- Multi-language support

## Example Output

See `AI_INTEGRATION_DESIGN.md` for detailed examples and technical details.

## Disable AI

If you don't want AI features:

**Command line:**
```bash
python autopdfbinder.py  # Just don't use --use-ai flag
```

**Config file:**
```json
{
  "ai_features": {
    "enabled": false
  }
}
```

The script works perfectly fine without AI - it's a completely optional enhancement!

---

**Ready to make your PDF bundles truly intelligent?** 🧠✨

```bash
pip install anthropic
export ANTHROPIC_API_KEY=your-key
python autopdfbinder.py --use-ai
```
