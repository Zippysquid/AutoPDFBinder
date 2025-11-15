# AI-Powered Document Intelligence

AutoPDFBinder now includes **optional AI-powered document analysis** using OpenAI or Claude AI to dramatically improve document processing quality.

**Multi-Provider Support:**
- **OpenAI GPT-4o** (Primary) - Latest vision model with excellent document understanding
- **Anthropic Claude** (Fallback) - Backup provider if OpenAI unavailable
- **Auto Fallback** - Automatically tries OpenAI first, falls back to Claude, or skips AI if both fail

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

### 1. Install AI SDKs

**Option A: OpenAI (Recommended)**
```bash
pip install openai
```

**Option B: Anthropic Claude**
```bash
pip install anthropic
```

**Option C: Both (Best - automatic fallback)**
```bash
pip install openai anthropic
```

### 2. Get API Keys

**OpenAI:**
- Sign up at https://platform.openai.com/
- Create an API key under "API keys"
- Set environment variable:
```bash
export OPENAI_API_KEY=sk-proj-...
```

**Anthropic (Optional Fallback):**
- Sign up at https://console.anthropic.com/
- Create an API key
- Set environment variable:
```bash
export ANTHROPIC_API_KEY=sk-ant-api03-...
```

**Using .env File (Recommended):**
```bash
# Create .env file in project root
echo "OPENAI_API_KEY=sk-proj-..." >> .env
echo "ANTHROPIC_API_KEY=sk-ant-api03-..." >> .env

# Use the helper script
./run_with_ai.sh
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
    "provider": "auto",
    "openai_model": "gpt-4o",
    "anthropic_model": "claude-3-5-sonnet-20241022"
  }
}
```

**Provider Options:**
- `"auto"` - Try OpenAI first, fallback to Anthropic (recommended)
- `"openai"` - Use OpenAI only
- `"anthropic"` - Use Anthropic only

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

### OpenAI GPT-4o (Primary)
- **Input**: $2.50 / million tokens
- **Output**: $10 / million tokens

**Per document** (3 pages analyzed):
- Images: ~2000 tokens
- Analysis: ~200 tokens
- **Cost: ~$0.005 per document**

**Real-world examples:**
- 10 documents: ~$0.05
- 100 documents: ~$0.50
- 1000 documents: ~$5.00

### Anthropic Claude 3.5 Sonnet (Fallback)
- **Input**: $3 / million tokens
- **Output**: $15 / million tokens

**Per document** (3 pages analyzed):
- **Cost: ~$0.006 per document**

**Real-world examples:**
- 10 documents: ~$0.06
- 100 documents: ~$0.60
- 1000 documents: ~$6.00

### GPT-4o Mini (Budget Option)
Set `"openai_model": "gpt-4o-mini"` in config:
- **Cost: ~$0.0005 per document** (10x cheaper!)

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

## Multi-Provider Fallback System

The AI integration supports multiple providers with automatic fallback:

**Priority Order (with `provider: "auto"`):**
1. **OpenAI GPT-4o** - Tried first if API key available
2. **Anthropic Claude** - Fallback if OpenAI fails or unavailable
3. **No AI** - Continues without AI if both fail (graceful degradation)

**Benefits:**
- ✓ Higher reliability (redundancy if one provider is down)
- ✓ Cost optimization (use cheaper provider)
- ✓ Quota management (switch providers if one runs out)
- ✓ Never blocks processing (always continues even if AI fails)

## Enhanced Cover Pages

Cover pages now include:
- ✓ Real title instead of filename
- ✓ Document category in header
- ✓ Concise description
- ✓ Extracted date and parties
- ✓ Reference/case numbers
- ✓ Quality warnings if detected
- ✓ "Analyzed by AI" attribution (shows which provider used)

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
- **Accuracy**: Both GPT-4o and Claude 3.5 Sonnet have excellent document understanding
- **Provider Speed**: OpenAI GPT-4o is typically faster than Claude
- **Caching**: Results can be cached to avoid re-analyzing same documents
- **Parallel**: Could be enhanced to analyze multiple documents simultaneously

## Privacy & Security

- **Your data**: Sent to OpenAI and/or Anthropic's API (see respective privacy policies)
- **Retention**: Neither provider trains on API data by default
- **API Keys**: Keep OPENAI_API_KEY and ANTHROPIC_API_KEY secure (use .env file)
- **Sensitive docs**: Disable AI for confidential documents if needed
- **Provider choice**: Use `"provider": "openai"` or `"provider": "anthropic"` to limit to one provider

## Troubleshooting

### "AI features enabled but no AI packages installed"
```bash
# Install one or both providers
pip install openai anthropic
```

### "AI features enabled but no API keys set"
```bash
# OpenAI (primary)
export OPENAI_API_KEY=sk-proj-your-key-here

# Anthropic (fallback)
export ANTHROPIC_API_KEY=sk-ant-api03-your-key-here

# Or use .env file (recommended)
echo "OPENAI_API_KEY=sk-proj-..." >> .env
echo "ANTHROPIC_API_KEY=sk-ant-..." >> .env
./run_with_ai.sh
```

### "OpenAI analysis failed" followed by "Anthropic analysis failed"
- Both providers unavailable - check:
  - API keys are valid
  - Account has credits/quota
  - Internet connection working
- Script will continue without AI (graceful degradation)

### "Error code: 429 - insufficient_quota"
- Account out of credits
- Check billing at platform.openai.com or console.anthropic.com
- Fallback provider will be tried automatically

### "AI returned invalid JSON"
- Rare issue with AI response format
- File will still process, just without AI metadata
- Logged in script_log.txt with details

### Provider-specific issues

**OpenAI:**
- Verify key at https://platform.openai.com/api-keys
- Check usage at https://platform.openai.com/usage

**Anthropic:**
- Verify key at https://console.anthropic.com/settings/keys
- Check usage at https://console.anthropic.com/settings/billing

## Limitations

**Current:**
- Only analyzes PDF files (DOCX requires conversion)
- Sequential processing (one at a time)
- Requires internet connection
- Requires API credits from at least one provider

**Future Enhancements:**
- Batch/parallel processing for speed
- Support for DOCX analysis
- Persistent caching of results
- Custom prompts per document type
- Multi-language support
- Additional AI providers (Azure OpenAI, Google Gemini, etc.)

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
