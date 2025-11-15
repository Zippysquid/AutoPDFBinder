#!/bin/bash
# Load API keys from .env and run autopdfbinder with AI

# Load .env file
if [ -f .env ]; then
    export $(cat .env | grep -v '^#' | xargs)
    echo "✓ Loaded API keys from .env"
else
    echo "✗ Error: .env file not found"
    exit 1
fi

# Check which API keys are available
has_openai=false
has_anthropic=false

if [ -n "$OPENAI_API_KEY" ]; then
    echo "✓ OpenAI API key found"
    has_openai=true
fi

if [ -n "$ANTHROPIC_API_KEY" ]; then
    echo "✓ Anthropic API key found"
    has_anthropic=true
fi

if [ "$has_openai" = false ] && [ "$has_anthropic" = false ]; then
    echo "✗ Error: No API keys found in .env (need OPENAI_API_KEY or ANTHROPIC_API_KEY)"
    exit 1
fi

echo ""
echo "Running AutoPDFBinder with AI features..."
echo "Provider priority: OpenAI (primary) → Anthropic (fallback) → None (default)"
echo ""

# Run the script with AI enabled
python3 autopdfbinder.py --use-ai "$@"
