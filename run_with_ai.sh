#!/bin/bash
# Load API key from .env and run autopdfbinder with AI

# Load .env file
if [ -f .env ]; then
    export $(cat .env | grep -v '^#' | xargs)
    echo "✓ Loaded API key from .env"
else
    echo "✗ Error: .env file not found"
    exit 1
fi

# Check if API key is set
if [ -z "$ANTHROPIC_API_KEY" ]; then
    echo "✗ Error: ANTHROPIC_API_KEY not found in .env"
    exit 1
fi

echo "✓ API key loaded"
echo ""
echo "Running AutoPDFBinder with AI features..."
echo ""

# Run the script with AI enabled
python3 autopdfbinder.py --use-ai "$@"
