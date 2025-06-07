"""Compatibility wrapper for the renamed script.

This module forwards execution to ``autopdfbinder`` for users accustomed to
launching ``enhanced-document-generator.py``. The primary functionality now
lives in :mod:`autopdfbinder`.
"""

from autopdfbinder import main

if __name__ == "__main__":
    main()
