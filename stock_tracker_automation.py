"""
Wrapper for stock_tracker.py to maintain backward compatibility with stock_tracker_automation.py
"""
import sys
from stock_tracker import main

if __name__ == "__main__":
    sys.exit(main())
