#!/usr/bin/env python3
"""
Streamlit entry point for Community Cloud
Imports and runs the main dashboard application.
"""

import sys
import os

# Add current directory to path so imports work
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

# Import the dashboard module directly
import importlib.util

spec = importlib.util.spec_from_file_location("despatch_dashboard", "despatch_dashboard")
dashboard = importlib.util.module_from_spec(spec)
spec.loader.exec_module(dashboard)
