"""Vercel serverless entry for ResumeConverter Flask app."""
import os
import sys

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
sys.path.insert(0, ROOT)
os.chdir(ROOT)

from app import app  # noqa: E402
