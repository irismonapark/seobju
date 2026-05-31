"""Vercel serverless entry for the payroll Flask app."""
import os
import sys

PAYROLL_DIR = os.path.abspath(
    os.path.join(os.path.dirname(__file__), "..", "artifacts", "payroll-system")
)

if not os.path.isdir(PAYROLL_DIR):
    raise RuntimeError(f"Payroll app directory not found: {PAYROLL_DIR}")

sys.path.insert(0, PAYROLL_DIR)
os.chdir(PAYROLL_DIR)

from app import app  # noqa: E402
