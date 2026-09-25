import os
import re
import io
import json
import time
import hashlib
import logging
import zipfile
import datetime as dt
import urllib.parse

from dataclasses import dataclass, asdict, field
from pathlib import Path
from concurrent.futures import ThreadPoolExecutor, as_completed

import requests
from bs4 import BeautifulSoup
import pandas as pd
import gspread

from google.oauth2 import service_account
from openai import OpenAI

from pydantic import BaseModel, Field
from pypdf import PdfReader
from docx import Document as DocxDocument

try:
    from playwright.sync_api import sync_playwright
except Exception:
    sync_playwright = None


# =============================================================================
# 1. CONFIGURATION
# =============================================================================

APP_NAME = "Pan-India Tender Intelligence Agent"

IST = dt.timezone(dt.timedelta(hours=5, minutes=30))

SERVICE_ACCOUNT_FILE = os.getenv(
    "SERVICE_ACCOUNT_FILE",
    "service_account.json",
)

SPREADSHEET_ID = os.getenv(
    "SPREADSHEET_ID",
    "",
)

EXCEL_MASTER_FILE = os.getenv(
    "EXCEL_MASTER_FILE",
    "Pan_India_Tender_URL_Master.xlsx",
)

OPENAI_API_KEY = os.getenv(
    "OPENAI_API_KEY",
    "",
)

# Current cost-sensitive OpenAI API model.
OPENAI_MODEL = os.getenv(
    "OPENAI_MODEL",
    "gpt-5.6-luna",
)

GEM_LISTING_URL = os.getenv(
    "GEM_LISTING_URL",
    "https://bidplus-global.gem.gov.in/",
)

DOWNLOAD_DIR = Path(
    os.getenv(
        "DOWNLOAD_DIR",
        "tender_downloads",
    )
)

DOWNLOAD_DIR.mkdir(
    parents=True,
    exist_ok=True,
)

REQUEST_TIMEOUT = int(
    os.getenv(
        "REQUEST_TIMEOUT",
        "20",
    )
)

CRAWL_WORKERS = int(
    os.getenv(
        "CRAWL_WORKERS",
        "10",
    )
