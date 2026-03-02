# -*- coding: utf-8 -*-
# ==========================
# ===== INCIO DA PARTE 1 =====
# ==========================
import json
import logging
import os
import re
import sqlite3
import sys
import tempfile
import tkinter as tk
import urllib.error
import urllib.parse
import urllib.request
import webbrowser
import shutil
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from tkinter import filedialog, messagebox, simpledialog, ttk
try:
    import pandas as pd
except Exception:
    pd = None
from datetime import datetime, timedelta
try:
    from reportlab.pdfgen import canvas
    from reportlab.lib.pagesizes import A4
except Exception:
    canvas = None
    A4 = None
from contextlib import contextmanager
import random
import string
import base64
import hashlib
import hmac
import secrets
import ctypes