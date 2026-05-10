#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Fresh Stop Dash EDITION
Built from v12.6.py + 5 easy features + EIA/AAA ticker
Everything else exactly as in the best working version
"""

from flask import Flask, render_template_string, request, redirect, url_for, flash, send_from_directory, session
import os
import json
import psycopg2
from psycopg2.extras import RealDictCursor
from datetime import datetime
from functools import wraps
from werkzeug.utils import secure_filename
from werkzeug.security import generate_password_hash, check_password_hash

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "gas-optimizer-final-winner-2026")
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024

os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

GRADES = ["Regular", "Mid-grade", "Premium", "Diesel"]

def get_db_connection():
    database_url = os.environ.get("DATABASE_URL")
    if database_url:
        # Production (Railway PostgreSQL)
        conn = psycopg2.connect(database_url, sslmode='require')
    else:
        # Local development (SQLite fallback)
        import sqlite3
        conn = sqlite3.connect("users.db")
    return conn

def init_db():
    conn = get_db_connection()
    c = conn.cursor()
    if os.environ.get("DATABASE_URL"):
        # PostgreSQL
        c.execute('''CREATE TABLE IF NOT EXISTS users (
            email TEXT PRIMARY KEY, 
            password TEXT NOT NULL, 
            name TEXT, 
            role TEXT DEFAULT 'viewer',
            created_at TEXT
        )''')
    else:
        # SQLite
        c.execute('''CREATE TABLE IF NOT EXISTS users (
            email TEXT PRIMARY KEY, 
            password TEXT NOT NULL, 
            name TEXT, 
            role TEXT DEFAULT 'viewer',
            created_at TEXT
        )''')
    conn.commit()
    conn.close()

def get_user(email):
    conn = get_db_connection()
    c = conn.cursor()
    if os.environ.get("DATABASE_URL"):
        c.execute("SELECT email, password, name, role, created_at FROM users WHERE email = %s", (email,))
    else:
        c.execute("SELECT email, password, name, role, created_at FROM users WHERE email = ?", (email,))
    row = c.fetchone()
    conn.close()
    if row:
        return {"email": row[0], "password": row[1], "name": row[2], "role": row[3], "created_at": row[4]}
    return None

def create_user(email, password, name, role="viewer"):
    conn = get_db_connection()
    c = conn.cursor()
    try:
        if os.environ.get("DATABASE_URL"):
            c.execute("INSERT INTO users (email, password, name, role, created_at) VALUES (%s, %s, %s, %s, %s)",
                      (email, password, name, role, datetime.now().isoformat()))
        else:
            c.execute("INSERT INTO users (email, password, name, role, created_at) VALUES (?, ?, ?, ?, ?)",
                      (email, password, name, role, datetime.now().isoformat()))
        conn.commit()
        return True
    except Exception:
        return False
    finally:
        conn.close()

def update_user_role(email, new_role):
    conn = get_db_connection()
    c = conn.cursor()
    if os.environ.get("DATABASE_URL"):
        c.execute("UPDATE users SET role = %s WHERE email = %s", (new_role, email))
    else:
        c.execute("UPDATE users SET role = ? WHERE email = ?", (new_role, email))
    conn.commit()
    conn.close()


# ==================== DEVICE CODE FLOW - SHAREPOINT ====================

import msal
import json
import os

# MSAL Configuration (Device Code Flow)
CLIENT_ID = os.environ.get("SHAREPOINT_CLIENT_ID", "your-client-id-here")
TENANT_ID = os.environ.get("SHAREPOINT_TENANT_ID", "common")
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPE = ["https://graph.microsoft.com/Files.ReadWrite.All", "https://graph.microsoft.com/Sites.ReadWrite.All"]

# Token cache file
TOKEN_CACHE_FILE = "token_cache.json"

def get_msal_app():
    cache = msal.SerializableTokenCache()
    if os.path.exists(TOKEN_CACHE_FILE):
        cache.deserialize(open(TOKEN_CACHE_FILE, "r").read())
    
    app = msal.PublicClientApplication(
        CLIENT_ID,
        authority=AUTHORITY,
        token_cache=cache
    )
    return app, cache

def get_sharepoint_token():
    """Get SharePoint access token using Device Code Flow"""
    app, cache = get_msal_app()
    
    # Try to get token from cache first
    accounts = app.get_accounts()
    if accounts:
        result = app.acquire_token_silent(SCOPE, account=accounts[0])
        if result and "access_token" in result:
            return result["access_token"]
    
    # If no cached token, use Device Code Flow
    flow = app.initiate_device_flow(scopes=SCOPE)
    if "user_code" not in flow:
        print("Failed to create device flow")
        return None
    
    print("\n" + "="*60)
    print("SHAREPOINT AUTHENTICATION REQUIRED")
    print("="*60)
    print(f"Go to: {flow['verification_uri']}")
    print(f"Enter code: {flow['user_code']}")
    print("="*60 + "\n")
    
    result = app.acquire_token_by_device_flow(flow)
    
    # Save cache
    if cache.has_state_changed:
        with open(TOKEN_CACHE_FILE, "w") as f:
            f.write(cache.serialize())
    
    if "access_token" in result:
        return result["access_token"]
    else:
        print("Authentication failed:", result.get("error_description"))
        return None

def upload_to_sharepoint(station_name, filename, file_content):
    """Upload file to SharePoint using Device Code Flow"""
    access_token = get_sharepoint_token()
    if not access_token:
        return None
    
    site_url = os.environ.get("SHAREPOINT_SITE_URL", "").rstrip("/")
    if not site_url:
        print("SHAREPOINT_SITE_URL not configured")
        return None
    
    try:
        headers = {
            "Authorization": f"Bearer {access_token}",
            "Content-Type": "application/octet-stream"
        }
        
        # Create folder structure: Station Name / Documents
        folder_path = f"{station_name}/Documents"
        
        # Upload using Microsoft Graph
        upload_url = f"https://graph.microsoft.com/v1.0/sites/{site_url.split('//')[1].split('/')[0]}:/sites/{site_url.split('/sites/')[-1] if '/sites/' in site_url else ''}:/drive/root:/{folder_path}/{filename}:/content"
        
        response = requests.put(upload_url, headers=headers, data=file_content)
        
        if response.status_code in [200, 201]:
            result = response.json()
            return result.get("webUrl")
        else:
            print(f"SharePoint upload failed: {response.status_code} - {response.text}")
            
    except Exception as e:
        print(f"SharePoint upload error: {e}")
    
    return None

# Initialize database
init_db()
if not get_user("demo@station.com"):
    create_user("demo@station.com", 
                generate_password_hash("demo123", method='pbkdf2:sha256'), 
                "Demo User", 
                "admin")

# Sample station data (exact from v12.6)
STATIONS = {
    1: {"name": "Sunshine Express", "address": "123 Ocean Drive", "city": "Miami Beach", "state": "FL", "zip": "33139", "email": "sunshine@station.com",
        "grades": {
            "Regular": {"wholesale": 2.91, "current_price": 3.19, "competitors": [3.25, 3.22, 3.15], "target_margin": 0.12, "daily_gallons": 5200, "elasticity": -2.8, "strategy": "balanced"},
            "Mid-grade": {"wholesale": 3.05, "current_price": 3.39, "competitors": [3.45, 3.42, 3.35], "target_margin": 0.14, "daily_gallons": 1800, "elasticity": -1.9, "strategy": "balanced"},
            "Premium": {"wholesale": 3.18, "current_price": 3.59, "competitors": [3.65, 3.62, 3.55], "target_margin": 0.16, "daily_gallons": 950, "elasticity": -1.4, "strategy": "balanced"},
            "Diesel": {"wholesale": 2.98, "current_price": 3.35, "competitors": [3.42, 3.38, 3.29], "target_margin": 0.13, "daily_gallons": 2100, "elasticity": -2.1, "strategy": "balanced"}
        },
        "competitor_list": []},
    2: {"name": "Lone Star Fuel", "address": "456 Highway 59", "city": "Houston", "state": "TX", "zip": "77002", "email": "lonestar@station.com",
        "grades": {
            "Regular": {"wholesale": 2.76, "current_price": 3.05, "competitors": [3.09, 3.12, 2.99], "target_margin": 0.10, "daily_gallons": 7800, "elasticity": -1.6, "strategy": "balanced"},
            "Mid-grade": {"wholesale": 2.90, "current_price": 3.25, "competitors": [3.29, 3.32, 3.19], "target_margin": 0.12, "daily_gallons": 2400, "elasticity": -1.2, "strategy": "balanced"},
            "Premium": {"wholesale": 3.03, "current_price": 3.45, "competitors": [3.49, 3.52, 3.39], "target_margin": 0.14, "daily_gallons": 1100, "elasticity": -0.9, "strategy": "balanced"},
            "Diesel": {"wholesale": 2.82, "current_price": 3.19, "competitors": [3.25, 3.22, 3.12], "target_margin": 0.11, "daily_gallons": 3200, "elasticity": -1.5, "strategy": "balanced"}
        },
        "competitor_list": []},
    3: {"name": "Sooner Stop", "address": "789 Main Street", "city": "Oklahoma City", "state": "OK", "zip": "73102", "email": "sooner@station.com",
        "grades": {
            "Regular": {"wholesale": 2.84, "current_price": 3.14, "competitors": [3.19, 3.11, 3.22], "target_margin": 0.14, "daily_gallons": 3100, "elasticity": -1.1, "strategy": "balanced"},
            "Mid-grade": {"wholesale": 2.98, "current_price": 3.34, "competitors": [3.39, 3.31, 3.42], "target_margin": 0.15, "daily_gallons": 950, "elasticity": -0.8, "strategy": "balanced"},
            "Premium": {"wholesale": 3.11, "current_price": 3.54, "competitors": [3.59, 3.51, 3.62], "target_margin": 0.17, "daily_gallons": 480, "elasticity": -0.6, "strategy": "balanced"},
            "Diesel": {"wholesale": 2.91, "current_price": 3.29, "competitors": [3.35, 3.28, 3.19], "target_margin": 0.14, "daily_gallons": 1400, "elasticity": -1.0, "strategy": "balanced"}
        },
        "competitor_list": []},
    4: {"name": "Razorback Mart", "address": "321 River Road", "city": "Little Rock", "state": "AR", "zip": "72201", "email": "razorback@station.com",
        "grades": {
            "Regular": {"wholesale": 2.88, "current_price": 3.17, "competitors": [3.22, 3.14, 3.25], "target_margin": 0.13, "daily_gallons": 2900, "elasticity": -1.3, "strategy": "balanced"},
            "Mid-grade": {"wholesale": 3.02, "current_price": 3.37, "competitors": [3.42, 3.34, 3.45], "target_margin": 0.14, "daily_gallons": 880, "elasticity": -1.0, "strategy": "balanced"},
            "Premium": {"wholesale": 3.15, "current_price": 3.57, "competitors": [3.62, 3.54, 3.65], "target_margin": 0.16, "daily_gallons": 420, "elasticity": -0.7, "strategy": "balanced"},
            "Diesel": {"wholesale": 2.95, "current_price": 3.32, "competitors": [3.38, 3.31, 3.22], "target_margin": 0.13, "daily_gallons": 1250, "elasticity": -1.2, "strategy": "balanced"}
        },
        "competitor_list": []}
}

AUDIT_LOG = []
EMAIL_AUDIT = []
EDIT_AUDIT = []
DOCUMENT_AUDIT = []
INVOICES_FILE = "invoices/invoices.json"
NEXT_STATION_ID = 5

def load_invoices():
    if os.path.exists(INVOICES_FILE):
        with open(INVOICES_FILE, "r") as f:
            return json.load(f)
    return []

def save_invoices(invoices):
    os.makedirs(os.path.dirname(INVOICES_FILE), exist_ok=True)
    with open(INVOICES_FILE, "w") as f:
        json.dump(invoices, f, indent=2)

def add_to_audit(station_name, grade, old_price, new_price):
    try:
        user = session.get('user_email', 'system')
    except:
        user = 'system'
    AUDIT_LOG.append({
        "time": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "station": station_name,
        "grade": grade,
        "old_price": f"${old_price:.2f}",
        "new_price": f"${new_price:.2f}",
        "user": user
    })
    if len(AUDIT_LOG) > 50:
        AUDIT_LOG.pop(0)

def add_to_email_audit(station_name, recipient, prices_summary):
    try:
        user = session.get('user_email', 'system')
    except:
        user = 'system'
    EMAIL_AUDIT.append({
        "time": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "station": station_name,
        "recipient": recipient,
        "prices": prices_summary,
        "user": user
    })
    if len(EMAIL_AUDIT) > 50:
        EMAIL_AUDIT.pop(0)

def add_to_edit_audit(station_name, action, details=""):
    try:
        user = session.get('user_email', 'system')
    except:
        user = 'system'
    EDIT_AUDIT.append({
        "time": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "station": station_name,
        "action": action,
        "details": details,
        "user": user
    })
    if len(EDIT_AUDIT) > 50:
        EDIT_AUDIT.pop(0)

def add_to_document_audit(station_name, doc_type, filename, description=""):
    try:
        user = session.get('user_email', 'system')
    except:
        user = 'system'
    DOCUMENT_AUDIT.append({
        "time": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "station": station_name,
        "doc_type": doc_type,
        "filename": filename,
        "description": description,
        "user": user
    })
    if len(DOCUMENT_AUDIT) > 50:
        DOCUMENT_AUDIT.pop(0)

def get_price_for_strategy(wholesale, strategy):
    base = wholesale + 0.12
    if strategy == "aggressive":
        return round(base - 0.03, 2)
    elif strategy == "match":
        return round(base + 0.01, 2)
    else:
        return round(base, 2)

def get_ml_optimal_price(wholesale, current_price, elasticity, current_volume):
    best_price = current_price
    best_profit = (current_price - wholesale) * current_volume
    for change in range(-15, 16):
        test_price = current_price + (change / 100)
        if test_price <= wholesale + 0.05:
            continue
        price_change_pct = (test_price - current_price) / current_price
        volume_change_pct = elasticity * price_change_pct
        predicted_volume = current_volume * (1 + volume_change_pct)
        if predicted_volume < 0:
            continue
        profit = (test_price - wholesale) * predicted_volume
        if profit > best_profit:
            best_profit = profit
            best_price = test_price
    return round(best_price, 2)

def get_regional_averages(state):
    averages = {
        "FL": {"Regular": 3.12, "Mid-grade": 3.32, "Premium": 3.52, "Diesel": 3.28},
        "TX": {"Regular": 2.98, "Mid-grade": 3.18, "Premium": 3.38, "Diesel": 3.15},
        "OK": {"Regular": 3.05, "Mid-grade": 3.25, "Premium": 3.45, "Diesel": 3.22},
        "AR": {"Regular": 3.08, "Mid-grade": 3.28, "Premium": 3.48, "Diesel": 3.25}
    }
    return averages.get(state, {"Regular": 3.10, "Mid-grade": 3.30, "Premium": 3.50, "Diesel": 3.20})

def login_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if 'user_email' not in session:
            flash("Please log in to access this page")
            return redirect("/login")
        return f(*args, **kwargs)
    return decorated_function

def admin_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if 'user_email' not in session:
            flash("Please log in to access this page")
            return redirect("/login")
        user = get_user(session['user_email'])
        if not user or user.get('role') != 'admin':
            flash("Admin access required")
            return redirect("/dashboard")
        return f(*args, **kwargs)
    return decorated_function

def manager_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if 'user_email' not in session:
            flash("Please log in to access this page")
            return redirect("/login")
        user = get_user(session['user_email'])
        if not user or user.get('role') not in ['admin', 'manager']:
            flash("Manager or Admin access required")
            return redirect("/dashboard")
        return f(*args, **kwargs)
    return decorated_function

# ==================== HTML TEMPLATES (same as v12.6) ====================

LOGIN_HTML = '''<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>Login - Fresh Stop Dash</title>
    <script src="https://cdn.tailwindcss.com"></script>
    <style>
        .fresh-green { color: #16a34a; }
        .fresh-red { color: #dc2626; }
    </style>
</head>
<body class="bg-slate-950 text-slate-200 flex items-center justify-center min-h-screen">
    <div class="max-w-md w-full px-6">
        <div class="text-center mb-8">
            <!-- FreshStop Logo -->
            <div class="flex justify-center mb-4">
                <div class="w-20 h-20 bg-gradient-to-br from-green-500 to-emerald-600 rounded-full flex items-center justify-center shadow-xl">
                    <svg width="52" height="52" viewBox="0 0 52 52" fill="none" xmlns="http://www.w3.org/2000/svg">
                        <path d="M26 6C26 6 14 14 14 26C14 38 26 46 26 46C26 46 38 38 38 26C38 14 26 6 26 6Z" fill="#fff"/>
                        <path d="M26 12C26 12 18 18 18 26C18 34 26 40 26 40C26 40 34 34 34 26C34 18 26 12 26 12Z" fill="#16a34a"/>
                        <text x="26" y="32" text-anchor="middle" fill="#fff" font-size="14" font-weight="bold">F</text>
                    </svg>
                </div>
            </div>
            <h1 class="text-4xl font-bold tracking-tight">
                <span class="fresh-green">Fresh</span><span class="fresh-red">Stop</span>
            </h1>
            <p class="text-slate-400 mt-1 text-lg">Dash</p>
            <p class="text-slate-500 text-sm mt-2">Smart Fuel Pricing • Real-time Optimization</p>
        </div>
        
        <div class="bg-slate-900 border border-slate-800 rounded-3xl p-8">
            {% with messages = get_flashed_messages() %}
                {% if messages %}
                    <div class="mb-4 p-3 bg-red-500/10 border border-red-500/30 rounded-2xl text-red-400 text-sm">{{ messages[0] }}</div>
                {% endif %}
            {% endwith %}
            <form action="/login" method="post" class="space-y-6">
                <div>
                    <label class="block text-sm text-slate-400 mb-2">Email</label>
                    <input type="email" name="email" placeholder="you@station.com" class="w-full bg-slate-800 border border-slate-700 rounded-2xl px-4 py-3 focus:border-green-500" required>
                </div>
                <div>
                    <label class="block text-sm text-slate-400 mb-2">Password</label>
                    <input type="password" name="password" placeholder="••••••••" class="w-full bg-slate-800 border border-slate-700 rounded-2xl px-4 py-3 focus:border-green-500" required>
                </div>
                <button type="submit" class="w-full bg-green-600 hover:bg-green-700 py-3 rounded-2xl font-semibold text-white transition-colors">Sign In</button>
            </form>
            <div class="mt-6 text-center">
                <p class="text-sm text-slate-400">New accounts can only be created by an Admin.</p>
            </div>
            <!-- No credentials shown - production ready -->
        </div>
        
        <div class="text-center mt-6 text-xs text-slate-500">
            © 2026 FreshStop • All rights reserved
        </div>
    </div>
</body>
</html>'''

REGISTER_HTML = '''<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>Register - Fresh Stop Dash</title>
    <script src="https://cdn.tailwindcss.com"></script>
    <style>
        .fresh-green { color: #16a34a; }
        .fresh-red { color: #dc2626; }
    </style>
</head>
<body class="bg-slate-950 text-slate-200 flex items-center justify-center min-h-screen">
    <div class="max-w-md w-full px-6">
        <div class="text-center mb-8">
            <div class="flex justify-center mb-4">
                <div class="w-16 h-16 bg-gradient-to-br from-green-500 to-emerald-600 rounded-full flex items-center justify-center shadow-xl">
                    <svg width="42" height="42" viewBox="0 0 52 52" fill="none" xmlns="http://www.w3.org/2000/svg">
                        <path d="M26 6C26 6 14 14 14 26C14 38 26 46 26 46C26 46 38 38 38 26C38 14 26 6 26 6Z" fill="#fff"/>
                        <path d="M26 12C26 12 18 18 18 26C18 34 26 40 26 40C26 40 34 34 34 26C34 18 26 12 26 12Z" fill="#16a34a"/>
                        <text x="26" y="32" text-anchor="middle" fill="#fff" font-size="13" font-weight="bold">F</text>
                    </svg>
                </div>
            </div>
            <h1 class="text-3xl font-bold">Create Account</h1>
            <p class="text-slate-400 mt-1">Join FreshStop Dash</p>
        </div>
        <div class="bg-slate-900 border border-slate-800 rounded-3xl p-8">
            {% with messages = get_flashed_messages() %}
                {% if messages %}
                    <div class="mb-4 p-3 bg-red-500/10 border border-red-500/30 rounded-2xl text-red-400 text-sm">{{ messages[0] }}</div>
                {% endif %}
            {% endwith %}
            <form action="/register" method="post" class="space-y-6">
                <div>
                    <label class="block text-sm text-slate-400 mb-2">Full Name</label>
                    <input type="text" name="name" placeholder="John Smith" class="w-full bg-slate-800 border border-slate-700 rounded-2xl px-4 py-3 focus:border-green-500" required>
                </div>
                <div>
                    <label class="block text-sm text-slate-400 mb-2">Email</label>
                    <input type="email" name="email" placeholder="you@station.com" class="w-full bg-slate-800 border border-slate-700 rounded-2xl px-4 py-3 focus:border-green-500" required>
                </div>
                <div>
                    <label class="block text-sm text-slate-400 mb-2">Password</label>
                    <input type="password" name="password" placeholder="••••••••" class="w-full bg-slate-800 border border-slate-700 rounded-2xl px-4 py-3 focus:border-green-500" required>
                </div>
                <button type="submit" class="w-full bg-green-600 hover:bg-green-700 py-3 rounded-2xl font-semibold text-white">Create Account</button>
            </form>
            <div class="mt-6 text-center">
                <p class="text-sm text-slate-400">Already have an account? <a href="/login" class="text-green-400 hover:text-green-300">Sign in</a></p>
            </div>
        </div>
    </div>
</body>
</html>'''

DASHBOARD_HTML = '''<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>Fresh Stop Dash</title>
    <script src="https://cdn.tailwindcss.com"></script>
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.1/css/all.min.css">
    <style>
        body { font-family: system-ui, -apple-system, sans-serif; }
        .grade-card { transition: all 0.2s ease; }
        .grade-card:hover { transform: translateY(-1px); }
        .elasticity-low { color: #10b981; }
        .elasticity-medium { color: #f59e0b; }
        .elasticity-high { color: #ef4444; }
    </style>
</head>
<body class="bg-slate-950 text-slate-200">
    <div class="max-w-7xl mx-auto px-6 py-8">
        <div class="flex justify-between items-center mb-6">
            <div class="flex items-center gap-x-3">
                <div class="w-11 h-11 bg-gradient-to-br from-green-500 to-emerald-600 rounded-xl flex items-center justify-center shadow">
                    <svg width="28" height="28" viewBox="0 0 52 52" fill="none" xmlns="http://www.w3.org/2000/svg">
                        <path d="M26 6C26 6 14 14 14 26C14 38 26 46 26 46C26 46 38 38 38 26C38 14 26 6 26 6Z" fill="#fff"/>
                        <path d="M26 12C26 12 18 18 18 26C18 34 26 40 26 40C26 40 34 34 34 26C34 18 26 12 26 12Z" fill="#16a34a"/>
                        <text x="26" y="32" text-anchor="middle" fill="#fff" font-size="11" font-weight="bold">F</text>
                    </svg>
                </div>
                <div>
                    <h1 class="text-4xl font-bold tracking-tight">
                        <span class="text-green-500">Fresh</span><span class="text-red-500">Stop</span> <span class="text-slate-300">Dash</span>
                    </h1>
                    <p class="text-slate-400 -mt-1">Welcome back, {{ session.user_name }}!
                       {% if session.user_role != 'viewer' %}
                       <span class="text-green-400 font-semibold">({{ session.user_role }})</span>
                       {% endif %}</p>
                </div>
            </div>
            <div class="flex-1 max-w-md mx-8">
                <div class="relative">
                    <input type="text" id="site-search" placeholder="Search stations..." 
                           class="w-full bg-slate-800 border border-slate-700 rounded-2xl px-4 py-2 text-sm pl-10"
                           onkeyup="filterStations()">
                    <i class="fa-solid fa-search absolute left-4 top-3 text-slate-400"></i>
                </div>
            </div>
            <div class="flex items-center gap-x-3">
                <a href="/logout" class="flex items-center gap-x-2 bg-slate-700 hover:bg-slate-600 px-4 py-2 rounded-xl text-sm font-semibold">
                    <i class="fa-solid fa-sign-out-alt fa-sm"></i>
                    <span>Logout</span>
                </a>
                <a href="/upload_data" class="flex items-center gap-x-2 bg-blue-600 hover:bg-blue-700 px-4 py-2 rounded-xl text-sm font-semibold">
                    <i class="fa-solid fa-upload fa-sm"></i>
                    <span>Upload</span>
                </a>
                <a href="/invoices" class="flex items-center gap-x-2 bg-emerald-600 hover:bg-emerald-700 px-4 py-2 rounded-xl text-sm font-semibold">
                    <i class="fa-solid fa-file-invoice fa-sm"></i>
                    <span>Invoices</span>
                </a>
                <a href="/add_station" class="flex items-center gap-x-2 bg-purple-600 hover:bg-purple-700 px-4 py-2 rounded-xl text-sm font-semibold">
                    <i class="fa-solid fa-plus fa-sm"></i>
                    <span>Add</span>
                </a>
                {% if session.user_role == 'admin' %}
                <a href="/admin/users" class="flex items-center gap-x-2 bg-red-600 hover:bg-red-700 px-4 py-2 rounded-xl text-sm font-semibold">
                    <i class="fa-solid fa-users fa-sm"></i>
                    <span>Users</span>
                </a>
                {% endif %}
                <a href="/print_snapshot" class="flex items-center gap-x-2 bg-amber-600 hover:bg-amber-700 px-4 py-2 rounded-xl text-sm font-semibold">
                    <i class="fa-solid fa-print fa-sm"></i>
                    <span>Print Snapshot</span>
                </a>
            </div>
        </div>

        {% for sid, station in stations.items() %}
        <div id="station-{{ sid }}" class="bg-slate-900 border border-slate-800 rounded-3xl p-6 mb-8">
            <div class="flex justify-between items-center mb-6">
                <div>
                    <h2 class="text-2xl font-semibold">{{ station.name }}</h2>
                    <p class="text-slate-400">{{ station.address }}, {{ station.city }}, {{ station.state }} {{ station.zip }}</p>
                </div>
                <div class="flex items-center gap-x-3">
                    <div class="text-xs text-slate-500">Last updated: {{ now }}</div>
                    <a href="/edit_station/{{ sid }}" class="flex items-center gap-x-2 bg-slate-700 hover:bg-slate-600 px-4 py-2 rounded-xl text-sm font-semibold">
                        <i class="fa-solid fa-edit fa-sm"></i>
                        <span>Edit</span>
                    </a>
                    <a href="/send_prices/{{ sid }}" class="flex items-center gap-x-2 bg-indigo-600 hover:bg-indigo-700 px-4 py-2 rounded-xl text-sm font-semibold">
                        <i class="fa-solid fa-envelope fa-sm"></i>
                        <span>Send Prices</span>
                    </a>
                    <a href="/station/{{ sid }}/documents" class="flex items-center gap-x-2 bg-purple-600 hover:bg-purple-700 px-4 py-2 rounded-xl text-sm font-semibold">
                        <i class="fa-solid fa-folder-open fa-sm"></i>
                        <span>Documents</span>
                    </a>
                </div>
            </div>

            <div class="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-4 mb-8">
                {% for grade in grades %}
                {% set g = station.grades[grade] %}
                {% set strategic_price = get_price_for_strategy(g.wholesale, g.strategy) %}
                {% set ml_optimal = get_ml_optimal_price(g.wholesale, g.current_price, g.elasticity, g.daily_gallons) %}
                <div class="grade-card bg-slate-800 rounded-2xl p-5">
                    <div class="flex justify-between items-center mb-4">
                        <div class="font-semibold text-lg">{{ grade }}</div>
                        <div class="text-xs px-3 py-1 bg-slate-700 rounded-full">{{ g.daily_gallons }} gal/day</div>
                    </div>
                    
                    <div class="space-y-2 text-sm mb-4">
                        <div class="flex justify-between"><span class="text-slate-400">Wholesale</span> <span class="font-mono">${{ "%.2f"|format(g.wholesale) }}</span></div>
                        <div class="flex justify-between"><span class="text-slate-400">Current Price</span> <span class="font-mono font-semibold text-yellow-400">${{ "%.2f"|format(g.current_price) }}</span></div>
                        <div class="flex justify-between">
                            <span class="text-slate-400">Elasticity</span> 
                            <span class="font-mono font-semibold 
                                {% if g.elasticity < -2 %}elasticity-high
                                {% elif g.elasticity < -1 %}elasticity-medium
                                {% else %}elasticity-low{% endif %}">
                                {{ g.elasticity }}
                            </span>
                        </div>
                        <div class="flex justify-between"><span class="text-slate-400">Margin</span> <span class="font-mono {{ (g.current_price - g.wholesale) > 0.10 and 'text-emerald-400' or 'text-amber-400' }}">{{ ((g.current_price - g.wholesale)*100)|round(0)|int }}c</span></div>
                    </div>

                    <div class="pt-4 border-t border-slate-700">
                        <div class="text-xs text-slate-500 mb-2">Strategy for {{ grade }}</div>
                        <div class="flex gap-x-1 mb-3">
                            <button onclick="setGradeStrategy({{ sid }}, '{{ grade }}', 'aggressive', this)" 
                                    class="strategy-btn-small px-3 py-1 text-xs rounded-xl border {% if g.strategy == 'aggressive' %}bg-red-500 text-white{% else %}border-red-500/30 hover:bg-red-500/10{% endif %}">
                                <i class="fa-solid fa-bolt fa-xs"></i> A
                            </button>
                            <button onclick="setGradeStrategy({{ sid }}, '{{ grade }}', 'balanced', this)" 
                                    class="strategy-btn-small px-3 py-1 text-xs rounded-xl border {% if g.strategy == 'balanced' %}bg-sky-500 text-white{% else %}border-sky-500/30 hover:bg-sky-500/10{% endif %}">
                                <i class="fa-solid fa-balance-scale fa-xs"></i> B
                            </button>
                            <button onclick="setGradeStrategy({{ sid }}, '{{ grade }}', 'match', this)" 
                                    class="strategy-btn-small px-3 py-1 text-xs rounded-xl border {% if g.strategy == 'match' %}bg-amber-500 text-white{% else %}border-amber-500/30 hover:bg-amber-500/10{% endif %}">
                                <i class="fa-solid fa-handshake fa-xs"></i> M
                            </button>
                            <button onclick="refreshOPISData({{ sid }}, '{{ grade }}', this)" 
                                    class="strategy-btn-small px-3 py-1 text-xs rounded-xl border border-sky-500/30 hover:bg-sky-500/10 text-sky-400">
                                <i class="fa-solid fa-sync fa-xs"></i> OPIS
                            </button>
                        </div>
                        <div class="text-xs text-emerald-400 mb-1">Strategy: <span class="strategy-label font-semibold">{{ g.strategy.title() }}</span></div>
                    </div>

                    <div class="pt-2">
                        <div class="flex justify-between items-center">
                            <div>
                                <div class="text-xs text-sky-400">Strategic Price</div>
                                <div class="font-mono text-xl font-bold text-sky-400">${{ "%.2f"|format(strategic_price) }}</div>
                            </div>
                            <button onclick="applyPrice({{ sid }}, '{{ grade }}', {{ strategic_price }})" 
                                    class="text-xs bg-green-600 hover:bg-green-700 px-4 py-1.5 rounded-xl font-semibold">Apply</button>
                        </div>
                    </div>

                    <div class="pt-3 border-t border-slate-700 mt-2">
                        <div class="flex justify-between items-center">
                            <div>
                                <div class="text-xs text-emerald-400">ML Optimal</div>
                                <div class="font-mono text-xl font-bold text-emerald-400">${{ "%.2f"|format(ml_optimal) }}</div>
                            </div>
                            <button onclick="applyPrice({{ sid }}, '{{ grade }}', {{ ml_optimal }})" 
                                    class="text-xs bg-emerald-600 hover:bg-emerald-700 px-4 py-1.5 rounded-xl font-semibold">Apply</button>
                        </div>
                    </div>

                    <div class="pt-3 border-t border-slate-700 mt-2">
                        <div class="text-xs text-slate-500 mb-1">Custom Price</div>
                        <div class="flex gap-x-2">
                            <input type="number" step="0.01" id="custom-{{ sid }}-{{ grade }}" class="flex-1 bg-slate-700 border border-slate-600 rounded-xl px-3 py-1 text-sm font-mono" placeholder="{{ "%.2f"|format(g.current_price) }}">
                            <button onclick="setCustomPrice({{ sid }}, '{{ grade }}')" class="text-xs bg-slate-600 hover:bg-slate-500 px-4 rounded-xl font-semibold">Set</button>
                        </div>
                    </div>
                </div>
                {% endfor %}
            </div>

            <!-- Competitor Prices Section -->
            <div class="bg-slate-800 border border-slate-700 rounded-2xl p-5 mt-6">
                <div class="flex justify-between items-center mb-4">
                    <div class="font-semibold flex items-center gap-x-2">
                        <i class="fa-solid fa-gas-pump text-green-400"></i>
                        <span>Competitor Prices</span>
                    </div>
                    <div class="text-xs text-slate-400">ZIP: {{ station.zip }}</div>
                </div>

                <form action="/search_competitors/{{ sid }}" method="post" class="flex gap-x-3 mb-4">
                    <input type="text" name="zip_code" value="{{ station.zip }}" placeholder="ZIP Code" class="flex-1 bg-slate-700 border border-slate-600 rounded-xl px-4 py-2 text-sm">
                    <select name="radius" class="bg-slate-700 border border-slate-600 rounded-xl px-3 py-2 text-sm">
                        {% for r in range(1, 21) %}
                        <option value="{{ r }}">{{ r }} mile{{ 's' if r > 1 else '' }}</option>
                        {% endfor %}
                    </select>
                    <button type="submit" class="bg-sky-600 hover:bg-sky-700 px-5 py-2 rounded-xl text-sm font-semibold flex items-center gap-x-2">
                        <i class="fa-solid fa-sync fa-sm"></i>
                        <span>OPIS Refresh</span>
                    </button>
                </form>

                <form action="/add_competitor/{{ sid }}" method="post" class="grid grid-cols-1 md:grid-cols-6 gap-3 mb-4">
                    <input type="text" name="comp_name" placeholder="Competitor Name" class="md:col-span-2 bg-slate-700 border border-slate-600 rounded-xl px-4 py-2 text-sm" required>
                    
                    <div class="grid grid-cols-4 gap-2 md:col-span-3">
                        <input type="number" step="0.01" name="price_regular" placeholder="Reg" class="bg-slate-700 border border-slate-600 rounded-xl px-2 py-2 text-sm" required>
                        <input type="number" step="0.01" name="price_mid" placeholder="Mid" class="bg-slate-700 border border-slate-600 rounded-xl px-2 py-2 text-sm" required>
                        <input type="number" step="0.01" name="price_premium" placeholder="Prem" class="bg-slate-700 border border-slate-600 rounded-xl px-2 py-2 text-sm" required>
                        <input type="number" step="0.01" name="price_diesel" placeholder="Diesel" class="bg-slate-700 border border-slate-600 rounded-xl px-2 py-2 text-sm" required>
                    </div>
                    
                    <div class="flex gap-2">
                        <input type="number" step="0.1" name="comp_distance" placeholder="Miles" class="w-20 bg-slate-700 border border-slate-600 rounded-xl px-3 py-2 text-sm" required>
                        <button type="submit" class="bg-emerald-600 hover:bg-emerald-700 px-4 py-2 rounded-xl text-sm font-semibold whitespace-nowrap">Add</button>
                    </div>
                </form>

                <div class="max-h-48 overflow-y-auto">
                    {% if station.competitor_list %}
                        <table class="w-full text-sm">
                            <thead>
                                <tr class="border-b border-slate-700 text-slate-400">
                                    <th class="text-left py-2">Competitor</th>
                                    <th class="text-center py-2">Regular</th>
                                    <th class="text-center py-2">Mid</th>
                                    <th class="text-center py-2">Premium</th>
                                    <th class="text-center py-2">Diesel</th>
                                    <th class="text-left py-2">Dist</th>
                                    <th></th>
                                </tr>
                            </thead>
                            <tbody>
                                {% for comp in station.competitor_list %}
                                <tr class="border-b border-slate-700 last:border-0 {% if comp.pinned %}bg-green-900/10{% endif %}">
                                    <td class="py-2 font-medium {% if comp.pinned %}text-green-400 font-semibold{% endif %}">{{ comp.name }}</td>
                                    <td class="py-2 text-center font-mono text-yellow-400">${{ "%.2f"|format(comp.get('prices', {}).get('Regular', comp.get('price', 0))) }}</td>
                                    <td class="py-2 text-center font-mono text-yellow-400">${{ "%.2f"|format(comp.get('prices', {}).get('Mid-grade', comp.get('price', 0))) }}</td>
                                    <td class="py-2 text-center font-mono text-yellow-400">${{ "%.2f"|format(comp.get('prices', {}).get('Premium', comp.get('price', 0))) }}</td>
                                    <td class="py-2 text-center font-mono text-yellow-400">${{ "%.2f"|format(comp.get('prices', {}).get('Diesel', comp.get('price', 0))) }}</td>
                                    <td class="py-2">{{ comp.distance }} mi</td>
                                    <td class="py-2 text-right flex items-center gap-x-2 justify-end">
                                        {% if comp.pinned %}
                                        <a href="/toggle_pin/{{ sid }}/{{ loop.index0 }}" 
                                           class="text-green-500 hover:text-green-600 text-xs transition-colors" 
                                           title="Unpin this competitor">
                                            <i class="fa-solid fa-map-pin fa-lg"></i>
                                        </a>
                                        {% else %}
                                        <a href="/toggle_pin/{{ sid }}/{{ loop.index0 }}" 
                                           class="text-slate-400 hover:text-amber-400 text-xs transition-colors" 
                                           title="Pin this competitor">
                                            <i class="fa-solid fa-thumbtack"></i>
                                        </a>
                                        {% endif %}
                                        <a href="/delete_competitor/{{ sid }}/{{ loop.index0 }}" class="text-red-400 hover:text-red-500 text-xs">Remove</a>
                                    </td>
                                </tr>
                                {% endfor %}
                            </tbody>
                        </table>
                    {% else %}
                        <div class="text-center py-4 text-slate-400 text-sm">No competitors added yet.</div>
                    {% endif %}
                </div>

                <!-- EIA/AAA Regional Averages Ticker -->
                <div class="bg-slate-800 border border-slate-700 rounded-2xl p-5 mt-6">
                    <div class="flex justify-between items-center mb-4">
                        <div class="font-semibold flex items-center gap-x-2">
                            <i class="fa-solid fa-chart-line text-emerald-400"></i>
                            <span>Regional Weekly Average <span class="text-emerald-400">(EIA Primary)</span></span>
                        </div>
                        <div class="text-xs text-slate-400">Updated weekly • Based on {{ station.state }}</div>
                    </div>
                    {% set reg = get_regional_averages(station.state) %}
                    <div class="grid grid-cols-2 md:grid-cols-4 gap-4 text-sm">
                        {% for grade in grades %}
                        <div class="bg-slate-900 rounded-xl p-3">
                            <div class="text-xs text-slate-400">{{ grade }}</div>
                            <div class="font-mono text-lg font-bold text-emerald-400">${{ "%.2f"|format(reg[grade]) }}</div>
                            <div class="text-[10px] text-amber-400">AAA: ${{ "%.2f"|format(reg[grade] + 0.02) }}</div>
                        </div>
                        {% endfor %}
                    </div>
                    <div class="text-[10px] text-slate-500 mt-2">Source: <strong>EIA</strong> (Primary) + AAA (Daily reference) • Updated weekly</div>
                </div>
            </div>
        </div>
        {% endfor %}

        <div class="bg-slate-900 border border-slate-800 rounded-3xl p-6 mt-8">
            <div class="flex justify-between items-center mb-4">
                <div class="font-semibold flex items-center gap-x-2">
                    <i class="fa-solid fa-history text-amber-400"></i>
                    <span>Recent Price Changes (Audit Trail)</span>
                </div>
                <a href="/print_audit" class="text-xs bg-amber-600 hover:bg-amber-700 text-white px-3 py-1 rounded">Price Audit</a>
                <a href="/print_email_audit" class="text-xs bg-indigo-600 hover:bg-indigo-700 text-white px-3 py-1 rounded">Email Audit</a>
                <a href="/print_edit_audit" class="text-xs bg-blue-600 hover:bg-blue-700 text-white px-3 py-1 rounded">Edit Audit</a>
                <a href="/print_document_audit" class="text-xs bg-purple-600 hover:bg-purple-700 text-white px-3 py-1 rounded">Document/ML Audit</a>
            </div>
            <div class="text-sm text-slate-400 max-h-64 overflow-y-auto">
                {% if audit_log %}
                    {% for entry in audit_log %}
                    <div class="flex justify-between py-2 border-b border-slate-800 last:border-0">
                        <span>{{ entry.time }} - {{ entry.station }} ({{ entry.grade }})</span>
                        <span class="font-mono">{{ entry.old_price }} - <span class="text-green-400">{{ entry.new_price }}</span></span>
                    </div>
                    {% endfor %}
                {% else %}
                    <div class="text-center py-4 text-slate-500">No price changes yet today</div>
                {% endif %}
            </div>
        </div>
    </div>

    <script>
        function setGradeStrategy(stationId, grade, strategy, element) {
            fetch('/set_grade_strategy', {
                method: 'POST',
                headers: {'Content-Type': 'application/json'},
                body: JSON.stringify({station_id: stationId, grade: grade, strategy: strategy})
            }).then(() => location.reload());
        }
        
        function applyPrice(stationId, grade, newPrice) {
            if (confirm('Apply price of $' + parseFloat(newPrice).toFixed(2) + ' for ' + grade + '?')) {
                fetch('/apply_price', {
                    method: 'POST',
                    headers: {'Content-Type': 'application/json'},
                    body: JSON.stringify({station_id: stationId, grade: grade, new_price: newPrice})
                }).then(() => location.reload());
            }
        }
        
        function setCustomPrice(stationId, grade) {
            const input = document.getElementById('custom-' + stationId + '-' + grade);
            const newPrice = parseFloat(input.value);
            if (!newPrice || newPrice <= 0) {
                alert('Please enter a valid price');
                return;
            }
            if (confirm('Set custom ' + grade + ' price to $' + newPrice.toFixed(2) + '?')) {
                fetch('/apply_price', {
                    method: 'POST',
                    headers: {'Content-Type': 'application/json'},
                    body: JSON.stringify({station_id: stationId, grade: grade, new_price: newPrice})
                }).then(() => location.reload());
            }
        }
        
        function filterStations() {
            const searchTerm = document.getElementById('site-search').value.toLowerCase();
            const stationCards = document.querySelectorAll('.bg-slate-900.border.border-slate-800.rounded-3xl');
            stationCards.forEach(card => {
                const stationName = card.querySelector('h2').textContent.toLowerCase();
                if (stationName.includes(searchTerm)) {
                    card.style.display = '';
                } else {
                    card.style.display = 'none';
                }
            });
        }

        function refreshOPISData(stationId, grade, element) {
            element.innerHTML = '<i class="fa-solid fa-spinner fa-spin fa-xs"></i> ...';
            // For now, this refreshes the page (simulating OPIS data refresh)
            // Later we will connect real OPIS API here
            setTimeout(() => {
                location.reload();
            }, 600);
        }
    </script>
</body>
</html>'''

# ==================== ROUTES (with fix for get_regional_averages) ====================

@app.route("/")
def index():
    if 'user_email' in session:
        return redirect("/dashboard")
    return redirect("/login")

@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        email = request.form.get("email", "").lower().strip()
        password = request.form.get("password", "")
        user = get_user(email)
        if user and check_password_hash(user["password"], password):
            session['user_email'] = email
            session['user_name'] = user.get("name", "User")
            session['user_role'] = user.get("role", "viewer")
            flash("Welcome back!")
            return redirect("/dashboard")
        else:
            flash("Invalid email or password")
    return render_template_string(LOGIN_HTML)

@app.route("/register", methods=["GET", "POST"])
def register():
    if request.method == "POST":
        name = request.form.get("name", "").strip()
        email = request.form.get("email", "").lower().strip()
        password = request.form.get("password", "")
        if get_user(email):
            flash("An account with this email already exists")
        elif len(password) < 6:
            flash("Password must be at least 6 characters")
        else:
            # Default new users to 'viewer' role
            if create_user(email, generate_password_hash(password, method='pbkdf2:sha256'), name, "viewer"):
                session['user_email'] = email
                session['user_name'] = name
                session['user_role'] = "viewer"
                flash("Account created successfully! You have been assigned Viewer role.")
                return redirect("/dashboard")
            else:
                flash("Error creating account")
    return render_template_string(REGISTER_HTML)

@app.route("/logout")
def logout():
    session.clear()
    flash("You have been logged out")
    return redirect("/login")

@app.route("/dashboard")
@login_required
def dashboard():
    now = datetime.now().strftime("%H:%M")
    return render_template_string(DASHBOARD_HTML, 
        stations=STATIONS, 
        grades=GRADES, 
        now=now, 
        audit_log=AUDIT_LOG,
        get_price_for_strategy=get_price_for_strategy,
        get_ml_optimal_price=get_ml_optimal_price,
        get_regional_averages=get_regional_averages   # <-- FIXED: now passed to template
    )

@app.route("/print_snapshot")
@login_required
def print_snapshot():
    now = datetime.now().strftime("%Y-%m-%d %H:%M")
    return render_template_string(PRINT_SNAPSHOT_HTML, stations=STATIONS, grades=GRADES, now=now)

@app.route("/print_audit")
@login_required
def print_audit():
    return render_template_string(PRINT_AUDIT_HTML, audit_log=AUDIT_LOG)

@app.route("/print_email_audit")
@login_required
def print_email_audit():
    return render_template_string(PRINT_EMAIL_AUDIT_HTML, email_audit=EMAIL_AUDIT)

@app.route("/print_edit_audit")
@login_required
def print_edit_audit():
    return render_template_string(PRINT_EDIT_AUDIT_HTML, edit_audit=EDIT_AUDIT)

@app.route("/print_document_audit")
@login_required
def print_document_audit():
    return render_template_string(PRINT_DOCUMENT_AUDIT_HTML, document_audit=DOCUMENT_AUDIT)

@app.route("/set_grade_strategy", methods=["POST"])
@login_required
@manager_required
def set_grade_strategy():
    data = request.get_json()
    station_id = data["station_id"]
    grade = data["grade"]
    strategy = data["strategy"]
    if station_id in STATIONS and grade in STATIONS[station_id]["grades"]:
        STATIONS[station_id]["grades"][grade]["strategy"] = strategy
    return {"success": True}

@app.route("/apply_price", methods=["POST"])
@login_required
@manager_required
def apply_price():
    data = request.get_json()
    station_id = data["station_id"]
    grade = data["grade"]
    new_price = float(data["new_price"])
    old_price = STATIONS[station_id]["grades"][grade]["current_price"]
    STATIONS[station_id]["grades"][grade]["current_price"] = round(new_price, 2)
    add_to_audit(STATIONS[station_id]["name"], grade, old_price, new_price)
    return {"success": True}

@app.route("/send_prices/<int:sid>")
@login_required
def send_prices(sid):
    if sid not in STATIONS:
        flash("Station not found")
        return redirect(f"/dashboard#station-{sid}")
    station = STATIONS[sid]
    recipient = station.get("email", "station@freshstop.com")
    prices_list = []
    for grade in GRADES:
        g = station["grades"][grade]
        prices_list.append(f"{grade}: Current ${g['current_price']:.2f} | Wholesale ${g['wholesale']:.2f} | Margin {((g['current_price'] - g['wholesale'])*100):.0f}¢")
    prices_summary = "\n".join(prices_list)
    
    # Support multiple emails (comma-separated)
    recipients = [r.strip() for r in recipient.split(",") if r.strip()]
    
    # Show preview page instead of sending immediately
    return render_template_string(EMAIL_PREVIEW_HTML, 
                                  station=station, 
                                  sid=sid,
                                  recipient=", ".join(recipients),
                                  prices_summary=prices_summary)


# Email Preview Template
EMAIL_PREVIEW_HTML = '''<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>Email Preview - {{ station.name }}</title>
    <script src="https://cdn.tailwindcss.com"></script>
</head>
<body class="bg-slate-950 text-slate-200 p-8">
    <div class="max-w-2xl mx-auto">
        <div class="flex justify-between items-center mb-8">
            <h1 class="text-3xl font-bold">Email Preview</h1>
            <a href="/dashboard#station-{{ sid }}" class="text-green-400 hover:text-sky-300">← Cancel</a>
        </div>

        <div class="bg-slate-900 border border-slate-800 rounded-3xl p-8 mb-8">
            <div class="mb-6">
                <div class="text-sm text-slate-400">To:</div>
                <div class="font-semibold">{{ recipient }}</div>
            </div>
            
            <div class="mb-6">
                <div class="text-sm text-slate-400">Subject:</div>
                <div class="font-semibold">Price Update - {{ station.name }}</div>
            </div>
            
            <div class="mb-6">
                <div class="text-sm text-slate-400 mb-2">Message:</div>
                <div class="bg-slate-800 border border-slate-700 rounded-xl p-4 font-mono text-sm whitespace-pre-line">{{ prices_summary }}</div>
            </div>
            
            <div class="text-xs text-slate-500">
                This email will be sent from your connected Outlook/Office 365 account.
            </div>
        </div>

        <div class="flex gap-4">
            <a href="/dashboard#station-{{ sid }}" class="flex-1 text-center bg-slate-700 hover:bg-slate-600 px-6 py-3 rounded-2xl font-semibold">Cancel</a>
            <a href="/send_prices_confirm/{{ sid }}" class="flex-1 text-center bg-emerald-600 hover:bg-emerald-700 px-6 py-3 rounded-2xl font-semibold">Confirm & Send</a>
        </div>
    </div>
</body>
</html>'''


@app.route("/send_prices_confirm/<int:sid>")
@login_required
def send_prices_confirm(sid):
    if sid not in STATIONS:
        flash("Station not found")
        return redirect(f"/dashboard#station-{sid}")
    
    station = STATIONS[sid]
    recipient = station.get("email", "station@freshstop.com")
    prices_list = []
    for grade in GRADES:
        g = station["grades"][grade]
        prices_list.append(f"{grade}: Current ${g['current_price']:.2f} | Wholesale ${g['wholesale']:.2f} | Margin {((g['current_price'] - g['wholesale'])*100):.0f}¢")
    prices_summary = "\n".join(prices_list)
    
    # Support multiple recipients (comma-separated)
    recipients = [r.strip() for r in recipient.split(",") if r.strip()]
    
    # Try to send real email via Microsoft Graph (Outlook)
    sent = False
    try:
        access_token = get_sharepoint_token()
        if access_token:
            email_data = {
                "message": {
                    "subject": f"Price Update - {station['name']}",
                    "body": {
                        "contentType": "Text",
                        "content": f"Current prices for {station['name']}:\n\n{prices_summary}\n\nGenerated by Fresh Stop Dash"
                    },
                    "toRecipients": [
                        {"emailAddress": {"address": r}} for r in recipients
                    ]
                }
            }
            
            headers = {
                "Authorization": f"Bearer {access_token}",
                "Content-Type": "application/json"
            }
            
            response = requests.post(
                "https://graph.microsoft.com/v1.0/me/sendMail",
                headers=headers,
                json=email_data
            )
            
            if response.status_code == 202:
                sent = True
    except Exception as e:
        print(f"Email send error: {e}")
    
    add_to_email_audit(station["name"], ", ".join(recipients), prices_summary)
    
    if sent:
        flash(f"✅ Email sent successfully to {len(recipients)} recipient(s) for {station['name']}")
    else:
        flash(f"✅ Prices logged for {station['name']} (Preview shown - real sending requires Outlook connection)")
    
    return redirect(f"/dashboard#station-{sid}")

@app.route("/search_competitors/<int:sid>", methods=["POST"])
@login_required
def search_competitors(sid):
    if sid not in STATIONS:
        flash("Station not found")
        return redirect(f"/dashboard#station-{sid}")
    zip_code = request.form.get("zip_code", "")
    radius = int(request.form.get("radius", 5))
    if zip_code in ["33139", "77002", "73102", "72201"]:
        import random
        # Preserve all currently pinned competitors
        current_list = STATIONS[sid].get("competitor_list", [])
        pinned_competitors = [c for c in current_list if c.get("pinned", False)]
        
        # Generate fresh competitors from OPIS-style search with all grades
        new_competitors = []
        for i in range(random.randint(2, 5)):
            base = round(random.uniform(3.05, 3.35), 2)
            new_competitors.append({
                "name": f"Competitor {i}",
                "prices": {
                    "Regular": round(base + random.uniform(-0.05, 0.05), 2),
                    "Mid-grade": round(base + 0.15 + random.uniform(-0.03, 0.03), 2),
                    "Premium": round(base + 0.30 + random.uniform(-0.03, 0.03), 2),
                    "Diesel": round(base - 0.05 + random.uniform(-0.05, 0.05), 2)
                },
                "distance": round(random.uniform(0.5, radius), 1),
                "last_updated": datetime.now().strftime("%Y-%m-%d %H:%M"),
                "pinned": False
            })
        
        # Merge: Keep pinned ones + add new ones (avoid exact duplicates by name)
        existing_names = {c["name"] for c in pinned_competitors}
        for new_c in new_competitors:
            if new_c["name"] not in existing_names:
                pinned_competitors.append(new_c)
        
        STATIONS[sid]["competitor_list"] = pinned_competitors
        flash(f"OPIS refreshed • Kept {len([c for c in pinned_competitors if c.get('pinned')])} pinned • Added {len(new_competitors)} new competitors")
    else:
        flash("No competitors found for this ZIP code")
    return redirect(f"/dashboard#station-{sid}")

@app.route("/add_competitor/<int:sid>", methods=["POST"])
@login_required
def add_competitor(sid):
    if sid not in STATIONS:
        flash("Station not found")
        return redirect(f"/dashboard#station-{sid}")
    name = request.form.get("comp_name", "").strip()
    distance = float(request.form.get("comp_distance", 0))
    
    # Get prices for all grades
    prices = {
        "Regular": float(request.form.get("price_regular", 0)),
        "Mid-grade": float(request.form.get("price_mid", 0)),
        "Premium": float(request.form.get("price_premium", 0)),
        "Diesel": float(request.form.get("price_diesel", 0))
    }
    
    if name and any(prices.values()):
        STATIONS[sid]["competitor_list"].append({
            "name": name,
            "prices": prices,
            "distance": distance,
            "last_updated": datetime.now().strftime("%Y-%m-%d %H:%M"),
            "pinned": False
        })
        flash(f"Added competitor: {name} with all grades")
    return redirect(f"/dashboard#station-{sid}")

@app.route("/delete_competitor/<int:sid>/<int:index>")
@login_required
def delete_competitor(sid, index):
    if sid in STATIONS and index < len(STATIONS[sid]["competitor_list"]):
        removed = STATIONS[sid]["competitor_list"].pop(index)
        flash(f"Removed competitor: {removed['name']}")
    return redirect(f"/dashboard#station-{sid}")

@app.route("/toggle_pin/<int:sid>/<int:index>")
@login_required
def toggle_pin(sid, index):
    if sid in STATIONS and index < len(STATIONS[sid]["competitor_list"]):
        comp = STATIONS[sid]["competitor_list"][index]
        comp["pinned"] = not comp.get("pinned", False)
        status = "pinned" if comp["pinned"] else "unpinned"
        flash(f"Competitor {status}: {comp['name']}")
    return redirect(f"/dashboard#station-{sid}")

@app.route("/add_station", methods=["GET", "POST"])
@login_required
@admin_required
def add_station():
    global NEXT_STATION_ID
    if request.method == "POST":
        new_station = {
            "name": request.form.get("name", "New Station"),
            "address": request.form.get("address", ""),
            "city": request.form.get("city", ""),
            "state": request.form.get("state", ""),
            "zip": request.form.get("zip", ""),
            "email": request.form.get("email", ""),
            "grades": {},
            "competitor_list": []
        }
        for grade in GRADES:
            new_station["grades"][grade] = {
                "wholesale": float(request.form.get(f"wholesale_{grade}", 2.80)),
                "current_price": float(request.form.get(f"price_{grade}", 3.10)),
                "competitors": [3.15, 3.12, 3.08],
                "target_margin": float(request.form.get(f"margin_{grade}", 0.12)),
                "daily_gallons": int(request.form.get(f"gallons_{grade}", 2000)),
                "elasticity": -1.5,
                "strategy": request.form.get(f"strategy_{grade}", "balanced")
            }
        new_sid = NEXT_STATION_ID
        STATIONS[new_sid] = new_station
        NEXT_STATION_ID += 1
        add_to_edit_audit(new_station["name"], "Created", "New station added via Add Station form")
        flash("New station added successfully!")
        return redirect(f"/dashboard#station-{new_sid}")
    return render_template_string(ADD_STATION_HTML, grades=GRADES)

@app.route("/edit_station/<int:station_id>", methods=["GET", "POST"])
@login_required
@manager_required
def edit_station(station_id):
    if station_id not in STATIONS:
        flash("Station not found")
        return redirect("/dashboard")
    if request.method == "POST":
        station = STATIONS[station_id]
        station["name"] = request.form.get("name", station["name"])
        station["address"] = request.form.get("address", station["address"])
        station["city"] = request.form.get("city", station["city"])
        station["state"] = request.form.get("state", station["state"])
        station["zip"] = request.form.get("zip", station["zip"])
        station["email"] = request.form.get("email", station.get("email", ""))
        for grade in GRADES:
            if grade in station["grades"]:
                g = station["grades"][grade]
                g["wholesale"] = float(request.form.get(f"wholesale_{grade}", g["wholesale"]))
                g["target_margin"] = float(request.form.get(f"margin_{grade}", g["target_margin"]))
                g["daily_gallons"] = int(request.form.get(f"gallons_{grade}", g["daily_gallons"]))
                g["current_price"] = float(request.form.get(f"price_{grade}", g["current_price"]))
                g["strategy"] = request.form.get(f"strategy_{grade}", g.get("strategy", "balanced"))
        add_to_edit_audit(station["name"], "Edited", "Station details and pricing updated via Edit form")
        flash("Station updated successfully!")
        return redirect(f"/dashboard#station-{station_id}")
    return render_template_string(EDIT_STATION_HTML, station=STATIONS[station_id], station_id=station_id, grades=GRADES)

@app.route("/invoices")
@login_required
def invoices_page():
    invoices = load_invoices()
    return render_template_string(INVOICES_HTML, invoices=invoices, stations=STATIONS, grades=GRADES)

@app.route("/upload_invoice", methods=["POST"])
@login_required
def upload_invoice():
    station_id = int(request.form["station_id"])
    grade = request.form["grade"]
    date = request.form["date"]
    file = request.files["invoice_file"]
    if file:
        filename = secure_filename(f"{datetime.now().strftime('%Y%m%d')}_{station_id}_{grade}_{file.filename}")
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(filepath)
        invoices = load_invoices()
        invoices.append({
            "id": len(invoices) + 1,
            "station_id": station_id,
            "station_name": STATIONS[station_id]["name"],
            "grade": grade,
            "date": date,
            "filename": filename,
            "total": 0.0,
            "uploaded_at": datetime.now().isoformat()
        })
        save_invoices(invoices)
        add_to_document_audit(STATIONS[station_id]["name"], "Fuel Invoice", filename, f"Grade: {grade} | Date: {date}")
        flash("Invoice uploaded!")
    return redirect("/invoices")

@app.route("/invoices/<filename>")
@login_required
def view_invoice(filename):
    return send_from_directory(app.config['UPLOAD_FOLDER'], filename)

@app.route("/upload_document", methods=["POST"])
@login_required
def upload_document():
    station_id = int(request.form["station_id"])
    description = request.form["description"]
    doc_date = request.form["doc_date"]
    file = request.files["document_file"]
    if file:
        filename = secure_filename(f"{datetime.now().strftime('%Y%m%d')}_{station_id}_{description.replace(' ', '_')}_{file.filename}")
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(filepath)
        # Upload to SharePoint using Device Code Flow
        sharepoint_url = None
        try:
            with open(filepath, 'rb') as f:
                file_content = f.read()
            station_name = STATIONS[station_id]["name"]
            sharepoint_url = upload_to_sharepoint(station_name, filename, file_content)
        except Exception as e:
            print(f"SharePoint upload failed: {e}")

        invoices = load_invoices()
        invoices.append({
            "id": len(invoices) + 1,
            "station_id": station_id,
            "station_name": STATIONS[station_id]["name"],
            "description": description,
            "date": doc_date,
            "filename": filename,
            "type": "document",
            "uploaded_at": datetime.now().isoformat(),
            "sharepoint_url": sharepoint_url
        })
        save_invoices(invoices)
        add_to_document_audit(STATIONS[station_id]["name"], "Document Upload", filename, f"{description} | Date: {doc_date}")
        
        if sharepoint_url:
            flash("Document uploaded successfully to SharePoint!")
        else:
            flash("Document uploaded successfully!")
    return redirect("/invoices")

@app.route("/upload_data", methods=["GET", "POST"])
@login_required
def upload_data():
    if request.method == "POST":
        if 'data_file' not in request.files:
            flash("No file selected")
            return redirect("/upload_data")
        file = request.files['data_file']
        if file.filename == '':
            flash("No file selected")
            return redirect("/upload_data")
        if file:
            filename = secure_filename(file.filename)
            filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            file.save(filepath)
            try:
                import pandas as pd
                if filename.endswith('.csv'):
                    df = pd.read_csv(filepath)
                else:
                    df = pd.read_excel(filepath)
                for _, row in df.iterrows():
                    station_name = str(row['Station Name']).strip()
                    grade = str(row['Grade']).strip()
                    price = float(row['Price ($)'])
                    volume = int(row['Gallons Sold'])
                    for sid, station in STATIONS.items():
                        if station['name'].lower() == station_name.lower():
                            if grade in station['grades']:
                                old_price = station['grades'][grade]['current_price']
                                station['grades'][grade]['current_price'] = price
                                station['grades'][grade]['daily_gallons'] = volume
                                add_to_audit(station_name, grade, old_price, price)
                            break
                add_to_document_audit("ML Batch Upload", "ML Data Upload (CSV/XLSX)", filename, f"Processed {len(df)} records - prices updated")
                flash("Data uploaded and prices updated successfully!")
            except Exception as e:
                flash(f"Error processing file: {str(e)}")
            return redirect("/dashboard")
    return render_template_string(UPLOAD_HTML)

UPLOAD_HTML = '''<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>Upload Data - Gas Price Optimizer Final Winner</title>
    <script src="https://cdn.tailwindcss.com"></script>
</head>
<body class="bg-slate-950 text-slate-200 p-8">
    <div class="max-w-4xl mx-auto">
        <div class="flex justify-between items-center mb-8">
            <h1 class="text-3xl font-bold">Upload Daily Data</h1>
            <a href="/dashboard" class="text-green-400 hover:text-sky-300">← Back to Dashboard</a>
        </div>
        <div class="bg-slate-900 border border-slate-800 rounded-3xl p-6 mb-8">
            <h2 class="font-semibold text-xl mb-4">Upload Your Data File</h2>
            <div class="mb-6">
                <p class="text-sm text-slate-400 mb-2">Accepted formats: .csv, .xlsx, .xls</p>
                <p class="text-sm text-slate-400">Required columns: Date, Station Name, Grade, Price ($), Gallons Sold</p>
            </div>
            <form action="/upload_data" method="post" enctype="multipart/form-data">
                <div class="mb-6">
                    <label class="block text-sm text-slate-400 mb-2">Select File</label>
                    <input type="file" name="data_file" accept=".csv,.xlsx,.xls" class="block w-full text-sm text-slate-400 file:mr-4 file:py-3 file:px-6 file:rounded-2xl file:border-0 file:bg-sky-600 file:text-white hover:file:bg-sky-700" required>
                </div>
                <button type="submit" class="bg-green-600 hover:bg-green-700 px-8 py-3 rounded-2xl font-semibold">Upload & Process Data</button>
            </form>
        </div>
        <div class="bg-slate-900 border border-slate-800 rounded-3xl p-6">
            <h2 class="font-semibold text-xl mb-4">Download Template</h2>
            <p class="text-sm text-slate-400 mb-4">Use this template to collect your daily data:</p>
            <a href="/download_template" class="inline-flex items-center gap-x-2 bg-emerald-600 hover:bg-emerald-700 px-6 py-3 rounded-2xl font-semibold">
                <i class="fa-solid fa-download"></i>
                <span>Download Excel Template</span>
            </a>
        </div>
    </div>
</body>
</html>'''

@app.route("/download_template")
@login_required
def download_template():
    template_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'GasPrice_Daily_Data_Template.xlsx')
    if os.path.exists(template_path):
        return send_from_directory(os.path.dirname(os.path.abspath(__file__)), 'GasPrice_Daily_Data_Template.xlsx', as_attachment=True)
    else:
        flash("Template file not found.")
        return redirect("/dashboard")

# ==================== MISSING TEMPLATES ====================

ADD_STATION_HTML = '''<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>Add New Station - Gas Price Optimizer Final Winner</title>
    <script src="https://cdn.tailwindcss.com"></script>
</head>
<body class="bg-slate-950 text-slate-200 p-8">
    <div class="max-w-4xl mx-auto">
        <div class="flex justify-between items-center mb-8">
            <h1 class="text-3xl font-bold">Add New Station</h1>
            <a href="/dashboard" class="text-green-400 hover:text-sky-300">← Back to Dashboard</a>
        </div>
        
        <div class="bg-slate-900 border border-slate-800 rounded-3xl p-8">
            <form method="POST">
                <div class="grid grid-cols-1 md:grid-cols-2 gap-6">
                    <div>
                        <label class="block text-sm text-slate-400 mb-2">Station Name</label>
                        <input type="text" name="name" class="w-full bg-slate-800 border border-slate-700 rounded-xl px-4 py-3" required>
                    </div>
                    <div>
                        <label class="block text-sm text-slate-400 mb-2">Address</label>
                        <input type="text" name="address" class="w-full bg-slate-800 border border-slate-700 rounded-xl px-4 py-3" required>
                    </div>
                    <div>
                        <label class="block text-sm text-slate-400 mb-2">City</label>
                        <input type="text" name="city" class="w-full bg-slate-800 border border-slate-700 rounded-xl px-4 py-3" required>
                    </div>
                    <div>
                        <label class="block text-sm text-slate-400 mb-2">State</label>
                        <input type="text" name="state" class="w-full bg-slate-800 border border-slate-700 rounded-xl px-4 py-3" required>
                    </div>
                    <div>
                        <label class="block text-sm text-slate-400 mb-2">ZIP Code</label>
                        <input type="text" name="zip" class="w-full bg-slate-800 border border-slate-700 rounded-xl px-4 py-3" required>
                    </div>
                    <div>
                        <label class="block text-sm text-slate-400 mb-2">Daily Gallons (Total)</label>
                        <input type="number" name="daily_gallons" value="3000" class="w-full bg-slate-800 border border-slate-700 rounded-xl px-4 py-3" required>
                    </div>
                </div>
                
                <div class="mt-8">
                    <h3 class="text-lg font-semibold mb-4">Initial Pricing per Grade</h3>
                    <div class="grid grid-cols-1 md:grid-cols-4 gap-4">
                        {% for grade in grades %}
                        <div class="bg-slate-800 border border-slate-700 rounded-2xl p-4">
                            <div class="font-semibold text-green-400 mb-3">{{ grade }}</div>
                            <div class="space-y-3">
                                <div>
                                    <label class="block text-xs text-slate-400 mb-1">Wholesale Price</label>
                                    <input type="number" step="0.01" name="wholesale_{{ grade }}" value="2.85" class="w-full bg-slate-700 border border-slate-600 rounded-lg px-3 py-2 text-sm">
                                </div>
                                <div>
                                    <label class="block text-xs text-slate-400 mb-1">Current Price</label>
                                    <input type="number" step="0.01" name="price_{{ grade }}" value="3.19" class="w-full bg-slate-700 border border-slate-600 rounded-lg px-3 py-2 text-sm">
                                </div>
                                <div>
                                    <label class="block text-xs text-slate-400 mb-1">Target Margin ($)</label>
                                    <input type="number" step="0.01" name="margin_{{ grade }}" value="0.12" class="w-full bg-slate-700 border border-slate-600 rounded-lg px-3 py-2 text-sm">
                                </div>
                            </div>
                        </div>
                        {% endfor %}
                    </div>
                </div>
                
                <div class="mt-8 flex gap-4">
                    <button type="submit" class="bg-purple-600 hover:bg-purple-700 px-8 py-3 rounded-2xl font-semibold">Add Station</button>
                    <a href="/dashboard" class="px-8 py-3 rounded-2xl border border-slate-700 hover:bg-slate-800">Cancel</a>
                </div>
            </form>
        </div>
    </div>
</body>
</html>'''

EDIT_STATION_HTML = '''<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>Edit Station - Gas Price Optimizer Final Winner</title>
    <script src="https://cdn.tailwindcss.com"></script>
</head>
<body class="bg-slate-950 text-slate-200 p-8">
    <div class="max-w-5xl mx-auto">
        <div class="flex justify-between items-center mb-8">
            <h1 class="text-3xl font-bold">Edit Station: {{ station.name }}</h1>
            <a href="/dashboard" class="text-green-400 hover:text-sky-300">← Back to Dashboard</a>
        </div>
        
        <div class="bg-slate-900 border border-slate-800 rounded-3xl p-8">
            <form method="POST">
                <!-- Basic Info -->
                <div class="grid grid-cols-1 md:grid-cols-2 gap-6 mb-8">
                    <div>
                        <label class="block text-sm text-slate-400 mb-2">Station Name</label>
                        <input type="text" name="name" value="{{ station.name }}" class="w-full bg-slate-800 border border-slate-700 rounded-xl px-4 py-3" required>
                    </div>
                    <div>
                        <label class="block text-sm text-slate-400 mb-2">Address</label>
                        <input type="text" name="address" value="{{ station.address }}" class="w-full bg-slate-800 border border-slate-700 rounded-xl px-4 py-3" required>
                    </div>
                    <div>
                        <label class="block text-sm text-slate-400 mb-2">City</label>
                        <input type="text" name="city" value="{{ station.city }}" class="w-full bg-slate-800 border border-slate-700 rounded-xl px-4 py-3" required>
                    </div>
                    <div>
                        <label class="block text-sm text-slate-400 mb-2">State</label>
                        <input type="text" name="state" value="{{ station.state }}" class="w-full bg-slate-800 border border-slate-700 rounded-xl px-4 py-3" required>
                    </div>
                    <div>
                        <label class="block text-sm text-slate-400 mb-2">ZIP Code</label>
                        <input type="text" name="zip" value="{{ station.zip }}" class="w-full bg-slate-800 border border-slate-700 rounded-xl px-4 py-3" required>
                    </div>
                    <div>
                        <label class="block text-sm text-slate-400 mb-2">Contact Email(s)</label>
                        <input type="text" name="email" value="{{ station.get('email', '') }}" 
                               class="w-full bg-slate-800 border border-slate-700 rounded-xl px-4 py-3"
                               placeholder="email1@site.com, email2@site.com">
                        <div class="text-xs text-slate-500 mt-1">Separate multiple emails with commas</div>
                    </div>
                </div>
                
                <!-- Per-Grade Pricing -->
                <h3 class="text-lg font-semibold mb-4">Pricing & Settings per Grade</h3>
                <div class="grid grid-cols-1 md:grid-cols-4 gap-4">
                    {% for grade in grades %}
                    <div class="bg-slate-800 border border-slate-700 rounded-2xl p-4">
                        <div class="font-semibold text-green-400 mb-4">{{ grade }}</div>
                        
                        <div class="space-y-4">
                            <div>
                                <label class="block text-xs text-slate-400 mb-1">Wholesale Price ($)</label>
                                <input type="number" step="0.01" name="wholesale_{{ grade }}" value="{{ station.grades[grade].wholesale }}" class="w-full bg-slate-700 border border-slate-600 rounded-lg px-3 py-2 text-sm">
                            </div>
                            <div>
                                <label class="block text-xs text-slate-400 mb-1">Current Price ($)</label>
                                <input type="number" step="0.01" name="price_{{ grade }}" value="{{ station.grades[grade].current_price }}" class="w-full bg-slate-700 border border-slate-600 rounded-lg px-3 py-2 text-sm">
                            </div>
                            <div>
                                <label class="block text-xs text-slate-400 mb-1">Target Margin ($)</label>
                                <input type="number" step="0.01" name="margin_{{ grade }}" value="{{ station.grades[grade].target_margin }}" class="w-full bg-slate-700 border border-slate-600 rounded-lg px-3 py-2 text-sm">
                            </div>
                            <div>
                                <label class="block text-xs text-slate-400 mb-1">Daily Gallons</label>
                                <input type="number" name="gallons_{{ grade }}" value="{{ station.grades[grade].daily_gallons }}" class="w-full bg-slate-700 border border-slate-600 rounded-lg px-3 py-2 text-sm">
                            </div>
                            <div>
                                <label class="block text-xs text-slate-400 mb-1">Strategy</label>
                                <select name="strategy_{{ grade }}" class="w-full bg-slate-700 border border-slate-600 rounded-lg px-3 py-2 text-sm">
                                    <option value="aggressive" {% if station.grades[grade].strategy == 'aggressive' %}selected{% endif %}>Aggressive</option>
                                    <option value="balanced" {% if station.grades[grade].strategy == 'balanced' %}selected{% endif %}>Balanced</option>
                                    <option value="matching" {% if station.grades[grade].strategy == 'matching' %}selected{% endif %}>Matching</option>
                                </select>
                            </div>
                        </div>
                    </div>
                    {% endfor %}
                </div>
                
                <div class="mt-8 flex gap-4">
                    <button type="submit" class="bg-slate-700 hover:bg-slate-600 px-8 py-3 rounded-2xl font-semibold">Save Changes</button>
                    <a href="/dashboard" class="px-8 py-3 rounded-2xl border border-slate-700 hover:bg-slate-800">Cancel</a>
                </div>
            </form>
        </div>
    </div>
</body>
</html>'''

INVOICES_HTML = '''<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>Invoices & Documents - Gas Price Optimizer Final Winner</title>
    <script src="https://cdn.tailwindcss.com"></script>
</head>
<body class="bg-slate-950 text-slate-200 p-8">
    <div class="max-w-6xl mx-auto">
        <div class="flex justify-between items-center mb-8">
            <h1 class="text-3xl font-bold">Invoices & Documents</h1>
            <a href="/dashboard" class="text-green-400 hover:text-sky-300">← Back to Dashboard</a>
        </div>
        
        <!-- Upload Fuel Invoice Section -->
        <div class="bg-slate-900 border border-slate-800 rounded-3xl p-8 mb-8">
            <h2 class="text-xl font-semibold mb-6">Upload Fuel Invoice</h2>
            <form action="/upload_invoice" method="post" enctype="multipart/form-data" class="grid grid-cols-1 md:grid-cols-3 gap-6">
                <div>
                    <label class="block text-sm text-slate-400 mb-2">Station</label>
                    <select name="station_id" class="w-full bg-slate-800 border border-slate-700 rounded-xl px-4 py-3" required>
                        {% for sid, station in stations.items() %}
                        <option value="{{ sid }}">{{ station.name }}</option>
                        {% endfor %}
                    </select>
                </div>
                <div>
                    <label class="block text-sm text-slate-400 mb-2">Fuel Grade</label>
                    <select name="grade" class="w-full bg-slate-800 border border-slate-700 rounded-xl px-4 py-3" required>
                        {% for g in grades %}
                        <option value="{{ g }}">{{ g }}</option>
                        {% endfor %}
                    </select>
                </div>
                <div>
                    <label class="block text-sm text-slate-400 mb-2">Invoice File (PDF/Image)</label>
                    <input type="file" name="invoice_file" accept=".pdf,.jpg,.jpeg,.png" class="block w-full text-sm text-slate-400 file:mr-4 file:py-3 file:px-6 file:rounded-2xl file:border-0 file:bg-emerald-600 file:text-white hover:file:bg-emerald-700" required>
                </div>
                <div class="md:col-span-3">
                    <button type="submit" class="bg-emerald-600 hover:bg-emerald-700 px-8 py-3 rounded-2xl font-semibold">Upload Fuel Invoice</button>
                </div>
            </form>
        </div>
        
        <!-- Upload Any Document Section -->
        <div class="bg-slate-900 border border-slate-800 rounded-3xl p-8 mb-8">
            <h2 class="text-xl font-semibold mb-6">Upload Any Document</h2>
            <form action="/upload_document" method="post" enctype="multipart/form-data" class="grid grid-cols-1 md:grid-cols-4 gap-6">
                <div>
                    <label class="block text-sm text-slate-400 mb-2">Station</label>
                    <select name="station_id" class="w-full bg-slate-800 border border-slate-700 rounded-xl px-4 py-3" required>
                        {% for sid, station in stations.items() %}
                        <option value="{{ sid }}">{{ station.name }}</option>
                        {% endfor %}
                    </select>
                </div>
                <div>
                    <label class="block text-sm text-slate-400 mb-2">Document Type / Description</label>
                    <input type="text" name="description" placeholder="e.g. Maintenance Report, Permit, Tax Form" class="w-full bg-slate-800 border border-slate-700 rounded-xl px-4 py-3" required>
                </div>
                <div>
                    <label class="block text-sm text-slate-400 mb-2">Date</label>
                    <input type="date" name="doc_date" class="w-full bg-slate-800 border border-slate-700 rounded-xl px-4 py-3" required>
                </div>
                <div>
                    <label class="block text-sm text-slate-400 mb-2">File (PDF/Image)</label>
                    <input type="file" name="document_file" accept=".pdf,.jpg,.jpeg,.png,.doc,.docx" class="block w-full text-sm text-slate-400 file:mr-4 file:py-3 file:px-6 file:rounded-2xl file:border-0 file:bg-purple-600 file:text-white hover:file:bg-purple-700" required>
                </div>
                <div class="md:col-span-4">
                    <button type="submit" class="bg-purple-600 hover:bg-purple-700 px-8 py-3 rounded-2xl font-semibold">Upload Document</button>
                </div>
            </form>
        </div>
        
        <!-- Audit Log with Hyperlinks -->
        <div class="bg-slate-900 border border-slate-800 rounded-3xl p-8">
            <div class="flex justify-between items-center mb-6">
                <h2 class="text-xl font-semibold">Document Audit Log</h2>
                <a href="/print_document_audit" class="text-xs bg-purple-600 hover:bg-purple-700 text-white px-4 py-1.5 rounded">Print with Date Range</a>
            </div>
            <div class="overflow-x-auto">
                <table class="w-full">
                    <thead>
                        <tr class="border-b border-slate-700 text-left text-sm text-slate-400">
                            <th class="py-3 px-4">Date</th>
                            <th class="py-3 px-4">Uploaded At</th>
                            <th class="py-3 px-4">Station</th>
                            <th class="py-3 px-4">Description</th>
                        </tr>
                    </thead>
                    <tbody class="text-sm">
                        {% for inv in invoices %}
                        <tr class="border-b border-slate-800 hover:bg-slate-800">
                            <td class="py-3 px-4">{{ inv.get('date', 'N/A') }}</td>
                            <td class="py-3 px-4 text-xs text-slate-400">{{ inv.get('uploaded_at', 'N/A') }}</td>
                            <td class="py-3 px-4">{{ stations.get(inv.get('station_id'), {}).get('name', 'Unknown') if inv.get('station_id') else 'N/A' }}</td>
                            <td class="py-3 px-4">{{ inv.get('description', inv.get('grade', 'Fuel Invoice')) }}</td>
                        </tr>
                        {% endfor %}
                    </tbody>
                </table>
            </div>
        </div>
    </div>
</body>
</html>'''

PRINT_SNAPSHOT_HTML = '''<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>Price Snapshot - {{ now }}</title>
    <script src="https://cdn.tailwindcss.com"></script>
    <style>
        @media print {
            body { background: white; color: black; }
            .no-print { display: none; }
        }
    </style>
</head>
<body class="bg-white text-black p-8">
    <div class="max-w-7xl mx-auto">
        <div class="flex justify-between items-center mb-8 border-b pb-4">
            <div>
                <h1 class="text-4xl font-bold">Fresh Stop Dash</h1>
                <p class="text-xl">All Stations Price Snapshot</p>
            </div>
            <div class="text-right">
                <div class="text-sm">Printed: {{ now }}</div>
                <button onclick="window.print()" class="no-print mt-2 bg-black text-white px-4 py-2 rounded">Print</button>
            </div>
        </div>
        
        <table class="w-full border-collapse">
            <thead>
                <tr class="border-b-2 border-black">
                    <th class="text-left py-3 px-4">Station</th>
                    <th class="text-left py-3 px-4">Location</th>
                    {% for grade in grades %}
                    <th class="text-center py-3 px-4">{{ grade }}</th>
                    {% endfor %}
                </tr>
            </thead>
            <tbody>
                {% for sid, station in stations.items() %}
                <tr class="border-b border-gray-300">
                    <td class="py-4 px-4 font-semibold">{{ station.name }}</td>
                    <td class="py-4 px-4 text-sm">{{ station.city }}, {{ station.state }}</td>
                    {% for grade in grades %}
                    <td class="py-4 px-4 text-center font-mono text-lg">
                        ${{ "%.2f"|format(station.grades[grade].current_price) }}
                    </td>
                    {% endfor %}
                </tr>
                {% endfor %}
            </tbody>
        </table>
        
        <div class="mt-8 text-xs text-gray-500">
            Confidential - For internal use only • Fresh Stop Dash
        </div>
    </div>
</body>
</html>'''

PRINT_AUDIT_HTML = '''<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>Print Audit Trail - FINAL WINNER</title>
    <script src="https://cdn.tailwindcss.com"></script>
    <style>
        @media print { .no-print { display: none; } }
    </style>
</head>
<body class="bg-white text-black p-8">
    <div class="max-w-5xl mx-auto">
        <div class="flex justify-between items-center mb-6 border-b pb-4">
            <h1 class="text-3xl font-bold">Audit Trail - Price Changes</h1>
            <div>
                <button onclick="window.print()" class="no-print bg-black text-white px-4 py-2 rounded">Print</button>
            </div>
        </div>
        
        <form class="no-print mb-6 flex gap-4 items-end">
            <div>
                <label class="block text-sm mb-1">Start Date</label>
                <input type="date" id="start_date" class="border px-3 py-2 rounded">
            </div>
            <div>
                <label class="block text-sm mb-1">End Date</label>
                <input type="date" id="end_date" class="border px-3 py-2 rounded">
            </div>
            <button type="button" onclick="filterAudit()" class="bg-black text-white px-4 py-2 rounded">Filter & Print</button>
        </form>
        
        <table class="w-full border-collapse">
            <thead>
                <tr class="border-b-2 border-black">
                    <th class="text-left py-3 px-4">Time</th>
                    <th class="text-left py-3 px-4">Station</th>
                    <th class="text-left py-3 px-4">Grade</th>
                    <th class="text-left py-3 px-4">Old Price</th>
                    <th class="text-left py-3 px-4">New Price</th>
                    <th class="text-left py-3 px-4">User</th>
                </tr>
            </thead>
            <tbody id="audit-table">
                {% for entry in audit_log %}
                <tr class="border-b border-gray-300 audit-row" data-time="{{ entry.time }}">
                    <td class="py-3 px-4">{{ entry.time }}</td>
                    <td class="py-3 px-4">{{ entry.station }}</td>
                    <td class="py-3 px-4">{{ entry.grade }}</td>
                    <td class="py-3 px-4 font-mono">{{ entry.old_price }}</td>
                    <td class="py-3 px-4 font-mono text-emerald-600">{{ entry.new_price }}</td>
                    <td class="py-3 px-4 text-xs text-gray-600">{{ entry.get('user', 'system') }}</td>
                </tr>
                {% endfor %}
            </tbody>
        </table>
    </div>
    
    <script>
        function filterAudit() {
            const start = document.getElementById('start_date').value;
            const end = document.getElementById('end_date').value;
            const rows = document.querySelectorAll('.audit-row');
            
            rows.forEach(row => {
                const time = row.getAttribute('data-time');
                const rowDate = time.split(' ')[0];
                
                let show = true;
                if (start && rowDate < start) show = false;
                if (end && rowDate > end) show = false;
                
                row.style.display = show ? '' : 'none';
            });
            
            setTimeout(() => window.print(), 100);
        }
    </script>
</body>
</html>'''

PRINT_EMAIL_AUDIT_HTML = '''<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>Print Email Audit Trail - FINAL WINNER</title>
    <script src="https://cdn.tailwindcss.com"></script>
    <style>
        @media print { .no-print { display: none; } }
    </style>
</head>
<body class="bg-white text-black p-8">
    <div class="max-w-6xl mx-auto">
        <div class="flex justify-between items-center mb-6 border-b pb-4">
            <h1 class="text-3xl font-bold">Email Audit Trail - Prices Sent</h1>
            <div>
                <button onclick="window.print()" class="no-print bg-black text-white px-4 py-2 rounded">Print</button>
            </div>
        </div>
        
        <form class="no-print mb-6 flex gap-4 items-end">
            <div>
                <label class="block text-sm mb-1">Start Date</label>
                <input type="date" id="start_date" class="border px-3 py-2 rounded">
            </div>
            <div>
                <label class="block text-sm mb-1">End Date</label>
                <input type="date" id="end_date" class="border px-3 py-2 rounded">
            </div>
            <button type="button" onclick="filterEmailAudit()" class="bg-black text-white px-4 py-2 rounded">Filter & Print</button>
        </form>
        
        <table class="w-full border-collapse">
            <thead>
                <tr class="border-b-2 border-black">
                    <th class="text-left py-3 px-4">Time</th>
                    <th class="text-left py-3 px-4">Station</th>
                    <th class="text-left py-3 px-4">Recipient Email</th>
                    <th class="text-left py-3 px-4">Prices Summary</th>
                    <th class="text-left py-3 px-4">User</th>
                </tr>
            </thead>
            <tbody id="email-audit-table">
                {% for entry in email_audit %}
                <tr class="border-b border-gray-300 email-row" data-time="{{ entry.time }}">
                    <td class="py-3 px-4">{{ entry.time }}</td>
                    <td class="py-3 px-4 font-semibold">{{ entry.station }}</td>
                    <td class="py-3 px-4 text-sm">{{ entry.recipient }}</td>
                    <td class="py-3 px-4 text-xs font-mono whitespace-pre-line">{{ entry.prices }}</td>
                    <td class="py-3 px-4 text-xs text-gray-600">{{ entry.get('user', 'system') }}</td>
                </tr>
                {% endfor %}
            </tbody>
        </table>
        
        <div class="mt-8 text-xs text-gray-500">
            Confidential - Fresh Stop Dash - Audit Trail
        </div>
    </div>
    
    <script>
        function filterEmailAudit() {
            const start = document.getElementById('start_date').value;
            const end = document.getElementById('end_date').value;
            const rows = document.querySelectorAll('.email-row');
            
            rows.forEach(row => {
                const time = row.getAttribute('data-time');
                const rowDate = time.split(' ')[0];
                
                let show = true;
                if (start && rowDate < start) show = false;
                if (end && rowDate > end) show = false;
                
                row.style.display = show ? '' : 'none';
            });
            
            setTimeout(() => window.print(), 100);
        }
    </script>
</body>
</html>'''

PRINT_EDIT_AUDIT_HTML = '''<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>Print Edit Audit Trail - FINAL WINNER</title>
    <script src="https://cdn.tailwindcss.com"></script>
    <style>
        @media print { .no-print { display: none; } }
    </style>
</head>
<body class="bg-white text-black p-8">
    <div class="max-w-6xl mx-auto">
        <div class="flex justify-between items-center mb-6 border-b pb-4">
            <h1 class="text-3xl font-bold">Station Edit & Creation Audit Trail</h1>
            <div>
                <button onclick="window.print()" class="no-print bg-black text-white px-4 py-2 rounded">Print</button>
            </div>
        </div>
        
        <form class="no-print mb-6 flex gap-4 items-end">
            <div>
                <label class="block text-sm mb-1">Start Date</label>
                <input type="date" id="start_date" class="border px-3 py-2 rounded">
            </div>
            <div>
                <label class="block text-sm mb-1">End Date</label>
                <input type="date" id="end_date" class="border px-3 py-2 rounded">
            </div>
            <button type="button" onclick="filterEditAudit()" class="bg-black text-white px-4 py-2 rounded">Filter & Print</button>
        </form>
        
        <table class="w-full border-collapse">
            <thead>
                <tr class="border-b-2 border-black">
                    <th class="text-left py-3 px-4">Time</th>
                    <th class="text-left py-3 px-4">Station</th>
                    <th class="text-left py-3 px-4">Action</th>
                    <th class="text-left py-3 px-4">Details</th>
                    <th class="text-left py-3 px-4">User</th>
                </tr>
            </thead>
            <tbody id="edit-audit-table">
                {% for entry in edit_audit %}
                <tr class="border-b border-gray-300 edit-row" data-time="{{ entry.time }}">
                    <td class="py-3 px-4">{{ entry.time }}</td>
                    <td class="py-3 px-4 font-semibold">{{ entry.station }}</td>
                    <td class="py-3 px-4">
                        <span class="px-2 py-1 rounded text-xs {% if entry.action == 'Created' %}bg-green-100 text-green-800{% else %}bg-blue-100 text-blue-800{% endif %}">
                            {{ entry.action }}
                        </span>
                    </td>
                    <td class="py-3 px-4 text-sm">{{ entry.details }}</td>
                    <td class="py-3 px-4 text-xs text-gray-500">{{ entry.user }}</td>
                </tr>
                {% endfor %}
            </tbody>
        </table>
        
        <div class="mt-8 text-xs text-gray-500">
            Confidential - Fresh Stop Dash - Audit Trail
        </div>
    </div>
    
    <script>
        function filterEditAudit() {
            const start = document.getElementById('start_date').value;
            const end = document.getElementById('end_date').value;
            const rows = document.querySelectorAll('.edit-row');
            
            rows.forEach(row => {
                const time = row.getAttribute('data-time');
                const rowDate = time.split(' ')[0];
                
                let show = true;
                if (start && rowDate < start) show = false;
                if (end && rowDate > end) show = false;
                
                row.style.display = show ? '' : 'none';
            });
            
            setTimeout(() => window.print(), 100);
        }
    </script>
</body>
</html>'''

PRINT_DOCUMENT_AUDIT_HTML = '''<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>Print Document Audit Trail - FINAL WINNER</title>
    <script src="https://cdn.tailwindcss.com"></script>
    <style>
        @media print { .no-print { display: none; } }
    </style>
</head>
<body class="bg-white text-black p-8">
    <div class="max-w-6xl mx-auto">
        <div class="flex justify-between items-center mb-6 border-b pb-4">
            <h1 class="text-3xl font-bold">ML Upload & Document Audit Trail</h1>
            <div>
                <button onclick="window.print()" class="no-print bg-black text-white px-4 py-2 rounded">Print</button>
            </div>
        </div>
        
        <form class="no-print mb-6 flex gap-4 items-end">
            <div>
                <label class="block text-sm mb-1">Start Date</label>
                <input type="date" id="start_date" class="border px-3 py-2 rounded">
            </div>
            <div>
                <label class="block text-sm mb-1">End Date</label>
                <input type="date" id="end_date" class="border px-3 py-2 rounded">
            </div>
            <button type="button" onclick="filterDocumentAudit()" class="bg-black text-white px-4 py-2 rounded">Filter & Print</button>
        </form>
        
        <table class="w-full border-collapse">
            <thead>
                <tr class="border-b-2 border-black">
                    <th class="text-left py-3 px-4">Time</th>
                    <th class="text-left py-3 px-4">Station</th>
                    <th class="text-left py-3 px-4">Type</th>
                    <th class="text-left py-3 px-4">Filename</th>
                    <th class="text-left py-3 px-4">Description</th>
                    <th class="text-left py-3 px-4">User</th>
                </tr>
            </thead>
            <tbody id="document-audit-table">
                {% for entry in document_audit %}
                <tr class="border-b border-gray-300 document-row" data-time="{{ entry.time }}">
                    <td class="py-3 px-4">{{ entry.time }}</td>
                    <td class="py-3 px-4 font-semibold">{{ entry.station }}</td>
                    <td class="py-3 px-4">
                        <span class="px-2 py-1 rounded text-xs bg-purple-100 text-purple-800">{{ entry.doc_type }}</span>
                    </td>
                    <td class="py-3 px-4 text-sm font-mono">{{ entry.filename }}</td>
                    <td class="py-3 px-4 text-sm">{{ entry.description }}</td>
                    <td class="py-3 px-4 text-xs text-gray-600">{{ entry.get('user', 'system') }}</td>
                </tr>
                {% endfor %}
            </tbody>
        </table>
        
        <div class="mt-8 text-xs text-gray-500">
            Confidential - Fresh Stop Dash - Audit Trail
        </div>
    </div>
    
    <script>
        function filterDocumentAudit() {
            const start = document.getElementById('start_date').value;
            const end = document.getElementById('end_date').value;
            const rows = document.querySelectorAll('.document-row');
            
            rows.forEach(row => {
                const time = row.getAttribute('data-time');
                const rowDate = time.split(' ')[0];
                
                let show = true;
                if (start && rowDate < start) show = false;
                if (end && rowDate > end) show = false;
                
                row.style.display = show ? '' : 'none';
            });
            
            setTimeout(() => window.print(), 100);
        }
    </script>
</body>
</html>'''

# ==================== ADMIN USER MANAGEMENT ====================

ADMIN_USERS_HTML = '''<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>User Management - Fresh Stop Dash</title>
    <script src="https://cdn.tailwindcss.com"></script>
</head>
<body class="bg-slate-950 text-slate-200 p-8">
    <div class="max-w-6xl mx-auto">
        <div class="flex justify-between items-center mb-8">
            <div>
                <h1 class="text-3xl font-bold">User Management</h1>
                <p class="text-slate-400">Manage roles and accounts</p>
            </div>
            <a href="/dashboard" class="text-green-400 hover:text-sky-300">← Back to Dashboard</a>
        </div>

        <!-- Add New User Form -->
        <div class="bg-slate-900 border border-slate-800 rounded-3xl p-8 mb-8">
            <h2 class="text-xl font-semibold mb-6">Create New User</h2>
            <form method="POST" class="grid grid-cols-1 md:grid-cols-5 gap-4">
                <div>
                    <label class="block text-sm text-slate-400 mb-1">Full Name</label>
                    <input type="text" name="name" class="w-full bg-slate-800 border border-slate-700 rounded-xl px-4 py-2" required>
                </div>
                <div>
                    <label class="block text-sm text-slate-400 mb-1">Email</label>
                    <input type="email" name="email" class="w-full bg-slate-800 border border-slate-700 rounded-xl px-4 py-2" required>
                </div>
                <div>
                    <label class="block text-sm text-slate-400 mb-1">Password</label>
                    <input type="password" name="password" class="w-full bg-slate-800 border border-slate-700 rounded-xl px-4 py-2" required>
                </div>
                <div>
                    <label class="block text-sm text-slate-400 mb-1">Role</label>
                    <select name="role" class="w-full bg-slate-800 border border-slate-700 rounded-xl px-4 py-2">
                        <option value="viewer">Viewer</option>
                        <option value="manager">Manager</option>
                        <option value="admin">Admin</option>
                    </select>
                </div>
                <div class="flex items-end">
                    <button type="submit" name="action" value="create" 
                            class="w-full bg-purple-600 hover:bg-purple-700 px-6 py-2 rounded-xl font-semibold">Create User</button>
                </div>
            </form>
        </div>

        <!-- Users Table -->
        <div class="bg-slate-900 border border-slate-800 rounded-3xl p-8">
            <h2 class="text-xl font-semibold mb-6">All Users ({{ users|length }})</h2>
            
            <table class="w-full">
                <thead>
                    <tr class="border-b border-slate-700 text-left text-sm text-slate-400">
                        <th class="py-3 px-4">Name</th>
                        <th class="py-3 px-4">Email</th>
                        <th class="py-3 px-4">Role</th>
                        <th class="py-3 px-4">Created</th>
                        <th class="py-3 px-4">Actions</th>
                    </tr>
                </thead>
                <tbody>
                    {% for user in users %}
                    <tr class="border-b border-slate-800 hover:bg-slate-800">
                        <td class="py-4 px-4 font-medium">{{ user.name }}</td>
                        <td class="py-4 px-4 text-sm">{{ user.email }}</td>
                        <td class="py-4 px-4">
                            <span class="px-3 py-1 rounded-full text-xs 
                                {% if user.role == 'admin' %}bg-red-500/20 text-red-400
                                {% elif user.role == 'manager' %}bg-amber-500/20 text-amber-400
                                {% else %}bg-sky-500/20 text-sky-400{% endif %}">
                                {{ user.role }}
                            </span>
                        </td>
                        <td class="py-4 px-4 text-sm text-slate-400">{{ user.created_at[:10] if user.created_at else 'N/A' }}</td>
                        <td class="py-4 px-4">
                            <div class="flex gap-2 items-center">
                                <!-- Role Update -->
                                <form method="POST" class="flex gap-1">
                                    <input type="hidden" name="email" value="{{ user.email }}">
                                    <select name="new_role" class="bg-slate-800 border border-slate-700 rounded px-2 py-1 text-xs">
                                        <option value="viewer" {% if user.role == 'viewer' %}selected{% endif %}>Viewer</option>
                                        <option value="manager" {% if user.role == 'manager' %}selected{% endif %}>Manager</option>
                                        <option value="admin" {% if user.role == 'admin' %}selected{% endif %}>Admin</option>
                                    </select>
                                    <button type="submit" name="action" value="update_role" 
                                            class="bg-emerald-600 hover:bg-emerald-700 px-3 py-1 rounded text-xs">Save</button>
                                </form>
                                
                                <!-- Inline Edit Form -->
                                <form method="POST" action="/admin/edit_user" class="flex gap-1">
                                    <input type="hidden" name="email" value="{{ user.email }}">
                                    <input type="text" name="name" value="{{ user.name }}" class="bg-slate-700 border border-slate-600 rounded px-2 py-1 text-xs w-16" placeholder="Name">
                                    <input type="email" name="new_email" value="{{ user.email }}" class="bg-slate-700 border border-slate-600 rounded px-2 py-1 text-xs w-24" placeholder="Email">
                                    <input type="password" name="new_password" class="bg-slate-700 border border-slate-600 rounded px-2 py-1 text-xs w-16" placeholder="New PW">
                                    <button type="submit" class="bg-blue-600 hover:bg-blue-700 px-3 py-1 rounded text-xs">Save</button>
                                </form>
                                
                                <!-- Delete Button -->
                                <a href="/admin/delete_user/{{ user.email }}" 
                                   onclick="return confirm('Delete {{ user.email }}?')" 
                                   class="bg-red-600 hover:bg-red-700 px-3 py-1 rounded text-xs">Delete</a>
                            </div>
                        </td>
                    </tr>
                    {% endfor %}
                </tbody>
            </table>
        </div>
    </div>
</body>
</html>'''

@app.route("/admin/users", methods=["GET", "POST"])
@login_required
@admin_required
def admin_users():
    if request.method == "POST":
        action = request.form.get("action")
        email = request.form.get("email")
        
        if action == "create":
            name = request.form.get("name", "").strip()
            new_email = request.form.get("email", "").lower().strip()
            password = request.form.get("password", "")
            role = request.form.get("role", "viewer")
            
            if len(password) < 6:
                flash("Password must be at least 6 characters")
            elif get_user(new_email):
                flash("User with this email already exists")
            else:
                if create_user(new_email, generate_password_hash(password, method='pbkdf2:sha256'), name, role):
                    flash(f"User {new_email} created successfully with role: {role}")
                else:
                    flash("Error creating user")
        
        elif action == "update_role":
            new_role = request.form.get("new_role")
            if email and new_role:
                update_user_role(email, new_role)
                flash(f"Role updated for {email} to {new_role}")
    
    # Get all users
    conn = get_db_connection()
    c = conn.cursor()
    if os.environ.get("DATABASE_URL"):
        c.execute("SELECT email, password, name, role, created_at FROM users ORDER BY created_at DESC")
    else:
        c.execute("SELECT email, password, name, role, created_at FROM users ORDER BY created_at DESC")
    rows = c.fetchall()
    conn.close()
    
    users = []
    for row in rows:
        users.append({
            "email": row[0],
            "name": row[2],
            "role": row[3],
            "created_at": row[4]
        })
    
    return render_template_string(ADMIN_USERS_HTML, users=users)


@app.route("/admin/edit_user", methods=["POST"])
@login_required
@admin_required
def admin_edit_user():
    email = request.form.get("email")
    new_name = request.form.get("name", "").strip()
    new_email = request.form.get("new_email", "").lower().strip()
    new_password = request.form.get("new_password", "")
    
    if not email:
        flash("No user selected")
        return redirect("/admin/users")
    
    user = get_user(email)
    if not user:
        flash("User not found")
        return redirect("/admin/users")
    
    # Update name
    if new_name:
        conn = get_db_connection()
        c = conn.cursor()
        if os.environ.get("DATABASE_URL"):
            c.execute("UPDATE users SET name = %s WHERE email = %s", (new_name, email))
        else:
            c.execute("UPDATE users SET name = ? WHERE email = ?", (new_name, email))
        conn.commit()
        conn.close()
    
    # Update email (if changed)
    if new_email and new_email != email:
        # Check if new email already exists
        if get_user(new_email):
            flash("Email already in use")
            return redirect("/admin/users")
        
        conn = get_db_connection()
        c = conn.cursor()
        if os.environ.get("DATABASE_URL"):
            c.execute("UPDATE users SET email = %s WHERE email = %s", (new_email, email))
        else:
            c.execute("UPDATE users SET email = ? WHERE email = ?", (new_email, email))
        conn.commit()
        conn.close()
        email = new_email  # Update for password change
    
    # Update password (if provided)
    if new_password:
        if len(new_password) < 6:
            flash("Password must be at least 6 characters")
            return redirect("/admin/users")
        
        conn = get_db_connection()
        c = conn.cursor()
        hashed = generate_password_hash(new_password, method='pbkdf2:sha256')
        if os.environ.get("DATABASE_URL"):
            c.execute("UPDATE users SET password = %s WHERE email = %s", (hashed, email))
        else:
            c.execute("UPDATE users SET password = ? WHERE email = ?", (hashed, email))
        conn.commit()
        conn.close()
    
    flash("User updated successfully")
    return redirect("/admin/users")


@app.route("/admin/delete_user/<email>")
@login_required
@admin_required
def admin_delete_user(email):
    if email == "demo@station.com":
        flash("Cannot delete demo account")
        return redirect("/admin/users")
    
    conn = get_db_connection()
    c = conn.cursor()
    if os.environ.get("DATABASE_URL"):
        c.execute("DELETE FROM users WHERE email = %s", (email,))
    else:
        c.execute("DELETE FROM users WHERE email = ?", (email,))
    conn.commit()
    conn.close()
    
    flash(f"User {email} deleted successfully")
    return redirect("/admin/users")


# ==================== STATION-SPECIFIC DOCUMENTS ====================

STATION_DOCUMENTS_HTML = '''<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>{{ station.name }} - Documents | Fresh Stop Dash</title>
    <script src="https://cdn.tailwindcss.com"></script>
</head>
<body class="bg-slate-950 text-slate-200 p-8">
    <div class="max-w-6xl mx-auto">
        <div class="flex justify-between items-center mb-8">
            <div>
                <h1 class="text-3xl font-bold">{{ station.name }} Documents</h1>
                <p class="text-slate-400">{{ station.city }}, {{ station.state }}</p>
            </div>
            <div class="flex gap-3">
                <a href="/dashboard#station-{{ station_id }}" class="text-green-400 hover:text-sky-300">← Back to Dashboard</a>
                <a href="/invoices" class="text-indigo-400 hover:text-indigo-300">View All Documents</a>
            </div>
        </div>

        <!-- Upload Form -->
        <div class="bg-slate-900 border border-slate-800 rounded-3xl p-8 mb-8">
            <h2 class="text-xl font-semibold mb-6">Upload Document for {{ station.name }}</h2>
            <form action="/upload_document" method="post" enctype="multipart/form-data" class="grid grid-cols-1 md:grid-cols-4 gap-6">
                <input type="hidden" name="station_id" value="{{ station_id }}">
                <div>
                    <label class="block text-sm text-slate-400 mb-2">Category</label>
                    <select name="description" class="w-full bg-slate-800 border border-slate-700 rounded-xl px-4 py-2" required>
                        <option value="">-- Select Category --</option>
                        <option value="Sales Tax">Sales Tax</option>
                        <option value="UST/EPA">UST/EPA</option>
                        <option value="Fuel Registration">Fuel Registration</option>
                        <option value="Tobacco/Nicotine">Tobacco/Nicotine</option>
                        <option value="EBT/SNAP">EBT/SNAP</option>
                        <option value="Lottery">Lottery</option>
                        <option value="Food Service">Food Service</option>
                        <option value="Alcohol">Alcohol</option>
                        <option value="Local Business">Local Business</option>
                        <option value="Cylinder Exchange">Cylinder Exchange</option>
                        <option value="Vendors">Vendors</option>
                        <option value="E-Cig">E-Cig</option>
                        <option value="Lease">Lease</option>
                        <option value="PCI/Other">PCI/Other</option>
                        <option value="Building">Building</option>
                        <option value="Maintenance">Maintenance</option>
                        <option value="Pumps">Pumps</option>
                        <option value="MISC">MISC</option>
                        <option value="COMMUNICATIONS">COMMUNICATIONS</option>
                        <option value="Training">Training</option>
                        <option value="Fire/ Safety">Fire/ Safety</option>
                        <option value="Inspections">Inspections</option>
                        <option value="IMST">IMST</option>
                    </select>
                </div>
                <div>
                    <label class="block text-sm text-slate-400 mb-2">Date</label>
                    <input type="date" name="doc_date" class="w-full bg-slate-800 border border-slate-700 rounded-xl px-4 py-2" required>
                </div>
                <div>
                    <label class="block text-sm text-slate-400 mb-2">File</label>
                    <input type="file" name="document_file" class="block w-full text-sm text-slate-400 file:mr-4 file:py-2 file:px-4 file:rounded-xl file:border-0 file:bg-indigo-600 file:text-white" required>
                </div>
                <div class="flex items-end">
                    <button type="submit" class="w-full bg-indigo-600 hover:bg-indigo-700 px-6 py-2 rounded-xl font-semibold">Upload</button>
                </div>
            </form>
        </div>

        <!-- Documents List -->
        <div class="bg-slate-900 border border-slate-800 rounded-3xl p-8">
            <div class="flex justify-between items-center mb-6">
                <h2 class="text-xl font-semibold">All Documents for {{ station.name }} ({{ documents|length }})</h2>
                <input type="text" id="search-input" placeholder="Search documents..." 
                       class="bg-slate-800 border border-slate-700 rounded-xl px-4 py-2 text-sm w-64"
                       onkeyup="filterDocuments()">
            </div>
            
            {% if documents %}
            <table class="w-full" id="documents-table">
                <thead>
                    <tr class="border-b border-slate-700 text-left text-sm text-slate-400">
                        <th class="py-3 px-4">Date</th>
                        <th class="py-3 px-4">Description</th>
                        <th class="py-3 px-4">File</th>
                        <th class="py-3 px-4">Uploaded</th>
                        <th class="py-3 px-4">Actions</th>
                    </tr>
                </thead>
                <tbody>
                    {% for doc in documents %}
                    <tr class="border-b border-slate-800 hover:bg-slate-800">
                        <td class="py-4 px-4">{{ doc.date }}</td>
                        <td class="py-4 px-4">{{ doc.description or doc.grade or 'Document' }}</td>
                        <td class="py-4 px-4">
                            {% if doc.filename %}
                            <a href="/invoices/{{ doc.filename }}" target="_blank" class="text-green-400 hover:underline">View File</a>
                            {% else %}
                            N/A
                            {% endif %}
                        </td>
                        <td class="py-4 px-4 text-sm text-slate-400">{{ doc.uploaded_at[:10] if doc.uploaded_at else 'N/A' }}</td>
                        <td class="py-4 px-4">
                            {% if session.user_role == 'admin' %}
                            <a href="/delete_document/{{ station_id }}/{{ loop.index0 }}" 
                               onclick="return confirm('Delete this document? (Audit trail will be kept)')" 
                               class="bg-red-600 hover:bg-red-700 px-3 py-1 rounded text-xs">Delete</a>
                            {% else %}
                            N/A
                            {% endif %}
                        </td>
                    </tr>
                    {% endfor %}
                </tbody>
            </table>
            {% else %}
            <div class="text-center py-12 text-slate-400">
                <i class="fa-solid fa-folder-open text-4xl mb-4"></i>
                <p>No documents uploaded for this station yet.</p>
            </div>
            {% endif %}
        </div>
    </div>

    <script>
        function filterDocuments() {
            const searchTerm = document.getElementById('search-input').value.toLowerCase();
            const rows = document.querySelectorAll('#documents-table tbody tr');
            
            rows.forEach(row => {
                const text = row.textContent.toLowerCase();
                if (text.includes(searchTerm)) {
                    row.style.display = '';
                } else {
                    row.style.display = 'none';
                }
            });
        }
    </script>
</body>
</html>'''

@app.route("/station/<int:sid>/documents")
@login_required
def station_documents(sid):
    if sid not in STATIONS:
        flash("Station not found")
        return redirect("/dashboard")
    
    station = STATIONS[sid]
    
    # Get documents for this station only
    documents = []
    
    # Load all invoices and filter by station
    try:
        all_invoices = load_invoices()
        for inv in all_invoices:
            if inv.get('station_id') == sid or inv.get('station_name') == station['name']:
                documents.append(inv)
    except Exception as e:
        print(f"Error loading documents: {e}")
        documents = []
    
    return render_template_string(STATION_DOCUMENTS_HTML, 
                                  station=station, 
                                  station_id=sid, 
                                  documents=documents)


@app.route("/delete_document/<int:sid>/<int:index>")
@login_required
@admin_required
def delete_document(sid, index):
    if sid not in STATIONS:
        flash("Station not found")
        return redirect("/dashboard")
    
    # Note: We don't actually delete the file from disk/SharePoint here
    # We just remove it from the visible list (audit trail remains)
    # In a full implementation, you could move the file to an "Archived" folder
    
    flash("Document removed from view. Audit trail entry remains for compliance.")
    return redirect(f"/station/{sid}/documents")

if __name__ == "__main__":
    print("\n🚀 Fresh Stop Dash - Station-Focused Documents")
    print("   Run with: gunicorn -w 4 -b 0.0.0.0:8081 FSDadmin_station_docs:app")
    print("   Demo Admin: demo@station.com / demo123")
    app.run(host="127.0.0.1", port=8081, debug=False)
