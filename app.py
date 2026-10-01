from flask import Flask, render_template, request, jsonify, session, redirect, Response
from werkzeug.security import generate_password_hash, check_password_hash
import pandas as pd
import numpy as np
import os
import uuid
import re
import sqlite3
import json
import math
import time
from datetime import datetime

from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib import colors
from functools import wraps

app = Flask(__name__)

# Change this in production. You can also set the DATALENS_SECRET_KEY
# environment variable instead of using the development fallback.
app.secret_key = os.environ.get(
    "DATALENS_SECRET_KEY",
    "datalens-development-secret-key-change-later"
)

UPLOAD_FOLDER = "uploads"
DATABASE = "datalens.db"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024


# =========================================================
# SQLITE DATABASE
# =========================================================

def get_db():
    conn = sqlite3.connect(DATABASE)
    conn.row_factory = sqlite3.Row
    return conn


def init_database():
    conn = get_db()
    cursor = conn.cursor()

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT UNIQUE NOT NULL,
            email TEXT,
            password_hash TEXT NOT NULL,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    """)

    # If an older DataLens database already has the users table
    # without an email column, add the column safely.
    columns = [
        row["name"]
        for row in cursor.execute("PRAGMA table_info(users)").fetchall()
    ]

    if "email" not in columns:
        cursor.execute("ALTER TABLE users ADD COLUMN email TEXT")

    # Keep the earlier default account available for testing.
    default_username = "arokiya"
    default_password = "datalens123"
    default_email = "arokiyaumaw@gmail.com"

    existing = cursor.execute(
        "SELECT id FROM users WHERE username = ?",
        (default_username,)
    ).fetchone()

    if existing is None:
        cursor.execute(
            """
            INSERT INTO users (username, email, password_hash)
            VALUES (?, ?, ?)
            """,
            (
                default_username,
                default_email,
                generate_password_hash(default_password)
            )
        )
    else:
        # Fill email for the existing default account if it is empty.
        cursor.execute(
            """
            UPDATE users
            SET email = COALESCE(NULLIF(email, ''), ?)
            WHERE username = ?
            """,
            (default_email, default_username)
        )

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS saved_dashboards (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id INTEGER NOT NULL,
            name TEXT NOT NULL,
            filename TEXT,
            dashboard_json TEXT NOT NULL,
            share_token TEXT UNIQUE,
            is_shared INTEGER DEFAULT 0,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY(user_id) REFERENCES users(id)
        )
    """)

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS engineering_runs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id INTEGER NOT NULL,
            source_filename TEXT,
            status TEXT NOT NULL DEFAULT 'SUCCESS',
            run_type TEXT NOT NULL DEFAULT 'FULL',
            rows_processed INTEGER DEFAULT 0,
            quality_score REAL DEFAULT 0,
            inserted_rows INTEGER DEFAULT 0,
            updated_rows INTEGER DEFAULT 0,
            deleted_rows INTEGER DEFAULT 0,
            unchanged_rows INTEGER DEFAULT 0,
            duration_ms REAL DEFAULT 0,
            message TEXT,
            result_json TEXT,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY(user_id) REFERENCES users(id)
        )
    """)

    conn.commit()
    conn.close()


init_database()


# =========================================================
# LOGIN REQUIRED
# =========================================================

def login_required(function):
    @wraps(function)
    def wrapper(*args, **kwargs):
        if "user_id" not in session:
            if request.path in ["/", "/login", "/register"]:
                return redirect("/login")

            return jsonify({
                "success": False,
                "error": "Please login first.",
                "redirect": "/login"
            }), 401

        return function(*args, **kwargs)

    return wrapper


# =========================================================
# LOAD CSV / EXCEL
# =========================================================

def load_file(filepath):
    extension = filepath.lower().split(".")[-1]

    if extension == "csv":
        return pd.read_csv(filepath)

    elif extension in ["xlsx", "xls"]:
        return pd.read_excel(filepath)

    raise ValueError("Unsupported file format")


# =========================================================
# SAFE VALUE CONVERSION
# =========================================================

def clean_value(value):
    if pd.isna(value):
        return None

    if isinstance(value, (np.integer,)):
        return int(value)

    if isinstance(value, (np.floating,)):
        return float(value)

    try:
        if hasattr(value, "item"):
            return value.item()
    except Exception:
        pass

    return str(value)


# =========================================================
# COLUMN NAME HELPERS
# =========================================================

def normalize_column_name(column):
    return re.sub(
        r"[^a-z0-9]+",
        " ",
        str(column).lower()
    ).strip()


# =========================================================
# INTELLIGENT COLUMN CLASSIFICATION
# =========================================================

def classify_columns(df):

    numeric_columns = []
    categorical_columns = []
    date_columns = []
    id_columns = []
    name_columns = []
    ignored_columns = []

    for column in df.columns:

        name = normalize_column_name(column)
        series = df[column]

        # -------------------------------------------------
        # ID DETECTION
        # -------------------------------------------------

        id_keywords = [
            "id",
            "student id",
            "customer id",
            "employee id",
            "order id",
            "product id",
            "transaction id",
            "user id",
            "roll no",
            "roll number",
            "registration number",
            "register number"
        ]

        if (
            name in id_keywords
            or name.endswith(" id")
            or name.startswith("id ")
            or "identifier" in name
        ):
            id_columns.append(column)
            continue

        # -------------------------------------------------
        # NAME DETECTION
        # -------------------------------------------------

        name_keywords = [
            "name",
            "student name",
            "customer name",
            "employee name",
            "person name",
            "full name"
        ]

        if (
            name in name_keywords
            or name.endswith(" name")
        ):
            name_columns.append(column)
            continue

        # -------------------------------------------------
        # DATE DETECTION
        # -------------------------------------------------

        if (
            "date" in name
            or "time" in name
            or "month" in name
            or "year" in name
        ):

            converted = pd.to_datetime(
                series,
                errors="coerce"
            )

            if converted.notna().mean() >= 0.70:

                # A plain academic year such as 1, 2, 3, 4
                # should NOT become a date.
                if name in ["year", "academic year", "study year"]:
                    pass
                else:
                    date_columns.append(column)
                    continue

        # -------------------------------------------------
        # ACADEMIC / ORDINAL YEAR DETECTION
        # -------------------------------------------------

        if name in [
            "year",
            "academic year",
            "study year",
            "student year",
            "class year"
        ]:
            categorical_columns.append(column)
            continue

        # -------------------------------------------------
        # NUMERIC DETECTION
        # -------------------------------------------------

        if pd.api.types.is_numeric_dtype(series):

            # Very low unique numeric values are usually
            # categories rather than measures.
            unique_count = series.nunique(dropna=True)

            if unique_count <= 10:

                # Boolean-like columns
                if unique_count <= 2:
                    categorical_columns.append(column)

                # Small ordinal values such as Year 1-4
                else:
                    categorical_columns.append(column)

            else:
                numeric_columns.append(column)

            continue

        # -------------------------------------------------
        # CATEGORICAL / TEXT DETECTION
        # -------------------------------------------------

        categorical_columns.append(column)

    return {
        "numeric_columns": numeric_columns,
        "categorical_columns": categorical_columns,
        "date_columns": date_columns,
        "id_columns": id_columns,
        "name_columns": name_columns,
        "ignored_columns": ignored_columns
    }


# =========================================================
# DATASET ANALYSIS
# =========================================================

def analyze_dataset(df):

    classification = classify_columns(df)

    return {
        "rows": int(len(df)),
        "columns": int(len(df.columns)),

        "numeric_columns":
            classification["numeric_columns"],

        "categorical_columns":
            classification["categorical_columns"],

        "date_columns":
            classification["date_columns"],

        "id_columns":
            classification["id_columns"],

        "name_columns":
            classification["name_columns"],

        "missing_values":
            int(df.isna().sum().sum()),

        "duplicate_rows":
            int(df.duplicated().sum())
    }


# =========================================================
# IMPORTANT MEASURE DETECTION
# =========================================================

def find_measure_by_keywords(df, keywords):

    numeric_columns = df.select_dtypes(
        include=["number"]
    ).columns.tolist()

    for column in numeric_columns:

        name = normalize_column_name(column)

        for keyword in keywords:

            if keyword in name:
                return column

    return None


# =========================================================
# KPI GENERATION
# =========================================================

def generate_kpis(df):

    classification = classify_columns(df)

    numeric_columns = classification["numeric_columns"]

    kpis = []

    # -----------------------------------------------------
    # TOTAL RECORDS
    # -----------------------------------------------------

    kpis.append({
        "title": "Total Records",
        "value": f"{len(df):,}",
        "type": "count"
    })

    # -----------------------------------------------------
    # SPECIAL KPI DETECTION
    # -----------------------------------------------------

    important_kpis = []

    # CGPA / GPA
    cgpa_column = find_measure_by_keywords(
        df,
        ["cgpa", "gpa"]
    )

    if cgpa_column:
        important_kpis.append({
            "title": "Average CGPA",
            "value":
                f"{df[cgpa_column].mean():,.2f}",
            "type": "average"
        })

    # Attendance
    attendance_column = find_measure_by_keywords(
        df,
        ["attendance", "attend"]
    )

    if attendance_column:
        important_kpis.append({
            "title": "Average Attendance",
            "value":
                f"{df[attendance_column].mean():,.2f}",
            "type": "average"
        })

    # Marks / Score
    score_column = find_measure_by_keywords(
        df,
        [
            "marks",
            "score",
            "percentage",
            "percent"
        ]
    )

    if score_column:
        important_kpis.append({
            "title": f"Average {score_column}",
            "value":
                f"{df[score_column].mean():,.2f}",
            "type": "average"
        })

    # Sales
    sales_column = find_measure_by_keywords(
        df,
        [
            "sales",
            "revenue",
            "amount",
            "total sales"
        ]
    )

    if sales_column:
        important_kpis.append({
            "title": f"Total {sales_column}",
            "value":
                f"{df[sales_column].sum():,.2f}",
            "type": "sum"
        })

    # Profit
    profit_column = find_measure_by_keywords(
        df,
        ["profit", "net profit"]
    )

    if profit_column:
        important_kpis.append({
            "title": f"Total {profit_column}",
            "value":
                f"{df[profit_column].sum():,.2f}",
            "type": "sum"
        })

    # -----------------------------------------------------
    # GENERIC NUMERIC KPIs
    # -----------------------------------------------------

    used_columns = {
        cgpa_column,
        attendance_column,
        score_column,
        sales_column,
        profit_column
    }

    for column in numeric_columns:

        if column in used_columns:
            continue

        if len(important_kpis) >= 5:
            break

        average = df[column].mean()

        if pd.notna(average):

            important_kpis.append({
                "title": f"Avg {column}",
                "value": f"{average:,.2f}",
                "type": "average"
            })

    # -----------------------------------------------------
    # DATA QUALITY
    # -----------------------------------------------------

    missing = int(
        df.isna().sum().sum()
    )

    duplicates = int(
        df.duplicated().sum()
    )

    kpis.extend(important_kpis)

    kpis.append({
        "title": "Missing Values",
        "value": f"{missing:,}",
        "type": "quality"
    })

    kpis.append({
        "title": "Duplicate Rows",
        "value": f"{duplicates:,}",
        "type": "quality"
    })

    return kpis[:8]


# =========================================================
# CHART GENERATION
# =========================================================

def generate_charts(df):

    classification = classify_columns(df)

    numeric_columns = classification["numeric_columns"]
    categorical_columns = classification["categorical_columns"]
    date_columns = classification["date_columns"]

    charts = []

    # =====================================================
    # DataLens smart chart selection
    # Choose chart types from the structure of the dataset.
    # Supported: line, area, bar, pie, histogram, box,
    # scatter and heatmap.
    # =====================================================

    # -----------------------------------------------------
    # 1. DATE + NUMERIC -> LINE + AREA
    # -----------------------------------------------------
    if date_columns and numeric_columns:

        date_col = date_columns[0]
        value_col = numeric_columns[0]

        temp = df[[date_col, value_col]].copy()

        temp[date_col] = pd.to_datetime(
            temp[date_col],
            errors="coerce"
        )

        temp[value_col] = pd.to_numeric(
            temp[value_col],
            errors="coerce"
        )

        temp = temp.dropna()

        if len(temp) > 0:

            grouped = (
                temp
                .groupby(date_col)[value_col]
                .sum()
                .reset_index()
                .sort_values(date_col)
                .head(120)
            )

            x_values = [str(x) for x in grouped[date_col]]
            y_values = [clean_value(x) for x in grouped[value_col]]

            charts.append({
                "type": "line",
                "title": f"{value_col} Trend",
                "column": date_col,
                "second_column": value_col,
                "x": x_values,
                "y": y_values
            })

            charts.append({
                "type": "area",
                "title": f"{value_col} Over Time",
                "column": date_col,
                "second_column": value_col,
                "x": x_values,
                "y": y_values
            })

    # -----------------------------------------------------
    # 2. LOW-CARDINALITY CATEGORY -> BAR + PIE
    # -----------------------------------------------------
    pie_added = False

    for column in categorical_columns:

        unique_count = df[column].nunique(dropna=True)

        if unique_count == 0 or unique_count > 30:
            continue

        counts = (
            df[column]
            .fillna("Missing")
            .astype(str)
            .value_counts()
            .head(10)
        )

        if len(counts) == 0:
            continue

        charts.append({
            "type": "bar",
            "title": f"{column} Breakdown",
            "column": column,
            "labels": counts.index.tolist(),
            "values": [int(x) for x in counts.values]
        })

        # Pie charts are most readable for a small number of categories.
        if not pie_added and 2 <= len(counts) <= 8:
            charts.append({
                "type": "pie",
                "title": f"{column} Share",
                "column": column,
                "labels": counts.index.tolist(),
                "values": [int(x) for x in counts.values]
            })
            pie_added = True

        # One main categorical breakdown is enough; another type will be
        # selected from the numeric columns below.
        break

    # -----------------------------------------------------
    # 3. CATEGORY + NUMERIC -> BAR + BOX
    # -----------------------------------------------------
    if categorical_columns and numeric_columns:

        category_col = categorical_columns[0]
        numeric_col = numeric_columns[0]

        temp = df[[category_col, numeric_col]].copy()
        temp[numeric_col] = pd.to_numeric(
            temp[numeric_col],
            errors="coerce"
        )
        temp = temp.dropna()

        if len(temp) > 0:

            grouped = (
                temp
                .groupby(category_col)[numeric_col]
                .mean()
                .sort_values(ascending=False)
                .head(10)
            )

            if len(grouped) > 0:
                charts.append({
                    "type": "bar",
                    "title": f"Average {numeric_col} by {category_col}",
                    "column": category_col,
                    "second_column": numeric_col,
                    "labels": [str(x) for x in grouped.index],
                    "values": [round(float(x), 2) for x in grouped.values]
                })

            # Box plot by category when the number of groups is manageable.
            box_temp = temp.copy()
            box_temp[category_col] = box_temp[category_col].fillna("Missing").astype(str)
            valid_groups = box_temp[category_col].value_counts().head(10).index.tolist()
            box_temp = box_temp[box_temp[category_col].isin(valid_groups)]

            if len(box_temp) > 0 and box_temp[category_col].nunique() >= 2:
                charts.append({
                    "type": "box",
                    "title": f"{numeric_col} Distribution by {category_col}",
                    "column": category_col,
                    "second_column": numeric_col,
                    "x": box_temp[category_col].tolist(),
                    "y": [clean_value(x) for x in box_temp[numeric_col]]
                })

    # -----------------------------------------------------
    # 4. NUMERIC -> HISTOGRAM + BOX
    # -----------------------------------------------------
    distribution_added = 0

    for column in numeric_columns:

        series = pd.to_numeric(
            df[column],
            errors="coerce"
        ).dropna()

        if len(series) == 0:
            continue

        sample = series.head(1000)

        charts.append({
            "type": "histogram",
            "title": f"{column} Distribution",
            "column": column,
            "labels": [],
            "values": [clean_value(x) for x in sample]
        })

        if series.nunique() >= 2:
            charts.append({
                "type": "box",
                "title": f"{column} Spread",
                "column": column,
                "labels": [],
                "values": [clean_value(x) for x in sample]
            })

        distribution_added += 1
        if distribution_added >= 1:
            break

    # -----------------------------------------------------
    # 5. NUMERIC vs NUMERIC -> SCATTER
    # -----------------------------------------------------
    if len(numeric_columns) >= 2:

        first = numeric_columns[0]
        second = numeric_columns[1]

        temp = df[[first, second]].copy()
        temp[first] = pd.to_numeric(temp[first], errors="coerce")
        temp[second] = pd.to_numeric(temp[second], errors="coerce")
        temp = temp.dropna().head(1000)

        if len(temp) > 0:
            charts.append({
                "type": "scatter",
                "title": f"{first} vs {second}",
                "column": first,
                "second_column": second,
                "x": [clean_value(x) for x in temp[first]],
                "y": [clean_value(x) for x in temp[second]]
            })

    # -----------------------------------------------------
    # 6. 3+ NUMERIC COLUMNS -> CORRELATION HEATMAP
    # -----------------------------------------------------
    if len(numeric_columns) >= 3:

        corr_columns = numeric_columns[:10]
        numeric_frame = df[corr_columns].apply(
            pd.to_numeric,
            errors="coerce"
        )

        corr = numeric_frame.corr().fillna(0)

        charts.append({
            "type": "heatmap",
            "title": "Numeric Correlation Heatmap",
            "column": corr_columns[0],
            "labels": corr_columns,
            "x": corr_columns,
            "y": corr_columns,
            "z": [
                [round(float(value), 3) for value in corr.loc[row].tolist()]
                for row in corr_columns
            ]
        })

    # Keep the dashboard compact while showing different applicable types.
    # Remove duplicate chart signatures before limiting.
    unique = []
    seen = set()

    for chart in charts:
        signature = (
            chart.get("type"),
            chart.get("column"),
            chart.get("second_column"),
            chart.get("title")
        )
        if signature in seen:
            continue
        seen.add(signature)
        unique.append(chart)

    return unique[:10]

# =========================================================
# REGISTER / CREATE ACCOUNT
# =========================================================

@app.route("/register", methods=["GET", "POST"])
def register():

    if request.method == "GET":
        # REGISTER IS A SEPARATE PAGE.
        return render_template("register.html")

    data = request.get_json(silent=True) or request.form

    username = str(data.get("username", "")).strip()
    email = str(data.get("email", "")).strip().lower()
    password = str(data.get("password", ""))
    confirm_password = str(data.get("confirm_password", ""))

    def form_error(message):
        if request.is_json:
            return jsonify({"success": False, "message": message}), 400
        return render_template(
            "register.html",
            error=message,
            username=username,
            email=email
        ), 400

    if len(username) < 3:
        return form_error("Username must contain at least 3 characters.")

    if not re.match(r"^[A-Za-z0-9_.-]+$", username):
        return form_error(
            "Username can contain only letters, numbers, dots, underscores and hyphens."
        )

    if not re.match(r"^[^\s@]+@[^\s@]+\.[^\s@]+$", email):
        return form_error("Please enter a valid email address.")

    if len(password) < 6:
        return form_error("Password must contain at least 6 characters.")

    if confirm_password and password != confirm_password:
        return form_error("Passwords do not match.")

    conn = get_db()

    try:
        existing_username = conn.execute(
            "SELECT id FROM users WHERE LOWER(username) = LOWER(?)",
            (username,)
        ).fetchone()

        if existing_username:
            if request.is_json:
                return jsonify({
                    "success": False,
                    "message": "Username already exists. Please choose another username."
                }), 409
            return render_template(
                "register.html",
                error="Username already exists. Please choose another username.",
                username=username,
                email=email
            ), 409

        existing_email = conn.execute(
            "SELECT id FROM users WHERE LOWER(email) = LOWER(?)",
            (email,)
        ).fetchone()

        if existing_email:
            if request.is_json:
                return jsonify({
                    "success": False,
                    "message": "An account with this email already exists."
                }), 409
            return render_template(
                "register.html",
                error="An account with this email already exists.",
                username=username,
                email=email
            ), 409

        password_hash = generate_password_hash(password)

        conn.execute(
            "INSERT INTO users (username, email, password_hash) VALUES (?, ?, ?)",
            (username, email, password_hash)
        )
        conn.commit()

        if request.is_json:
            return jsonify({
                "success": True,
                "message": "Account created successfully. Please login.",
                "redirect": "/login"
            })

        return redirect("/login")

    except sqlite3.IntegrityError:
        conn.rollback()
        if request.is_json:
            return jsonify({
                "success": False,
                "message": "Username or email already exists."
            }), 409
        return render_template(
            "register.html",
            error="Username or email already exists.",
            username=username,
            email=email
        ), 409

    except Exception as e:
        conn.rollback()
        if request.is_json:
            return jsonify({
                "success": False,
                "message": "Unable to create account.",
                "error": str(e)
            }), 500
        return render_template(
            "register.html",
            error="Unable to create account. Please try again.",
            username=username,
            email=email
        ), 500

    finally:
        conn.close()


# =========================================================
# LOGIN
# =========================================================

@app.route("/login", methods=["GET", "POST"])
def login():

    if request.method == "GET":
        # LOGIN IS A SEPARATE PAGE.
        # Never send the browser to the dashboard just because an old
        # session cookie exists. The user can login explicitly from here.
        return render_template("login.html")

    data = request.get_json(silent=True) or request.form

    username = str(data.get("username", "")).strip()
    password = str(data.get("password", ""))

    if not username or not password:
        return jsonify({
            "success": False,
            "message": "Please enter username and password."
        }), 400

    conn = get_db()

    try:
        user = conn.execute(
            """
            SELECT id, username, email, password_hash
            FROM users
            WHERE LOWER(username) = LOWER(?)
            """,
            (username,)
        ).fetchone()
    finally:
        conn.close()

    if user is None or not check_password_hash(
        user["password_hash"], password
    ):
        return jsonify({
            "success": False,
            "message": "Invalid username or password."
        }), 401

    session.clear()
    session["user_id"] = user["id"]
    session["username"] = user["username"]
    session["email"] = user["email"] or ""

    # The separate login.html uses a normal HTML form, so send the
    # browser directly to the protected dashboard after successful login.
    if request.is_json:
        return jsonify({
            "success": True,
            "message": "Login successful.",
            "redirect": "/dashboard"
        })

    return redirect("/dashboard")


# =========================================================
# LOGOUT
# =========================================================

@app.route("/logout")
def logout():
    session.clear()
    return redirect("/login")


# =========================================================
# FIRST PAGE
# =========================================================
# The website root ALWAYS opens the separate login page.
# The dashboard is only available at /dashboard after login.

@app.route("/")
def home():
    return render_template("login.html")


# =========================================================
# MAIN DATALENS DASHBOARD
# =========================================================
# This is the existing index.html. It is protected and can only
# be reached after a successful login.

@app.route("/dashboard")
@login_required
def index():
    return render_template("index.html")


# =========================================================
# CURRENT USER
# =========================================================

@app.route("/api/me")
@login_required
def current_user():
    return jsonify({
        "success": True,
        "user": {
            "id": session["user_id"],
            "username": session["username"],
            "email": session.get("email", "")
        }
    })


# =========================================================
# SMART INSIGHTS
# =========================================================

def generate_insights(df):
    classification = classify_columns(df)
    numeric_columns = classification["numeric_columns"]
    categorical_columns = classification["categorical_columns"]
    date_columns = classification["date_columns"]

    insights = []

    rows = len(df)
    missing = int(df.isna().sum().sum())
    duplicates = int(df.duplicated().sum())

    insights.append({
        "type": "summary",
        "text": f"The dataset contains {rows:,} rows across {len(df.columns):,} columns."
    })

    if missing:
        pct = (missing / max(rows * max(len(df.columns), 1), 1)) * 100
        insights.append({
            "type": "quality",
            "text": f"There are {missing:,} missing values ({pct:.1f}% of all cells). Consider reviewing the affected columns before deeper analysis."
        })
    else:
        insights.append({
            "type": "quality",
            "text": "No missing values were detected in the uploaded dataset."
        })

    if duplicates:
        insights.append({
            "type": "quality",
            "text": f"{duplicates:,} duplicate rows were detected and may affect aggregate results."
        })

    for col in numeric_columns[:3]:
        series = pd.to_numeric(df[col], errors="coerce").dropna()
        if len(series):
            mean = float(series.mean())
            mx = float(series.max())
            mn = float(series.min())
            insights.append({
                "type": "numeric",
                "text": f"{col}: average {mean:,.2f}, minimum {mn:,.2f}, maximum {mx:,.2f}."
            })
            if len(insights) >= 6:
                break

    if len(insights) < 6 and categorical_columns:
        for col in categorical_columns[:3]:
            series = df[col].dropna().astype(str)
            if len(series):
                counts = series.value_counts()
                top = counts.index[0]
                share = counts.iloc[0] / len(series) * 100
                insights.append({
                    "type": "category",
                    "text": f"{col}: '{top}' is the most common category at {share:.1f}% of non-empty records."
                })
                if len(insights) >= 6:
                    break

    if len(insights) < 6 and date_columns and numeric_columns:
        dc = date_columns[0]
        nc = numeric_columns[0]
        temp = df[[dc, nc]].copy()
        temp[dc] = pd.to_datetime(temp[dc], errors="coerce")
        temp[nc] = pd.to_numeric(temp[nc], errors="coerce")
        temp = temp.dropna().sort_values(dc)
        if len(temp) >= 2:
            first = float(temp[nc].iloc[0])
            last = float(temp[nc].iloc[-1])
            direction = "increased" if last >= first else "decreased"
            insights.append({
                "type": "trend",
                "text": f"Across {dc}, {nc} {direction} from {first:,.2f} to {last:,.2f} in the available time range."
            })

    return insights[:6]


# =========================================================
# UPLOAD DATASET
# =========================================================

@app.route("/upload", methods=["POST"])
@login_required
def upload():

    if "file" not in request.files:
        return jsonify({
            "success": False,
            "error": "No file selected."
        }), 400

    file = request.files["file"]

    if file.filename == "":
        return jsonify({
            "success": False,
            "error": "Please select a file."
        }), 400

    extension = file.filename.lower().split(".")[-1]

    if extension not in ["csv", "xlsx", "xls"]:
        return jsonify({
            "success": False,
            "error": "Only CSV and Excel files are supported."
        }), 400

    filename = f"{uuid.uuid4().hex}.{extension}"
    filepath = os.path.join(UPLOAD_FOLDER, filename)
    file.save(filepath)

    try:
        df = load_file(filepath)

        if df.empty:
            return jsonify({
                "success": False,
                "error": "The uploaded dataset is empty."
            }), 400

        analysis = analyze_dataset(df)
        kpis = generate_kpis(df)
        charts = generate_charts(df)

        preview = df.head(10).copy()
        preview_data = []

        for _, row in preview.iterrows():
            preview_data.append({
                column: clean_value(row[column])
                for column in df.columns
            })

        records = []
        record_limit = 20000
        for _, row in df.head(record_limit).iterrows():
            records.append({column: clean_value(row[column]) for column in df.columns})

        return jsonify({
            "success": True,
            "filename": file.filename,
            "analysis": analysis,
            "kpis": kpis,
            "charts": charts,
            "columns": df.columns.tolist(),
            "preview": preview_data,
            "records": records,
            "record_limit": record_limit,
            "insights": generate_insights(df),
            "active_filters": {},
            "layout_mode": "standard",
            "user": session.get("username")
        })

    except Exception as e:

        return jsonify({
            "success": False,
            "error": str(e)
        }), 500

    finally:
        # The analysis is already returned to the browser, so the
        # temporary upload can be removed. This also prevents the
        # uploads folder from growing indefinitely.
        try:
            if os.path.exists(filepath):
                os.remove(filepath)
        except Exception:
            pass


# =========================================================
# DASHBOARD STORAGE / AUTO-SAVE
# =========================================================

def normalize_dashboard_payload(payload):
    if not isinstance(payload, dict):
        raise ValueError("Invalid dashboard data.")
    safe = {
        "filename": payload.get("filename", "Dashboard"),
        "analysis": payload.get("analysis", {}),
        "kpis": payload.get("kpis", []),
        "charts": payload.get("charts", []),
        "columns": payload.get("columns", []),
        "preview": payload.get("preview", []),
        "records": payload.get("records", [])[:20000],
        "record_limit": int(payload.get("record_limit", 20000) or 20000),
        "insights": payload.get("insights", []),
        "active_filters": payload.get("active_filters", {}),
        "layout_mode": payload.get("layout_mode", "standard")
    }
    json.dumps(safe, ensure_ascii=False)
    return safe


def dashboard_row_to_dict(row):
    data = json.loads(row["dashboard_json"])
    return {
        "id": row["id"],
        "name": row["name"],
        "filename": row["filename"],
        "share_token": row["share_token"],
        "is_shared": bool(row["is_shared"]),
        "created_at": row["created_at"],
        "updated_at": row["updated_at"],
        "data": data
    }


@app.route("/api/dashboards", methods=["GET", "POST"])
@login_required
def dashboards():
    conn = get_db()
    try:
        if request.method == "GET":
            rows = conn.execute(
                """
                SELECT id, name, filename, share_token, is_shared,
                       created_at, updated_at
                FROM saved_dashboards
                WHERE user_id = ?
                ORDER BY updated_at DESC, id DESC
                """,
                (session["user_id"],)
            ).fetchall()
            return jsonify({"success": True, "dashboards": [dict(row) for row in rows]})

        payload = request.get_json(silent=True) or {}
        dashboard = normalize_dashboard_payload(payload.get("data"))
        name = str(payload.get("name", "")).strip()
        dashboard_id = payload.get("id")
        if not name:
            name = os.path.splitext(dashboard.get("filename", "Dashboard"))[0] or "Dashboard"

        if dashboard_id:
            dashboard_id = int(dashboard_id)
            existing = conn.execute(
                "SELECT id FROM saved_dashboards WHERE id = ? AND user_id = ?",
                (dashboard_id, session["user_id"])
            ).fetchone()
            if existing:
                conn.execute(
                    """
                    UPDATE saved_dashboards
                    SET name = ?, filename = ?, dashboard_json = ?, updated_at = CURRENT_TIMESTAMP
                    WHERE id = ? AND user_id = ?
                    """,
                    (name, dashboard.get("filename", ""), json.dumps(dashboard, ensure_ascii=False), dashboard_id, session["user_id"])
                )
                conn.commit()
                return jsonify({"success": True, "id": dashboard_id, "message": "Dashboard saved."})

        cursor = conn.execute(
            """
            INSERT INTO saved_dashboards (user_id, name, filename, dashboard_json)
            VALUES (?, ?, ?, ?)
            """,
            (session["user_id"], name, dashboard.get("filename", ""), json.dumps(dashboard, ensure_ascii=False))
        )
        conn.commit()
        return jsonify({"success": True, "id": cursor.lastrowid, "message": "Dashboard auto-saved."})
    except Exception as e:
        conn.rollback()
        return jsonify({"success": False, "message": str(e)}), 400
    finally:
        conn.close()


@app.route("/api/dashboards/<int:dashboard_id>")
@login_required
def get_saved_dashboard(dashboard_id):
    conn = get_db()
    try:
        row = conn.execute(
            "SELECT * FROM saved_dashboards WHERE id = ? AND user_id = ?",
            (dashboard_id, session["user_id"])
        ).fetchone()
        if row is None:
            return jsonify({"success": False, "message": "Dashboard not found."}), 404
        return jsonify({"success": True, **dashboard_row_to_dict(row)})
    finally:
        conn.close()


@app.route("/api/dashboards/<int:dashboard_id>/filter", methods=["POST"])
@login_required
def filter_saved_dashboard(dashboard_id):
    conn = get_db()
    try:
        row = conn.execute(
            "SELECT dashboard_json FROM saved_dashboards WHERE id = ? AND user_id = ?",
            (dashboard_id, session["user_id"])
        ).fetchone()
        if row is None:
            return jsonify({"success": False, "message": "Dashboard not found."}), 404

        stored = json.loads(row["dashboard_json"])
        records = stored.get("records") or []
        if not records:
            return jsonify({
                "success": False,
                "message": "This saved dashboard does not contain interactive source rows. Generate it again to enable Smart Filters."
            }), 400

        request_data = request.get_json(silent=True) or {}
        filters = request_data.get("filters") or {}
        numeric_ranges = request_data.get("numeric_ranges") or {}

        df = pd.DataFrame(records)
        original_count = len(df)

        for column, value in filters.items():
            if column not in df.columns or value in (None, "", "All"):
                continue
            df = df[df[column].fillna("Missing").astype(str) == str(value)]

        for column, bounds in numeric_ranges.items():
            if column not in df.columns or not isinstance(bounds, (list, tuple)) or len(bounds) != 2:
                continue
            numeric = pd.to_numeric(df[column], errors="coerce")
            mask = numeric.notna()
            low = bounds[0]
            high = bounds[1]
            try:
                if low not in (None, ""):
                    mask &= numeric >= float(low)
                if high not in (None, ""):
                    mask &= numeric <= float(high)
            except (TypeError, ValueError):
                pass
            df = df.loc[mask]

        analysis = analyze_dataset(df)
        kpis = generate_kpis(df)
        charts = generate_charts(df)
        preview = []
        for _, row_data in df.head(10).iterrows():
            preview.append({column: clean_value(row_data[column]) for column in df.columns})

        return jsonify({
            "success": True,
            "analysis": analysis,
            "kpis": kpis,
            "charts": charts,
            "preview": preview,
            "filtered_rows": int(len(df)),
            "original_rows": int(original_count),
            "active_filters": {
                "categorical": filters,
                "numeric_ranges": numeric_ranges
            },
            "insights": generate_insights(df)
        })
    except Exception as e:
        return jsonify({"success": False, "message": str(e)}), 400
    finally:
        conn.close()


@app.route("/api/dashboards/<int:dashboard_id>", methods=["DELETE"])
@login_required
def delete_saved_dashboard(dashboard_id):
    conn = get_db()
    try:
        cursor = conn.execute(
            "DELETE FROM saved_dashboards WHERE id = ? AND user_id = ?",
            (dashboard_id, session["user_id"])
        )
        conn.commit()
        if cursor.rowcount == 0:
            return jsonify({"success": False, "message": "Dashboard not found."}), 404
        return jsonify({"success": True, "message": "Dashboard deleted."})
    finally:
        conn.close()


# =========================================================
# DOWNLOAD DASHBOARD AS HTML
# =========================================================

def safe_json_for_script(value):
    return json.dumps(value, ensure_ascii=False, separators=(",", ":")).replace("<", "\\u003c").replace(">", "\\u003e").replace("&", "\\u0026")


def make_dashboard_html(title, data, readonly=False):
    """Create a standalone dashboard export with all generated chart types.

    The export is intentionally self-contained except for Plotly JS loaded from CDN.
    It renders every chart returned by DataLens (up to the backend's current limit),
    uses a dense responsive grid, and calls Plotly.resize after layout so charts do not
    collapse into blank cards.
    """
    payload = safe_json_for_script(data)
    import html as _html
    title_safe = _html.escape(str(title))
    readonly_note = "<div class='shared-note'>Read-only shared dashboard</div>" if readonly else ""
    return f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>{title_safe} | DataLens</title>
<script src="https://cdn.plot.ly/plotly-2.35.2.min.js"></script>
<style>
*{{box-sizing:border-box}}
html,body{{width:100%;min-height:100%;margin:0}}
body{{font-family:Segoe UI,Inter,Arial,sans-serif;background:radial-gradient(circle at 8% 0%,rgba(37,99,235,.10),transparent 25%),radial-gradient(circle at 92% 4%,rgba(6,182,212,.10),transparent 22%),#f4f7fb;color:#0f172a;overflow:auto}}
.page{{width:100%;min-height:100vh;padding:12px 14px;display:grid;grid-template-rows:auto auto auto 1fr auto;gap:8px}}
.topbar{{display:flex;align-items:center;justify-content:space-between;gap:12px;padding:9px 12px;background:rgba(255,255,255,.94);border:1px solid #dbe4f0;border-radius:16px;box-shadow:0 8px 22px rgba(15,23,42,.05)}}
.brand-wrap{{display:flex;align-items:center;gap:9px;min-width:0}}
.brand-mark{{width:32px;height:32px;border-radius:10px;display:grid;place-items:center;background:linear-gradient(135deg,#2563eb,#06b6d4);color:#fff;font-weight:900;font-size:17px}}
.brand{{font-size:18px;font-weight:900;letter-spacing:-.4px}}
.brand-sub{{font-size:9px;color:#64748b;font-weight:700}}
.shared-note{{font-size:8px;color:#64748b;margin-top:2px}}
.top-actions{{display:flex;align-items:center;gap:7px}}
.badge{{padding:6px 9px;border-radius:999px;background:#ecfdf5;color:#047857;border:1px solid #86efac;font-weight:800;font-size:9px;white-space:nowrap}}
.preview-btn{{padding:6px 9px;border:1px solid #cbd5e1;background:#fff;border-radius:9px;color:#334155;font-size:9px;font-weight:800;cursor:pointer}}
.hero{{padding:0 2px;min-width:0}}
h1{{margin:0;font-size:22px;line-height:1.1;letter-spacing:-.7px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}}
.meta{{color:#64748b;margin-top:3px;font-size:9px;font-weight:650;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}}
.kpis{{display:grid;grid-template-columns:repeat(8,minmax(0,1fr));gap:8px;min-width:0}}
.kpi{{position:relative;min-width:0;background:#fff;border:1px solid #dbe4f0;border-radius:13px;padding:9px 10px;box-shadow:0 6px 16px rgba(15,23,42,.045);overflow:hidden}}
.kpi::before{{content:"";position:absolute;left:0;top:0;bottom:0;width:3px;background:linear-gradient(180deg,#2563eb,#06b6d4)}}
.kpi .l{{font-size:8px;color:#64748b;font-weight:850;text-transform:uppercase;letter-spacing:.35px;padding-left:2px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}}
.kpi .v{{font-size:17px;font-weight:900;margin-top:4px;letter-spacing:-.3px;padding-left:2px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}}
.insights{{display:flex;align-items:center;gap:7px;min-width:0;background:linear-gradient(135deg,#eef6ff,#f7f5ff);border:1px solid #dbeafe;border-radius:13px;padding:6px 8px;overflow:hidden}}
.insight-title{{font-size:9px;font-weight:900;color:#334155;white-space:nowrap;display:flex;align-items:center;gap:5px}}
.insight-title span{{width:19px;height:19px;border-radius:7px;display:grid;place-items:center;background:#e0ecff;color:#1d4ed8}}
.insight-list{{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:6px;min-width:0;flex:1}}
.insight{{min-width:0;padding:5px 7px;background:rgba(255,255,255,.86);border:1px solid #dbeafe;border-radius:9px;color:#334155;font-size:8px;line-height:1.35;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}}
.charts{{display:grid;grid-template-columns:repeat(5,minmax(0,1fr));grid-template-rows:repeat(2,minmax(0,1fr));gap:8px;min-height:0}}
.card{{min-width:0;min-height:0;background:#fff;border:1px solid #dbe4f0;border-radius:15px;padding:8px 9px;box-shadow:0 6px 16px rgba(15,23,42,.045);display:grid;grid-template-rows:auto minmax(0,1fr);overflow:hidden}}
.card-head{{display:flex;align-items:center;justify-content:space-between;gap:6px;margin-bottom:2px;min-width:0}}
.card h3{{margin:0;font-size:10px;line-height:1.2;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}}
.chart-type{{font-size:7px;font-weight:900;letter-spacing:.4px;padding:3px 5px;border-radius:999px;background:#eef2ff;color:#4f46e5;white-space:nowrap;flex-shrink:0}}
.chart{{width:100%;height:100%;min-height:150px}}
.footer{{font-size:7px;color:#94a3b8;text-align:right;white-space:nowrap}}
.modal{{display:none;position:fixed;inset:0;background:rgba(15,23,42,.48);z-index:50;align-items:center;justify-content:center;padding:18px}}
.modal.open{{display:flex}}
.modal-card{{width:min(1100px,96vw);max-height:88vh;background:#fff;border-radius:18px;border:1px solid #dbe4f0;padding:15px;box-shadow:0 22px 60px rgba(15,23,42,.25);display:flex;flex-direction:column;gap:9px}}
.modal-head{{display:flex;align-items:center;justify-content:space-between;gap:10px}}
.modal-head h2{{margin:0;font-size:15px}}
.close{{border:1px solid #cbd5e1;background:#fff;border-radius:9px;padding:6px 9px;font-weight:800;cursor:pointer}}
.table-scroll{{overflow:auto;border:1px solid #e2e8f0;border-radius:12px;max-height:68vh}}
table{{width:100%;border-collapse:collapse;font-size:10px}}
th,td{{padding:7px 9px;border-bottom:1px solid #e2e8f0;text-align:left;white-space:nowrap}}
th{{position:sticky;top:0;background:#eff6ff;color:#1d4ed8;z-index:1}}
tbody tr:nth-child(even) td{{background:#fbfdff}}
@media(max-width:1400px){{.charts{{grid-template-columns:repeat(4,minmax(0,1fr));grid-template-rows:none}}.chart{{min-height:180px}}}}
@media(max-width:1050px){{.kpis{{grid-template-columns:repeat(4,minmax(0,1fr))}}.charts{{grid-template-columns:repeat(2,minmax(0,1fr))}}.insight-list{{grid-template-columns:1fr 1fr}}}}
@media(max-width:700px){{.kpis{{grid-template-columns:repeat(2,minmax(0,1fr))}}.charts{{grid-template-columns:1fr}}.card{{min-height:230px}}.insights{{align-items:flex-start}}.insight-list{{grid-template-columns:1fr}}}}
</style>
</head>
<body>
<div class="page">
  <div class="topbar">
    <div class="brand-wrap">
      <div class="brand-mark">⌁</div>
      <div>
        <div class="brand">DataLens</div>
        <div class="brand-sub">Interactive analytics dashboard</div>
        {readonly_note}
      </div>
    </div>
    <div class="top-actions">
      <button class="preview-btn" id="previewBtn" type="button">View Data</button>
      <div class="badge">● Analytics Ready</div>
    </div>
  </div>

  <section class="hero">
    <h1 id="title">{title_safe}</h1>
    <div class="meta" id="meta"></div>
  </section>

  <div>
    <div class="kpis" id="kpis"></div>
    <div class="insights" id="insightsShell">
      <div class="insight-title"><span>✦</span>Insights</div>
      <div class="insight-list" id="insights"></div>
    </div>
  </div>

  <div class="charts" id="charts"></div>

  <div class="footer">Generated by DataLens • Standalone dashboard export</div>
</div>

<div class="modal" id="previewModal" aria-hidden="true">
  <div class="modal-card">
    <div class="modal-head">
      <h2>Data Preview</h2>
      <button class="close" id="closePreview" type="button">Close</button>
    </div>
    <div class="table-scroll"><div id="preview"></div></div>
  </div>
</div>

<script>
const DATA = {payload};
function esc(v){{return String(v??'').replace(/[&<>\"]/g,m=>({{'&':'&amp;','<':'&lt;','>':'&gt;','\"':'&quot;'}}[m]))}}
const analysis = DATA.analysis || {{}};
document.getElementById('meta').textContent = `${{analysis.rows ?? 0}} rows • ${{analysis.columns ?? (DATA.columns||[]).length}} columns • ${{DATA.filename||''}}`;

const k=document.getElementById('kpis');
(DATA.kpis||[]).slice(0,8).forEach(x=>{{
  const d=document.createElement('div');d.className='kpi';
  d.innerHTML=`<div class="l">${{esc(x.title)}}</div><div class="v">${{esc(x.value)}}</div>`;
  k.appendChild(d);
}});

const ins=document.getElementById('insights');
(DATA.insights||[]).slice(0,3).forEach(x=>{{
  const d=document.createElement('div');d.className='insight';d.textContent=x.text||String(x);ins.appendChild(d);
}});
if(!ins.children.length){{
  const d=document.createElement('div');d.className='insight';d.textContent='No additional insights were detected for this dataset.';ins.appendChild(d);
}}

function baseLayout(type){{return {{
  margin:{{l:type==='heatmap'?38:28,r:type==='heatmap'?28:8,t:4,b:type==='pie'||type==='donut'?8:28}},
  paper_bgcolor:'transparent',plot_bgcolor:'transparent',
  font:{{family:'Segoe UI,Arial',size:8,color:'#64748b'}},
  showlegend:type==='pie'||type==='donut',
  legend:{{font:{{size:8}},orientation:'h',y:-.18}},
  xaxis:{{automargin:true,tickfont:{{size:7}}}},
  yaxis:{{automargin:true,tickfont:{{size:7}}}}
}}}}

function draw(el,c){{
  const type=c.type||'';
  let trace=[];
  if(type==='bar') trace=[{{x:c.labels||[],y:(c.values||[]).map(Number),type:'bar',marker:{{color:'#2563eb'}}}}];
  else if(type==='line') trace=[{{x:c.x||[],y:(c.y||[]).map(Number),type:'scatter',mode:'lines+markers',line:{{width:2,color:'#2563eb',shape:'spline'}},marker:{{color:'#06b6d4',size:3}}}}];
  else if(type==='area') trace=[{{x:c.x||[],y:(c.y||[]).map(Number),type:'scatter',mode:'lines',fill:'tozeroy',line:{{width:2,color:'#2563eb'}},fillcolor:'rgba(37,99,235,.18)'}}];
  else if(type==='pie'||type==='donut') trace=[{{labels:c.labels||[],values:(c.values||[]).map(Number),type:'pie',hole:type==='donut'?.55:0,textinfo:'percent',textposition:'inside'}}];
  else if(type==='histogram') trace=[{{x:(c.values||[]).map(Number),type:'histogram',marker:{{color:'#06b6d4'}},nbinsx:20}}];
  else if(type==='box') {{
    const xs=Array.isArray(c.x)?c.x:[];
    const ys=(Array.isArray(c.y)?c.y:(c.values||[])).map(Number).filter(Number.isFinite);
    trace=[{{x:xs.length===ys.length?xs:undefined,y:ys,type:'box',boxpoints:'outliers',jitter:.25,pointpos:0,marker:{{size:3,color:'#2563eb'}},line:{{color:'#2563eb'}}}}];
  }}
  else if(type==='scatter') trace=[{{x:(c.x||[]).map(Number),y:(c.y||[]).map(Number),type:'scatter',mode:'markers',marker:{{size:5,color:'#0284c7',opacity:.7}}}}];
  else if(type==='heatmap') trace=[{{x:c.x||[],y:c.y||[],z:c.z||[],type:'heatmap',colorscale:'Blues',zmin:-1,zmax:1,colorbar:{{title:'r',thickness:8,len:.75}}}}];
  else if(type==='treemap') trace=[{{labels:c.labels||[],parents:c.parents||[],values:(c.values||[]).map(Number),type:'treemap'}}];
  else if(type==='funnel') trace=[{{y:c.labels||[],x:(c.values||[]).map(Number),type:'funnel'}}];
  else trace=[{{x:c.labels||c.x||[],y:(c.values||c.y||[]).map(Number),type:'bar',marker:{{color:'#2563eb'}}}}];

  const layout=baseLayout(type);
  if(type==='heatmap') layout.xaxis={{side:'bottom',tickangle:-35,tickfont:{{size:7}},automargin:true}};
  Plotly.newPlot(el,trace,layout,{{displayModeBar:false,responsive:true,scrollZoom:false}})
    .then(()=>Plotly.Plots.resize(el));
}}

const cg=document.getElementById('charts');
const charts=DATA.charts||[];
charts.forEach((c,i)=>{{
  const card=document.createElement('div');card.className='card';
  card.innerHTML=`<div class="card-head"><h3>${{esc(c.title||'Data Insight')}}</h3><span class="chart-type">${{esc((c.type||'chart').toUpperCase())}}</span></div><div class="chart" id="chart${{i}}"></div>`;
  cg.appendChild(card);
}});

function renderAllCharts(){{
  charts.forEach((c,i)=>draw(document.getElementById('chart'+i),c));
  setTimeout(()=>document.querySelectorAll('.chart').forEach(el=>Plotly.Plots.resize(el)),120);
  setTimeout(()=>document.querySelectorAll('.chart').forEach(el=>Plotly.Plots.resize(el)),500);
}}

if(document.readyState==='loading'){{document.addEventListener('DOMContentLoaded',renderAllCharts);}}else{{renderAllCharts();}}
window.addEventListener('resize',()=>document.querySelectorAll('.chart').forEach(el=>Plotly.Plots.resize(el)));

const pv=document.getElementById('preview'),rows=DATA.preview||[];
if(rows.length){{
  const cols=Object.keys(rows[0]);
  const t=document.createElement('table');
  t.innerHTML='<thead><tr>'+cols.map(c=>`<th>${{esc(c)}}</th>`).join('')+'</tr></thead><tbody>'+rows.map(r=>'<tr>'+cols.map(c=>`<td>${{esc(r[c])}}</td>`).join('')+'</tr>').join('')+'</tbody>';
  pv.appendChild(t);
}}else{{pv.textContent='No preview available.';}}

const modal=document.getElementById('previewModal');
document.getElementById('previewBtn').addEventListener('click',()=>{{modal.classList.add('open');modal.setAttribute('aria-hidden','false')}});
document.getElementById('closePreview').addEventListener('click',()=>{{modal.classList.remove('open');modal.setAttribute('aria-hidden','true')}});
modal.addEventListener('click',e=>{{if(e.target===modal){{modal.classList.remove('open');modal.setAttribute('aria-hidden','true')}}}});
</script>
</body>
</html>"""


@app.route("/download-dashboard/<int:dashboard_id>")
@login_required
def download_dashboard(dashboard_id):
    conn = get_db()
    try:
        row = conn.execute(
            "SELECT name, dashboard_json FROM saved_dashboards WHERE id = ? AND user_id = ?",
            (dashboard_id, session["user_id"])
        ).fetchone()
        if row is None:
            return jsonify({"success": False, "message": "Dashboard not found."}), 404
        data = json.loads(row["dashboard_json"])
        html = make_dashboard_html(row["name"], data)
        filename = re.sub(r"[^A-Za-z0-9_-]+", "_", row["name"]).strip("_") or "datalens_dashboard"
        return Response(html, mimetype="text/html", headers={"Content-Disposition": f'attachment; filename="{filename}.html"'})
    finally:
        conn.close()


# =========================================================
# PDF EXPORT
# =========================================================

@app.route("/download-dashboard-pdf/<int:dashboard_id>")
@login_required
def download_dashboard_pdf(dashboard_id):
    conn = get_db()
    try:
        row = conn.execute(
            "SELECT name, dashboard_json FROM saved_dashboards WHERE id = ? AND user_id = ?",
            (dashboard_id, session["user_id"])
        ).fetchone()
        if row is None:
            return jsonify({"success": False, "message": "Dashboard not found."}), 404

        data = json.loads(row["dashboard_json"])
        filename = re.sub(r"[^A-Za-z0-9_-]+", "_", row["name"]).strip("_") or "datalens_dashboard"
        pdf_path = os.path.join(UPLOAD_FOLDER, f"{uuid.uuid4().hex}.pdf")
        c = canvas.Canvas(pdf_path, pagesize=A4)
        width, height = A4
        y = height - 48

        c.setFont("Helvetica-Bold", 20)
        c.setFillColor(colors.HexColor("#1d4ed8"))
        c.drawString(42, y, "DataLens")
        y -= 28
        c.setFont("Helvetica-Bold", 15)
        c.setFillColor(colors.black)
        c.drawString(42, y, str(row["name"])[:90])
        y -= 20
        c.setFont("Helvetica", 9)
        c.setFillColor(colors.HexColor("#64748b"))
        analysis = data.get("analysis", {})
        c.drawString(42, y, f"{analysis.get('rows', 0):,} rows • {analysis.get('columns', 0):,} columns • {data.get('filename', '')}")
        y -= 28

        c.setFont("Helvetica-Bold", 12)
        c.setFillColor(colors.black)
        c.drawString(42, y, "Key Performance Indicators")
        y -= 18
        c.setFont("Helvetica", 10)
        for kpi in data.get("kpis", [])[:8]:
            c.drawString(52, y, f"{str(kpi.get('title',''))[:40]}: {str(kpi.get('value',''))[:40]}")
            y -= 15
            if y < 90:
                c.showPage(); y = height - 48

        y -= 8
        c.setFont("Helvetica-Bold", 12)
        c.drawString(42, y, "DataLens Insights")
        y -= 18
        c.setFont("Helvetica", 9.5)
        for insight in data.get("insights", [])[:6]:
            text = "• " + str(insight.get("text", ""))
            words = text.split()
            line = ""
            for word in words:
                candidate = (line + " " + word).strip()
                if c.stringWidth(candidate, "Helvetica", 9.5) > width - 94:
                    c.drawString(52, y, line)
                    y -= 13
                    line = word
                    if y < 90:
                        c.showPage(); y = height - 48
                        c.setFont("Helvetica", 9.5)
                else:
                    line = candidate
            if line:
                c.drawString(52, y, line); y -= 15

        c.setFont("Helvetica-Bold", 11)
        if y < 100:
            c.showPage(); y = height - 48
        c.drawString(42, y, "Data Preview")
        y -= 18
        preview = data.get("preview", [])[:8]
        columns = data.get("columns", [])[:7]
        c.setFont("Helvetica", 7.5)
        if columns:
            header = " | ".join(str(col)[:14] for col in columns)
            c.drawString(42, y, header[:125]); y -= 12
            for row_data in preview:
                values = " | ".join(str(row_data.get(col, ""))[:14] for col in columns)
                c.drawString(42, y, values[:125])
                y -= 11
                if y < 50:
                    c.showPage(); y = height - 48; c.setFont("Helvetica", 7.5)

        c.save()
        return Response(
            open(pdf_path, "rb"),
            mimetype="application/pdf",
            headers={"Content-Disposition": f'attachment; filename="{filename}.pdf"'}
        )
    finally:
        conn.close()


# =========================================================
# SHARE DASHBOARD
# =========================================================

@app.route("/api/dashboards/<int:dashboard_id>/share", methods=["POST"])
@login_required
def share_dashboard(dashboard_id):
    conn = get_db()
    try:
        row = conn.execute(
            "SELECT id, share_token FROM saved_dashboards WHERE id = ? AND user_id = ?",
            (dashboard_id, session["user_id"])
        ).fetchone()
        if row is None:
            return jsonify({"success": False, "message": "Dashboard not found."}), 404
        token = row["share_token"] or uuid.uuid4().hex + uuid.uuid4().hex
        conn.execute(
            "UPDATE saved_dashboards SET share_token = ?, is_shared = 1, updated_at = CURRENT_TIMESTAMP WHERE id = ? AND user_id = ?",
            (token, dashboard_id, session["user_id"])
        )
        conn.commit()
        return jsonify({
            "success": True,
            "share_url": request.host_url.rstrip("/") + "/share/" + token,
            "message": "Share link created."
        })
    finally:
        conn.close()


@app.route("/share/<token>")
def public_shared_dashboard(token):
    conn = get_db()
    try:
        row = conn.execute(
            "SELECT name, dashboard_json FROM saved_dashboards WHERE share_token = ? AND is_shared = 1",
            (token,)
        ).fetchone()
        if row is None:
            return "<h2 style='font-family:Segoe UI,Arial'>This dashboard link is invalid or no longer available.</h2>", 404
        data = json.loads(row["dashboard_json"])
        return Response(make_dashboard_html(row["name"], data, readonly=True), mimetype="text/html")
    finally:
        conn.close()



# =========================================================
# DATA ENGINEERING MODE
# =========================================================


def _engineering_records_to_df(records):
    if not isinstance(records, list) or not records:
        return pd.DataFrame()
    safe_records = [r for r in records if isinstance(r, dict)]
    if not safe_records:
        return pd.DataFrame()
    return pd.DataFrame(safe_records)


def _json_safe_record(row):
    return {str(k): clean_value(v) for k, v in row.items()}


def _class_role_map(df):
    classification = classify_columns(df)
    role_map = {}
    for col in classification.get("id_columns", []):
        role_map[col] = "ID"
    for col in classification.get("name_columns", []):
        role_map[col] = "Name"
    for col in classification.get("date_columns", []):
        role_map[col] = "Date"
    for col in classification.get("numeric_columns", []):
        role_map[col] = "Numeric"
    for col in classification.get("categorical_columns", []):
        role_map[col] = "Categorical"
    for col in classification.get("ignored_columns", []):
        role_map[col] = "Ignored"
    return role_map


def _schema_snapshot(df):
    roles = _class_role_map(df)
    rows = []
    for col in df.columns:
        series = df[col]
        rows.append({
            "name": str(col),
            "dtype": str(series.dtype),
            "role": roles.get(col, "Other"),
            "nulls": int(series.isna().sum()),
            "unique": int(series.nunique(dropna=True)),
        })
    return rows


def _quality_analysis(df):
    classification = classify_columns(df)
    rows = len(df)
    columns = len(df.columns)
    null_cells = int(df.isna().sum().sum())
    duplicate_rows = int(df.duplicated().sum())
    issue_rows = {}

    def add_issue(mask, reason):
        if mask is None:
            return
        try:
            idxs = df.index[mask].tolist()
        except Exception:
            return
        for idx in idxs:
            issue_rows.setdefault(int(idx) if isinstance(idx, (int, np.integer)) else str(idx), []).append(reason)

    for col in classification.get("id_columns", []):
        add_issue(df[col].isna(), f"{col} is null")
        try:
            dup_mask = df[col].notna() & df[col].duplicated(keep=False)
            add_issue(dup_mask, f"Duplicate {col}")
        except Exception:
            pass

    for col in classification.get("numeric_columns", []):
        name = normalize_column_name(col)
        numeric = pd.to_numeric(df[col], errors="coerce")
        if any(k in name for k in ["cgpa", "gpa"]):
            add_issue(numeric.notna() & ((numeric < 0) | (numeric > 10)), f"{col} outside 0-10")
        elif any(k in name for k in ["attendance", "percentage", "percent", "score", "mark"]):
            add_issue(numeric.notna() & ((numeric < 0) | (numeric > 100)), f"{col} outside 0-100")

    invalid_rows = len(issue_rows)
    total_cells = max(rows * max(columns, 1), 1)
    completeness = 1.0 - (null_cells / total_cells)
    uniqueness = 1.0 - (duplicate_rows / max(rows, 1))
    validity = 1.0 - (invalid_rows / max(rows, 1))
    quality_score = round(max(0.0, min(100.0, 100.0 * (0.4 * completeness + 0.3 * uniqueness + 0.3 * validity))), 1)

    rejected = []
    for idx, reasons in list(issue_rows.items())[:120]:
        try:
            row = df.loc[idx]
            record = _json_safe_record(row)
        except Exception:
            record = {}
        rejected.append({
            "row": idx,
            "reasons": reasons,
            "record": record,
        })

    return {
        "total_rows": rows,
        "total_columns": columns,
        "null_cells": null_cells,
        "duplicate_rows": duplicate_rows,
        "invalid_rows": invalid_rows,
        "valid_rows": max(0, rows - invalid_rows),
        "quality_score": quality_score,
        "rejected_records": rejected,
        "rules": [
            "ID columns must not be null",
            "ID duplicates are flagged",
            "CGPA / GPA must be between 0 and 10",
            "Attendance / score / percentage fields must be between 0 and 100",
        ],
    }


def _silver_transform(df):
    if df.empty:
        return df.copy(), []
    out = df.copy()
    transformations = []

    original_cols = list(out.columns)
    normalized_cols = []
    seen = set()
    for col in original_cols:
        base = re.sub(r"[^A-Za-z0-9_]+", "_", str(col).strip()).strip("_") or "column"
        base = base.lower()
        candidate = base
        counter = 2
        while candidate in seen:
            candidate = f"{base}_{counter}"
            counter += 1
        seen.add(candidate)
        normalized_cols.append(candidate)
    if normalized_cols != [str(c) for c in original_cols]:
        transformations.append("Normalize column names")
    out.columns = normalized_cols

    before_dupes = len(out)
    out = out.drop_duplicates().reset_index(drop=True)
    if len(out) != before_dupes:
        transformations.append(f"Remove {before_dupes - len(out):,} duplicate rows")

    for col in out.columns:
        if pd.api.types.is_object_dtype(out[col]):
            cleaned = out[col].apply(lambda x: x.strip() if isinstance(x, str) else x)
            if not cleaned.equals(out[col]):
                transformations.append(f"Trim text in {col}")
            out[col] = cleaned

    for col in out.columns:
        converted = pd.to_numeric(out[col], errors="coerce")
        original_non_null = int(out[col].notna().sum())
        converted_non_null = int(converted.notna().sum())
        if original_non_null and converted_non_null == original_non_null and not pd.api.types.is_numeric_dtype(out[col]):
            out[col] = converted
            transformations.append(f"Cast {col} to numeric")

    return out, transformations


def _gold_summary(df):
    if df.empty:
        return {"columns": [], "rows": [], "description": "No rows available."}
    classification = classify_columns(df)
    cats = classification.get("categorical_columns", [])
    nums = classification.get("numeric_columns", [])
    if cats and nums:
        cat = cats[0]
        num = nums[0]
        temp = df[[cat, num]].copy()
        temp[num] = pd.to_numeric(temp[num], errors="coerce")
        temp = temp.dropna(subset=[num])
        grouped = temp.groupby(cat, dropna=False)[num].agg(["count", "mean", "sum"]).reset_index().head(12)
        cols = [str(cat), "count", f"avg_{num}", f"sum_{num}"]
        rows = []
        for _, r in grouped.iterrows():
            rows.append({
                str(cat): clean_value(r[cat]),
                "count": int(r["count"]),
                f"avg_{num}": round(float(r["mean"]), 3),
                f"sum_{num}": round(float(r["sum"]), 3),
            })
        return {"columns": cols, "rows": rows, "group_by": str(cat), "measure": str(num), "description": f"Aggregated {num} by {cat}."}
    if nums:
        num = nums[0]
        s = pd.to_numeric(df[num], errors="coerce").dropna()
        return {
            "columns": ["metric", "value"],
            "rows": [
                {"metric": "count", "value": int(len(s))},
                {"metric": "average", "value": round(float(s.mean()), 3) if len(s) else 0},
                {"metric": "minimum", "value": round(float(s.min()), 3) if len(s) else 0},
                {"metric": "maximum", "value": round(float(s.max()), 3) if len(s) else 0},
            ],
            "group_by": None,
            "measure": str(num),
            "description": f"Summary metrics for {num}.",
        }
    cat = classification.get("categorical_columns", [None])[0]
    if cat:
        counts = df[cat].fillna("Missing").astype(str).value_counts().head(12)
        return {"columns": [str(cat), "count"], "rows": [{str(cat): str(k), "count": int(v)} for k, v in counts.items()], "group_by": str(cat), "measure": None, "description": f"Category counts for {cat}."}
    return {"columns": [], "rows": [], "description": "No suitable aggregation found."}


def _compare_schemas(previous_df, current_df):
    prev_map = {str(c): str(previous_df[c].dtype) for c in previous_df.columns}
    curr_map = {str(c): str(current_df[c].dtype) for c in current_df.columns}
    added = [c for c in curr_map if c not in prev_map]
    removed = [c for c in prev_map if c not in curr_map]
    changed = [{"column": c, "previous": prev_map[c], "current": curr_map[c]} for c in curr_map if c in prev_map and curr_map[c] != prev_map[c]]
    return {"added": added, "removed": removed, "changed": changed, "has_drift": bool(added or removed or changed)}


def _row_key_column(df):
    classification = classify_columns(df)
    if classification.get("id_columns"):
        return classification["id_columns"][0]
    return str(df.columns[0]) if len(df.columns) else None


def _normalize_compare_value(v):
    if pd.isna(v):
        return None
    if isinstance(v, (np.integer, np.floating)):
        return v.item()
    if isinstance(v, (pd.Timestamp, datetime)):
        return str(v)
    return str(v).strip() if isinstance(v, str) else v


def _cdc_compare(previous_df, current_df):
    key = _row_key_column(current_df) or _row_key_column(previous_df)
    if not key or key not in current_df.columns or key not in previous_df.columns:
        return {"key_column": key, "inserted": 0, "updated": 0, "deleted": 0, "unchanged": 0, "changes": [], "message": "No common key column was detected."}

    def mapping(df):
        m = {}
        for _, row in df.iterrows():
            k = _normalize_compare_value(row.get(key))
            if k is None:
                continue
            record = {str(c): _normalize_compare_value(row[c]) for c in df.columns}
            m[str(k)] = record
        return m

    old = mapping(previous_df)
    new = mapping(current_df)
    inserted_keys = [k for k in new if k not in old]
    deleted_keys = [k for k in old if k not in new]
    updated_keys = []
    unchanged_keys = []
    for k in new:
        if k in old:
            if new[k] != old[k]:
                updated_keys.append(k)
            else:
                unchanged_keys.append(k)

    changes = []
    for k in inserted_keys[:80]:
        changes.append({"key": k, "operation": "INSERT", "before": None, "after": new[k]})
    for k in updated_keys[:80]:
        changes.append({"key": k, "operation": "UPDATE", "before": old[k], "after": new[k]})
    for k in deleted_keys[:80]:
        changes.append({"key": k, "operation": "DELETE", "before": old[k], "after": None})

    return {
        "key_column": key,
        "inserted": len(inserted_keys),
        "updated": len(updated_keys),
        "deleted": len(deleted_keys),
        "unchanged": len(unchanged_keys),
        "changes": changes,
        "message": f"Compared datasets using {key}."
    }


def _lineage(df, gold):
    columns = [str(c) for c in df.columns[:18]]
    source_nodes = [f"source.{c}" for c in columns]
    bronze_nodes = [f"bronze.{re.sub(r'[^A-Za-z0-9_]+', '_', c).strip('_').lower() or 'column'}" for c in columns]
    silver_nodes = [f"silver.{re.sub(r'[^A-Za-z0-9_]+', '_', c).strip('_').lower() or 'column'}" for c in columns]
    gold_nodes = [f"gold.{c}" for c in gold.get("columns", [])[:8]]
    edges = []
    for s, b, si in zip(source_nodes, bronze_nodes, silver_nodes):
        edges.append([s, b])
        edges.append([b, si])
    for g in gold_nodes:
        if columns:
            edges.append([silver_nodes[0], g])
    if gold_nodes:
        edges.append([gold_nodes[0], "dashboard.metrics"])
    elif silver_nodes:
        edges.append([silver_nodes[0], "dashboard.metrics"])
    return {"nodes": source_nodes + bronze_nodes + silver_nodes + gold_nodes + ["dashboard.metrics"], "edges": edges}


def _performance(df, duration_ms):
    rows = len(df)
    cols = len(df.columns)
    partitions = max(1, math.ceil(rows / 1000))
    shuffle_ops = []
    classification = classify_columns(df)
    if classification.get("categorical_columns") and classification.get("numeric_columns"):
        shuffle_ops.append("groupBy aggregation")
    if rows > 5000:
        shuffle_ops.append("wide row processing")
    recommendations = []
    if rows > 5000:
        recommendations.append("Process large files incrementally rather than re-reading all records.")
    if classification.get("date_columns"):
        recommendations.append("Partition time-series workloads by date when operating at production scale.")
    if classification.get("categorical_columns") and classification.get("numeric_columns"):
        recommendations.append("Review groupBy-heavy transformations for shuffle cost.")
    return {
        "rows": rows,
        "columns": cols,
        "estimated_partitions": partitions,
        "shuffle_operations": shuffle_ops,
        "duration_ms": round(float(duration_ms), 2),
        "recommendations": recommendations[:4],
    }


def _generate_pyspark_code(source_name, df, gold):
    cols = [str(c) for c in df.columns]
    key = _row_key_column(df) or (cols[0] if cols else "id")
    safe_cols = [re.sub(r"[^A-Za-z0-9_]+", "_", c).strip("_").lower() or "column" for c in cols]
    mapping = {c: s for c, s in zip(cols, safe_cols)}
    renamed = ",\n    ".join([f'"{c}": "{mapping[c]}"' for c in cols[:18]])
    gold_group = gold.get("group_by")
    gold_measure = gold.get("measure")
    if gold_group and gold_measure:
        g = mapping.get(gold_group, re.sub(r"[^A-Za-z0-9_]+", "_", gold_group).strip("_").lower())
        m = mapping.get(gold_measure, re.sub(r"[^A-Za-z0-9_]+", "_", gold_measure).strip("_").lower())
        gold_block = f'''\ngold_df = (\n    silver_df.groupBy("{g}")\n    .agg(\n        F.count("{m}").alias("record_count"),\n        F.avg("{m}").alias("avg_{m}"),\n        F.sum("{m}").alias("sum_{m}")\n    )\n)\n'''
    else:
        gold_block = '\ngold_df = silver_df\n'
    return f'''# DataLens generated Databricks-compatible PySpark example\nfrom pyspark.sql import functions as F\n\nsource_path = "{source_name}"\n\n# BRONZE: preserve source data with minimal transformation\nbronze_df = (\n    spark.read\n    .format("csv")\n    .option("header", True)\n    .option("inferSchema", True)\n    .load(source_path)\n)\n\n# SILVER: normalize, cast, and deduplicate\nrename_map = {{{renamed}}}\nsilver_df = bronze_df\nfor old_name, new_name in rename_map.items():\n    silver_df = silver_df.withColumnRenamed(old_name, new_name)\n\nsilver_df = (\n    silver_df\n    .dropDuplicates(["{mapping.get(key, key)}"])\n    .filter(F.col("{mapping.get(key, key)}").isNotNull())\n)\n{gold_block}\n# DELTA persistence examples\nbronze_df.write.format("delta").mode("append").saveAsTable("datalens.bronze_raw")\nsilver_df.write.format("delta").mode("overwrite").saveAsTable("datalens.silver_clean")\ngold_df.write.format("delta").mode("overwrite").saveAsTable("datalens.gold_summary")\n'''


def _generate_sql_code(df, gold):
    cols = [str(c) for c in df.columns]
    key = _row_key_column(df) or (cols[0] if cols else "id")
    normalized = [re.sub(r"[^A-Za-z0-9_]+", "_", c).strip("_").lower() or "column" for c in cols]
    source_cols = ",\n    ".join([f'`{c}` AS `{n}`' for c, n in zip(cols[:24], normalized[:24])])
    key_n = re.sub(r"[^A-Za-z0-9_]+", "_", str(key)).strip("_").lower() or "id"
    group = gold.get("group_by")
    measure = gold.get("measure")
    if group and measure:
        g = re.sub(r"[^A-Za-z0-9_]+", "_", str(group)).strip("_").lower()
        m = re.sub(r"[^A-Za-z0-9_]+", "_", str(measure)).strip("_").lower()
        gold_sql = f'''\nCREATE OR REPLACE TABLE gold_summary AS\nSELECT\n    `{g}` AS group_value,\n    COUNT(`{m}`) AS record_count,\n    AVG(`{m}`) AS avg_value,\n    SUM(`{m}`) AS sum_value\nFROM silver_clean\nGROUP BY `{g}`;\n'''
    else:
        gold_sql = '''\nCREATE OR REPLACE TABLE gold_summary AS\nSELECT COUNT(*) AS record_count FROM silver_clean;\n'''
    return f'''-- DataLens generated Databricks SQL example\n\nCREATE OR REPLACE TABLE bronze_raw AS\nSELECT * FROM read_files('/path/to/source');\n\nCREATE OR REPLACE TABLE silver_clean AS\nSELECT\n    {source_cols}\nFROM bronze_raw\nWHERE `{key_n}` IS NOT NULL\nQUALIFY ROW_NUMBER() OVER (PARTITION BY `{key_n}` ORDER BY `{key_n}`) = 1;\n{gold_sql}\n-- Delta Lake MERGE pattern\nMERGE INTO silver_clean AS target\nUSING bronze_raw AS source\nON target.`{key_n}` = source.`{key_n}`\nWHEN MATCHED THEN UPDATE SET *\nWHEN NOT MATCHED THEN INSERT *;\n'''


def _generate_delta_merge_sql(df):
    key = _row_key_column(df) or (str(df.columns[0]) if len(df.columns) else "id")
    key_n = re.sub(r"[^A-Za-z0-9_]+", "_", str(key)).strip("_").lower() or "id"
    return f'''MERGE INTO gold_target AS t\nUSING silver_updates AS s\nON t.`{key_n}` = s.`{key_n}`\nWHEN MATCHED THEN\n  UPDATE SET *\nWHEN NOT MATCHED THEN\n  INSERT *;'''


def _engineering_analysis_from_df(df, previous_df=None, source_filename="dataset", duration_ms=0):
    quality = _quality_analysis(df)
    silver_df, transformations = _silver_transform(df)
    gold = _gold_summary(silver_df)
    schema = _schema_snapshot(df)
    drift = _compare_schemas(previous_df, df) if previous_df is not None and not previous_df.empty else {"added": [], "removed": [], "changed": [], "has_drift": False}
    cdc = _cdc_compare(previous_df, df) if previous_df is not None and not previous_df.empty else {"key_column": _row_key_column(df), "inserted": 0, "updated": 0, "deleted": 0, "unchanged": 0, "changes": [], "message": "Upload a previous dataset version to compare incremental changes."}
    performance = _performance(df, duration_ms)
    lineage = _lineage(df, gold)
    pyspark = _generate_pyspark_code(source_filename, df, gold)
    sql = _generate_sql_code(df, gold)
    delta_merge = _generate_delta_merge_sql(df)
    return {
        "source": {"filename": source_filename, "rows": int(len(df)), "columns": int(len(df.columns))},
        "schema": schema,
        "schema_drift": drift,
        "quality": quality,
        "silver": {
            "rows_before": int(len(df)),
            "rows_after": int(len(silver_df)),
            "transformations": transformations,
            "preview": [_json_safe_record(r) for _, r in silver_df.head(12).iterrows()],
        },
        "gold": gold,
        "cdc": cdc,
        "scd": {
            "key_column": cdc.get("key_column"),
            "type1": "Overwrite the current value; history is not preserved.",
            "type2": "Create a new version and retain the previous value with effective dates / version metadata.",
            "sample_updates": cdc.get("changes", [])[:12],
        },
        "lineage": lineage,
        "performance": performance,
        "codes": {"pyspark": pyspark, "sql": sql, "delta_merge": delta_merge},
    }


def _save_engineering_run(result, run_type="FULL", status="SUCCESS", message="Pipeline completed."):
    conn = get_db()
    try:
        quality = result.get("quality", {}) if isinstance(result, dict) else {}
        cdc = result.get("cdc", {}) if isinstance(result, dict) else {}
        perf = result.get("performance", {}) if isinstance(result, dict) else {}
        cur = conn.execute('''
            INSERT INTO engineering_runs
            (user_id, source_filename, status, run_type, rows_processed, quality_score,
             inserted_rows, updated_rows, deleted_rows, unchanged_rows, duration_ms, message, result_json)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ''', (
            session["user_id"],
            result.get("source", {}).get("filename", "Dataset"),
            status,
            run_type,
            int(result.get("source", {}).get("rows", 0)),
            float(quality.get("quality_score", 0)),
            int(cdc.get("inserted", 0)),
            int(cdc.get("updated", 0)),
            int(cdc.get("deleted", 0)),
            int(cdc.get("unchanged", 0)),
            float(perf.get("duration_ms", 0)),
            message,
            json.dumps(result, ensure_ascii=False),
        ))
        conn.commit()
        return int(cur.lastrowid)
    finally:
        conn.close()


@app.route("/api/engineering/analyze", methods=["POST"])
@login_required
def engineering_analyze():
    started = time.perf_counter()
    payload = request.get_json(silent=True) or {}
    data = payload.get("data") or {}
    records = data.get("records") or []
    df = _engineering_records_to_df(records)
    if df.empty:
        return jsonify({"success": False, "message": "Generate a dashboard first."}), 400
    try:
        result = _engineering_analysis_from_df(
            df,
            source_filename=str(data.get("filename", "Dataset")),
            duration_ms=(time.perf_counter() - started) * 1000,
        )
        simulate_failure = bool(payload.get("simulate_failure"))
        if simulate_failure:
            result["performance"]["duration_ms"] = round((time.perf_counter() - started) * 1000, 2)
            run_id = _save_engineering_run(result, run_type="SIMULATED_FAILURE", status="FAILED", message="Simulated pipeline failure for retry demonstration.")
            return jsonify({"success": True, "status": "FAILED", "run_id": run_id, "message": "Simulated pipeline failure.", "result": result})
        result["performance"]["duration_ms"] = round((time.perf_counter() - started) * 1000, 2)
        run_id = _save_engineering_run(result, run_type="FULL", status="SUCCESS", message="Pipeline completed successfully.")
        return jsonify({"success": True, "status": "SUCCESS", "run_id": run_id, "result": result})
    except Exception as exc:
        return jsonify({"success": False, "message": str(exc)}), 400


@app.route("/api/engineering/compare", methods=["POST"])
@login_required
def engineering_compare():
    started = time.perf_counter()
    current_file = request.files.get("current_file")
    previous_file = request.files.get("previous_file")
    if current_file is None or previous_file is None:
        return jsonify({"success": False, "message": "Select both current and previous dataset files."}), 400
    try:
        current_ext = current_file.filename.lower().rsplit(".", 1)[-1]
        previous_ext = previous_file.filename.lower().rsplit(".", 1)[-1]
        if current_ext not in ["csv", "xlsx", "xls"] or previous_ext not in ["csv", "xlsx", "xls"]:
            return jsonify({"success": False, "message": "Only CSV and Excel files are supported."}), 400
        current_path = os.path.join(UPLOAD_FOLDER, f"{uuid.uuid4().hex}.{current_ext}")
        previous_path = os.path.join(UPLOAD_FOLDER, f"{uuid.uuid4().hex}.{previous_ext}")
        current_file.save(current_path)
        previous_file.save(previous_path)
        try:
            current_df = load_file(current_path)
            previous_df = load_file(previous_path)
            result = _engineering_analysis_from_df(
                current_df,
                previous_df=previous_df,
                source_filename=current_file.filename,
                duration_ms=(time.perf_counter() - started) * 1000,
            )
            result["comparison"] = {"previous_filename": previous_file.filename, "current_filename": current_file.filename}
            result["performance"]["duration_ms"] = round((time.perf_counter() - started) * 1000, 2)
            run_id = _save_engineering_run(result, run_type="INCREMENTAL_CDC", status="SUCCESS", message="Schema and CDC comparison completed.")
            return jsonify({"success": True, "status": "SUCCESS", "run_id": run_id, "result": result})
        finally:
            for path in [current_path, previous_path]:
                try:
                    if os.path.exists(path):
                        os.remove(path)
                except OSError:
                    pass
    except Exception as exc:
        return jsonify({"success": False, "message": str(exc)}), 400


@app.route("/api/engineering/runs", methods=["GET"])
@login_required
def engineering_runs():
    conn = get_db()
    try:
        rows = conn.execute('''
            SELECT id, source_filename, status, run_type, rows_processed, quality_score,
                   inserted_rows, updated_rows, deleted_rows, unchanged_rows, duration_ms,
                   message, created_at
            FROM engineering_runs
            WHERE user_id = ?
            ORDER BY id DESC
            LIMIT 20
        ''', (session["user_id"],)).fetchall()
        return jsonify({"success": True, "runs": [dict(r) for r in rows]})
    finally:
        conn.close()


@app.route("/api/engineering/runs/<int:run_id>/retry", methods=["POST"])
@login_required
def engineering_retry(run_id):
    conn = get_db()
    try:
        row = conn.execute("SELECT * FROM engineering_runs WHERE id = ? AND user_id = ?", (run_id, session["user_id"])).fetchone()
        if row is None:
            return jsonify({"success": False, "message": "Run not found."}), 404
        result = json.loads(row["result_json"] or "{}")
        result["performance"]["duration_ms"] = round(max(1.0, float(result.get("performance", {}).get("duration_ms", 0)) * 0.85), 2)
        new_id = _save_engineering_run(result, run_type="RETRY", status="SUCCESS", message=f"Retry of run #{run_id} completed successfully.")
        return jsonify({"success": True, "run_id": new_id, "result": result, "message": "Pipeline retry succeeded."})
    finally:
        conn.close()


@app.route("/api/engineering/schedule", methods=["POST"])
@login_required
def engineering_schedule():
    payload = request.get_json(silent=True) or {}
    frequency = str(payload.get("frequency", "Manual"))
    return jsonify({
        "success": True,
        "schedule": {
            "frequency": frequency,
            "status": "SIMULATED",
            "message": f"DataLens scheduled the pipeline as a local simulation: {frequency}."
        }
    })


# =========================================================
# ERROR HANDLERS
# =========================================================

@app.errorhandler(413)
def file_too_large(error):
    return jsonify({
        "success": False,
        "error": "File is too large. Maximum size is 50 MB."
    }), 413


@app.errorhandler(500)
def server_error(error):
    return jsonify({
        "success": False,
        "error": "An internal server error occurred."
    }), 500


# =========================================================
# START DATALENS
# =========================================================

if __name__ == "__main__":
    app.run(
        debug=True,
        host="127.0.0.1",
        port=5000
    )
