from flask import Flask, render_template, request, jsonify, session, redirect, Response
from werkzeug.security import generate_password_hash, check_password_hash
import pandas as pd
import numpy as np
import os
import uuid
import re
import sqlite3
import json
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
body{{margin:0;font-family:Segoe UI,Arial,sans-serif;background:linear-gradient(135deg,#f8faff,#eef5ff);color:#0f172a}}
.wrap{{max-width:1450px;margin:auto;padding:24px}}
.top{{display:flex;justify-content:space-between;align-items:center;gap:18px;margin-bottom:20px}}
.brand{{font-size:26px;font-weight:900}}
.badge{{padding:9px 14px;border-radius:999px;background:#ecfdf5;color:#047857;border:1px solid #86efac;font-weight:800;font-size:13px}}
.shared-note{{margin-top:7px;color:#64748b;font-size:13px}}
h1{{margin:4px 0 0;font-size:30px}}
.meta{{color:#64748b;margin-top:6px}}
.kpis{{display:grid;grid-template-columns:repeat(auto-fit,minmax(180px,1fr));gap:15px;margin:20px 0}}
.kpi{{background:#fff;border:1px solid #dbeafe;border-radius:18px;padding:18px;box-shadow:0 10px 28px rgba(30,64,175,.08)}}
.kpi .l{{font-size:13px;color:#64748b;font-weight:800}}
.kpi .v{{font-size:28px;font-weight:900;margin-top:8px}}
.charts{{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:18px}}
.layout-compact .charts{{grid-template-columns:repeat(3,minmax(0,1fr));gap:12px}}
.layout-wide .charts,.layout-focus .charts{{grid-template-columns:1fr}}
.layout-compact .chart{{height:245px}}
.layout-wide .chart{{height:360px}}
.layout-focus .chart{{height:420px}}
.card{{background:#fff;border:1px solid #dbeafe;border-radius:20px;padding:16px;box-shadow:0 10px 28px rgba(30,64,175,.07)}}
.card h3{{margin:0 0 8px;font-size:17px}}
.chart{{height:330px}}
.insights{{margin-top:18px;background:linear-gradient(135deg,#eef6ff,#f5f3ff);border:1px solid #dbeafe;border-radius:20px;padding:16px}}
.insights h2{{margin:0 0 10px;font-size:17px}}
.insight{{padding:9px 11px;margin:7px 0;background:#fff;border:1px solid #dbeafe;border-radius:12px;color:#334155;font-size:12px;line-height:1.5}}
.preview{{margin-top:18px;background:#fff;border:1px solid #dbeafe;border-radius:20px;padding:16px;overflow:auto}}
table{{width:100%;border-collapse:collapse}}
th,td{{padding:10px 12px;border-bottom:1px solid #e2e8f0;text-align:left;white-space:nowrap}}
th{{background:#eff6ff;color:#1d4ed8}}
@media(max-width:800px){{.charts{{grid-template-columns:1fr}}.wrap{{padding:14px}}}}
</style>
</head>
<body>
<div class="wrap layout-${{data.get('layout_mode', 'standard')}}">
  <div class="top"><div><div class="brand">DataLens</div>{readonly_note}</div><div class="badge">Analytics Dashboard</div></div>
  <h1 id="title">{title_safe}</h1>
  <div class="meta" id="meta"></div>
  <div class="insights"><h2>AI Insights</h2><div id="insights"></div></div>
  <div class="kpis" id="kpis"></div>
  <div class="charts" id="charts"></div>
  <div class="preview"><h3>Data Preview</h3><div id="preview"></div></div>
</div>
<script>
const DATA = {payload};
function esc(v){{return String(v??'').replace(/[&<>\"]/g,m=>({{'&':'&amp;','<':'&lt;','>':'&gt;','\"':'&quot;'}}[m]))}}
document.getElementById('meta').textContent = `${{DATA.analysis?.rows ?? 0}} rows • ${{DATA.analysis?.columns ?? (DATA.columns||[]).length}} columns • ${{DATA.filename||''}}`;
const ins=document.getElementById('insights');(DATA.insights||[]).forEach(x=>{{const d=document.createElement('div');d.className='insight';d.textContent=x.text||String(x);ins.appendChild(d)}});
const k=document.getElementById('kpis');
(DATA.kpis||[]).forEach(x=>{{const d=document.createElement('div');d.className='kpi';d.innerHTML=`<div class="l">${{esc(x.title)}}</div><div class="v">${{esc(x.value)}}</div>`;k.appendChild(d)}});
function draw(el,c){{let trace=[],layout={{margin:{{l:45,r:15,t:10,b:50}},paper_bgcolor:'transparent',plot_bgcolor:'transparent',font:{{family:'Segoe UI,Arial',size:11,color:'#64748b'}},showlegend:false,responsive:true}};if(c.type==='bar')trace=[{{x:c.labels||[],y:c.values||[],type:'bar'}}];else if(c.type==='line')trace=[{{x:c.x||[],y:c.y||[],type:'scatter',mode:'lines+markers',line:{{width:3,shape:'spline'}}}}];else if(c.type==='histogram')trace=[{{x:c.values||[],type:'histogram'}}];else if(c.type==='scatter')trace=[{{x:c.x||[],y:c.y||[],type:'scatter',mode:'markers'}}];Plotly.newPlot(el,trace,layout,{{displayModeBar:false,responsive:true}})}}
const cg=document.getElementById('charts');(DATA.charts||[]).forEach((c,i)=>{{const card=document.createElement('div');card.className='card';card.innerHTML=`<h3>${{esc(c.title||'Data Insight')}}</h3><div class="chart" id="chart${{i}}"></div>`;cg.appendChild(card);draw('chart'+i,c)}});
const pv=document.getElementById('preview'),rows=DATA.preview||[];if(rows.length){{const cols=Object.keys(rows[0]);const t=document.createElement('table');t.innerHTML='<thead><tr>'+cols.map(c=>`<th>${{esc(c)}}</th>`).join('')+'</tr></thead><tbody>'+rows.map(r=>'<tr>'+cols.map(c=>`<td>${{esc(r[c])}}</td>`).join('')+'</tr>').join('')+'</tbody>';pv.appendChild(t)}}else pv.textContent='No preview available.';
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
