import os
import traceback
from datetime import datetime
from flask import Flask, render_template, request, jsonify, redirect, session
import pandas as pd
from flask_cors import CORS

# ============================
# CONFIG
# ============================
app = Flask(__name__, static_folder="static")
app.secret_key = "inwmh_secret_123"
CORS(app)

BASE_DIR = os.path.abspath(os.path.dirname(__file__))
EXCEL_FILE = os.path.join(BASE_DIR, "bookings.xlsx")

# ============================
# EXCEL CONFIG
# ============================
COLUMNS = [
    "Timestamp", "Name", "Email", "Phone", "Company",
    "Setup", "People", "Package",
    "Date", "Time Slot", "Duration",
    "Requirements", "Referral"
]

BOOKINGS_SHEET = "Bookings"
OFFDAYS_SHEET = "OffDays"

# ============================
# INITIALIZE EXCEL
# ============================
def initialize_excel():
    if not os.path.exists(EXCEL_FILE):
        bookings_df = pd.DataFrame(columns=COLUMNS)
        offdays_df = pd.DataFrame(columns=["Date"])

        with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl") as writer:
            bookings_df.to_excel(writer, sheet_name=BOOKINGS_SHEET, index=False)
            offdays_df.to_excel(writer, sheet_name=OFFDAYS_SHEET, index=False)

def load_excel_sheets():
    initialize_excel()
    try:
        xls = pd.ExcelFile(EXCEL_FILE)

        if BOOKINGS_SHEET in xls.sheet_names:
            bookings_df = pd.read_excel(xls, BOOKINGS_SHEET)
        else:
            bookings_df = pd.read_excel(xls, xls.sheet_names[0])

        for col in COLUMNS:
            if col not in bookings_df.columns:
                bookings_df[col] = ""

        if OFFDAYS_SHEET in xls.sheet_names:
            offdays_df = pd.read_excel(xls, OFFDAYS_SHEET)
        else:
            offdays_df = pd.DataFrame(columns=["Date"])

        return bookings_df, offdays_df

    except Exception:
        traceback.print_exc()
        return pd.DataFrame(columns=COLUMNS), pd.DataFrame(columns=["Date"])

def save_excel_sheets(bookings_df, offdays_df):
    try:
        with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl") as writer:
            bookings_df.to_excel(writer, BOOKINGS_SHEET, index=False)
            offdays_df.to_excel(writer, OFFDAYS_SHEET, index=False)
        return True
    except Exception:
        traceback.print_exc()
        return False

initialize_excel()

# ============================
# PUBLIC ROUTES
# ============================
@app.route("/")
def home():
    return render_template("index.html")

@app.route("/booked_slots")
def booked_slots():
    date_str = request.args.get("date")
    if not date_str:
        return jsonify({"booked": [], "off_day": False})

    try:
        bookings_df, offdays_df = load_excel_sheets()

        booked = bookings_df[bookings_df["Date"].astype(str) == date_str]["Time Slot"].dropna().astype(str).tolist()
        off_day = date_str in offdays_df["Date"].astype(str).tolist()

        return jsonify({"booked": booked, "off_day": off_day})
    except Exception:
        traceback.print_exc()
        return jsonify({"booked": [], "off_day": False}), 500

@app.route("/submit", methods=["POST"])
def submit_form():
    try:
        data = request.get_json()
        if not data:
            return jsonify({"message": "No data received"}), 400

        bookings_df, offdays_df = load_excel_sheets()

        new_entry = {
            "Timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "Name": data.get("name", ""),
            "Email": data.get("email", ""),
            "Phone": data.get("phone", ""),
            "Company": data.get("company", ""),
            "Setup": data.get("setup", ""),
            "People": data.get("people", ""),
            "Package": data.get("package", ""),
            "Date": data.get("date", ""),
            "Time Slot": data.get("time_slot", ""),
            "Duration": data.get("duration", ""),
            "Requirements": data.get("requirements", ""),
            "Referral": data.get("referral", "")
        }

        bookings_df = pd.concat([bookings_df, pd.DataFrame([new_entry])], ignore_index=True)

        if save_excel_sheets(bookings_df, offdays_df):
            return jsonify({"message": "Booking saved successfully"}), 200

        return jsonify({"message": "Failed to save booking"}), 500

    except Exception:
        traceback.print_exc()
        return jsonify({"message": "Server error"}), 500

# ============================
# ADMIN AUTH
# ============================
ADMIN_USERNAME = "admin"
ADMIN_PASSWORD = "1234"

@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        if (
            request.form.get("username") == ADMIN_USERNAME and
            request.form.get("password") == ADMIN_PASSWORD
        ):
            session["admin"] = True
            return redirect("/admin")

        return render_template("login.html", error="Invalid credentials")

    return render_template("login.html")

@app.route("/logout")
def logout():
    session.pop("admin", None)
    return redirect("/login")

# ============================
# ADMIN ROUTES
# ============================
@app.route("/admin")
def admin_dashboard():
    if "admin" not in session:
        return redirect("/login")
    return render_template("admin.html")

@app.route("/api/bookings")
def get_all_bookings():
    try:
        bookings_df, _ = load_excel_sheets()

        # 🔥 FIX: Replace NaN with empty string (JSON-safe)
        bookings_df = bookings_df.where(pd.notna(bookings_df), "")

        return jsonify({
            "status": "success",
            "bookings": bookings_df.to_dict(orient="records")
        })
    except Exception as e:
        traceback.print_exc()
        return jsonify({
            "status": "error",
            "message": str(e)
        }), 500

# ============================
# START SERVER
# ============================
if __name__ == "__main__":
    print("🚀 Running on http://127.0.0.1:5000")
    app.run(debug=True)
