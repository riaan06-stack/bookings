from flask import Flask, render_template, request, jsonify
import pandas as pd
import os
from datetime import datetime
import traceback

app = Flask(__name__)
app.secret_key = "inwmh_secret_123"
EXCEL_FILE = "bookings.xlsx"

def initialize_excel():
    if not os.path.exists(EXCEL_FILE):
        df = pd.DataFrame(columns=[
            "Timestamp", "Name", "Email", "Phone", "Company",
            "Setup", "People", "Package", "Addons",
            "Date", "Duration", "Time Slot", "Frequency",
            "Requirements", "Referral"
        ])
        df.to_excel(EXCEL_FILE, index=False)

initialize_excel()

@app.route("/")
def home():
    return render_template("index.html")

@app.route("/submit", methods=["POST"])
def submit_form():
    try:
        data = request.get_json()
        print("\n🔹 Received data:", data, "\n")

        if not data:
            return jsonify({"message": "No data received"}), 400

        if os.path.exists(EXCEL_FILE):
            try:
                df = pd.read_excel(EXCEL_FILE)
            except Exception:
                df = pd.DataFrame(columns=[
                    "Timestamp", "Name", "Email", "Phone", "Company",
                    "Setup", "People", "Package", "Addons",
                    "Date", "Duration", "Time Slot", "Frequency",
                    "Requirements", "Referral"
                ])
        else:
            df = pd.DataFrame(columns=[
                "Timestamp", "Name", "Email", "Phone", "Company",
                "Setup", "People", "Package", "Addons",
                "Date", "Duration", "Time Slot", "Frequency",
                "Requirements", "Referral"
            ])

        new_entry = {
            "Timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "Name": data.get("name"),
            "Email": data.get("email"),
            "Phone": data.get("phone"),
            "Company": data.get("company"),
            "Setup": data.get("setup"),
            "People": data.get("people"),
            "Package": data.get("package"),
            "Addons": data.get("selected_addons"),
            "Date": data.get("date"),
            "Duration": data.get("duration"),
            "Time Slot": data.get("time_slot"),
            "Frequency": data.get("frequency"),
            "Requirements": data.get("requirements"),
            "Referral": data.get("referral")
        }

        df = pd.concat([df, pd.DataFrame([new_entry])], ignore_index=True)
        df.to_excel(EXCEL_FILE, index=False)

        return jsonify({"message": "Booking saved successfully!"}), 200

    except Exception as e:
        traceback.print_exc()
        return jsonify({"message": f"Error: {str(e)}"}), 500


if __name__ == "__main__":
    print("✅ Flask is starting... Visit http://127.0.0.1:5000")
    app.run(debug=True)
