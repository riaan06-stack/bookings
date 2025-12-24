import os
import traceback
from datetime import datetime
from flask import Flask, render_template, request, jsonify, redirect, session
from flask_mail import Mail, Message
import pandas as pd
from flask_cors import CORS
from threading import Lock
booking_lock = Lock()

# ============================
# CONFIG
# ============================
app = Flask(__name__, static_folder="static")
app.secret_key = "inwmh_secret_123"
CORS(app)

# Email Configuration
app.config['MAIL_SERVER'] = 'smtp.gmail.com'
app.config['MAIL_PORT'] = 587
app.config['MAIL_USE_TLS'] = True
app.config['MAIL_USERNAME'] = 'sanjanastudys@gmail.com'  # Replace with your email
app.config['MAIL_PASSWORD'] = 'ynyf lltf zdwu rreu'     # Replace with app password
app.config['MAIL_DEFAULT_SENDER'] = 'sanjanastudys@gmail.com'

mail = Mail(app)

BASE_DIR = os.path.abspath(os.path.dirname(__file__))
EXCEL_FILE = os.path.join(BASE_DIR, "bookings.xlsx")

# ============================
# EXCEL CONFIG
# ============================
COLUMNS = [
    "Timestamp", "Name", "Email", "Phone", "Company",
    "Setup", "People", "Package",
    "Date", "Time Slot", "Duration",
    "Base Price", "Saturday Surcharge", "Total Price",
    "Requirements", "Referral"
]

BOOKINGS_SHEET = "Bookings"
OFFDAYS_SHEET = "OffDays"

# Time slots configuration (must match frontend)
TIMESLOTS = [
    "10:00 AM", "11:00 AM",
    "12:00 PM", "1:00 PM", "2:00 PM",
    "3:00 PM", "4:00 PM", "5:00 PM",
    "6:00 PM"
]

# ============================
# EMAIL FUNCTIONS
# ============================
def send_booking_confirmation(booking_data):
    """Send booking confirmation emails to user and admin"""
    
    # Format date for display
    booking_date = datetime.strptime(booking_data['Date'], "%Y-%m-%d")
    formatted_date = booking_date.strftime("%B %d, %Y")
    
    # User email
    user_subject = f"INWMH Studios Booking Confirmation - {formatted_date}"
    user_body = f"""
Dear {booking_data['Name']},

Thank you for booking with INWMH Studios!

📅 **Booking Details:**
- Date: {formatted_date}
- Time: {booking_data['Time Slot']}
- Duration: {booking_data['Duration']} hours
- Studio Setup: {booking_data['Setup']}
- Package: {booking_data['Package']}
- Number of People: {booking_data.get('People', '1')}
- Total Amount: ${booking_data.get('Total Price', '0.00')}

📍 **Studio Address:**
INWMH Studios
[Your Studio Address Here]

📞 **Contact Information:**
Phone: [Your Phone Number]
Email: studio@inwmh.com

💡 **Important Notes:**
- Please arrive 15 minutes before your scheduled time
- Bring any necessary equipment or materials
- Contact us if you need to reschedule or cancel

We're excited to have you at our studio! Our team will contact you shortly if any additional information is needed.

Best regards,
INWMH Studios Team
www.inwmhstudios.com
"""
    
    # Admin email
    admin_subject = f"📋 New Booking: {booking_data['Name']} - {formatted_date}"
    admin_body = f"""
🚨 **NEW BOOKING RECEIVED** 🚨

**Customer Information:**
- Name: {booking_data['Name']}
- Email: {booking_data['Email']}
- Phone: {booking_data['Phone']}
- Company: {booking_data.get('Company', 'Not provided')}

**Booking Details:**
- Date: {formatted_date}
- Time: {booking_data['Time Slot']}
- Duration: {booking_data['Duration']} hours
- Setup: {booking_data['Setup']}
- Package: {booking_data['Package']}
- People: {booking_data.get('People', '1')}

**Financial Information:**
- Base Price: ${booking_data.get('Base Price', '0.00')}
- Saturday Surcharge: ${booking_data.get('Saturday Surcharge', '0.00')}
- **Total: ${booking_data.get('Total Price', '0.00')}**

**Additional Information:**
- Special Requirements: {booking_data.get('Requirements', 'None')}
- Referral Source: {booking_data.get('Referral', 'Not specified')}
- Booking Time: {booking_data['Timestamp']}

**Action Required:**
1. Confirm booking with customer
2. Prepare studio setup
3. Update schedule
"""
    
    try:
        # Send to user
        user_msg = Message(
            subject=user_subject,
            recipients=[booking_data['Email']],
            body=user_body
        )
        mail.send(user_msg)
        
        # Send to admin (replace with actual admin email)
        admin_msg = Message(
            subject=admin_subject,
            recipients=['sanjanastudys@gmail.com'],  # Replace with your admin email
            body=admin_body
        )
        mail.send(admin_msg)
        
        app.logger.info(f"Emails sent successfully for booking by {booking_data['Email']}")
        return True
        
    except Exception as e:
        app.logger.error(f"Failed to send emails: {str(e)}")
        return False

# ============================
# HELPER FUNCTIONS
# ============================
def get_time_slots_for_duration(start_time, duration_hours):
    """Get all time slots needed for a given duration"""
    start_index = TIMESLOTS.index(start_time) if start_time in TIMESLOTS else -1
    if start_index == -1:
        return []
    
    duration_slots = duration_hours  # Each hour = 1 slot
    required_slots = []
    
    for i in range(duration_slots):
        if start_index + i < len(TIMESLOTS):
            required_slots.append(TIMESLOTS[start_index + i])
    
    return required_slots

def check_time_slot_overlap(existing_bookings, new_start_time, new_duration):
    """Check if new booking overlaps with existing bookings"""
    new_slots = get_time_slots_for_duration(new_start_time, new_duration)
    if not new_slots:
        return True, []  # Invalid time slot
    
    for booking in existing_bookings:
        existing_start = booking["Time Slot"]
        existing_duration = int(booking["Duration"]) if str(booking["Duration"]).isdigit() else 2
        existing_slots = get_time_slots_for_duration(existing_start, existing_duration)
        
        # Check for any overlap
        if any(slot in existing_slots for slot in new_slots):
            return True, existing_slots
    
    return False, []

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
        
        # Filter bookings for the selected date
        date_bookings = bookings_df[bookings_df["Date"].astype(str) == date_str]
        
        # Create list of booked slots with duration info
        booked_slots_with_duration = []
        for _, row in date_bookings.iterrows():
            start_time = str(row["Time Slot"]) if pd.notna(row["Time Slot"]) else ""
            duration = int(row["Duration"]) if str(row["Duration"]).isdigit() else 2
            
            if start_time and start_time in TIMESLOTS:
                booked_slots_with_duration.append({
                    "startTime": start_time,
                    "duration": duration
                })
        
        # Check if date is an off day
        off_day = date_str in offdays_df["Date"].astype(str).tolist()
        
        return jsonify({
            "booked": booked_slots_with_duration, 
            "off_day": off_day
        })
        
    except Exception:
        traceback.print_exc()
        return jsonify({"booked": [], "off_day": False}), 500

@app.route("/submit", methods=["POST"])
def submit_form():
    with booking_lock:  # 🔒 ensures ONE booking at a time
        try:
            data = request.get_json()
            if not data:
                return jsonify({"message": "No data received"}), 400

            # -------------------------
            # Validate required fields
            # -------------------------
            required_fields = ["name", "email", "phone", "setup", "package", "date", "time_slot"]
            for field in required_fields:
                if field not in data or not data[field]:
                    return jsonify({"message": f"Missing required field: {field}"}), 400

            # -------------------------
            # Load Excel data
            # -------------------------
            bookings_df, offdays_df = load_excel_sheets()

            # -------------------------
            # Check Sunday (closed)
            # -------------------------
            date_obj = datetime.strptime(data["date"], "%Y-%m-%d")
            if date_obj.weekday() == 6:  # Sunday
                return jsonify({"message": "Studio is closed on Sundays"}), 400

            # -------------------------
            # Validate time slot
            # -------------------------
            if data["time_slot"] not in TIMESLOTS:
                return jsonify({"message": "Invalid time slot"}), 400

            # -------------------------
            # Duration
            # -------------------------
            duration = int(data.get("duration", 2))

            # -------------------------
            # Overlap check
            # -------------------------
            date_str = data["date"]
            existing_bookings = bookings_df[
                bookings_df["Date"].astype(str) == date_str
            ]
            existing_bookings_list = existing_bookings.to_dict(orient="records")

            has_overlap, _ = check_time_slot_overlap(
                existing_bookings_list,
                data["time_slot"],
                duration
            )

            if has_overlap:
                return jsonify({
                    "message": "Time slot overlaps with existing booking",
                    "overlap": True
                }), 409

            # -------------------------
            # Prices
            # -------------------------
            base_price = float(data.get("base_price", 0))
            saturday_surcharge = float(data.get("saturday_surcharge", 0))
            total_price = float(data.get("total_price", 0))

            # -------------------------
            # New booking entry
            # -------------------------
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
                "Duration": str(duration),
                "Base Price": str(base_price),
                "Saturday Surcharge": str(saturday_surcharge),
                "Total Price": str(total_price),
                "Requirements": data.get("requirements", ""),
                "Referral": data.get("referral", "")
            }

            # -------------------------
            # Save booking
            # -------------------------
            bookings_df = pd.concat(
                [bookings_df, pd.DataFrame([new_entry])],
                ignore_index=True
            )

            if not save_excel_sheets(bookings_df, offdays_df):
                return jsonify({"message": "Failed to save booking"}), 500

            # -------------------------
            # Send email (non-blocking logic)
            # -------------------------
            email_sent = send_booking_confirmation(new_entry)

            return jsonify({
                "message": "Booking saved successfully",
                "booking": new_entry,
                "email_sent": email_sent
            }), 200

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

        # Replace NaN with empty string (JSON-safe)
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

@app.route("/api/bookings/<booking_id>", methods=["DELETE"])
def delete_booking(booking_id):
    if "admin" not in session:
        return jsonify({"message": "Unauthorized"}), 401
    
    try:
        bookings_df, offdays_df = load_excel_sheets()
        
        # Check if booking exists (using index-based deletion)
        if len(bookings_df) == 0:
            return jsonify({"message": "No bookings found"}), 404
        
        try:
            idx = int(booking_id)
            if idx < 0 or idx >= len(bookings_df):
                return jsonify({"message": "Booking not found"}), 404
            
            # Remove the booking
            bookings_df = bookings_df.drop(idx).reset_index(drop=True)
            
            if save_excel_sheets(bookings_df, offdays_df):
                return jsonify({"message": "Booking deleted successfully"}), 200
            else:
                return jsonify({"message": "Failed to save changes"}), 500
                
        except ValueError:
            return jsonify({"message": "Invalid booking ID"}), 400
            
    except Exception as e:
        traceback.print_exc()
        return jsonify({"message": str(e)}), 500

# ============================
# START SERVER
# ============================
if __name__ == "__main__":
    print("🚀 Server running on http://127.0.0.1:5000")
    print("📅 Booking system with email notifications")
    print("📧 Email configuration active")
    app.run(debug=True)
