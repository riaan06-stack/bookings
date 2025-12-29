import os
import traceback
import uuid
from datetime import datetime
from threading import Lock
from flask import Flask, render_template, request, jsonify, redirect, session
from flask_mail import Mail, Message
import pandas as pd
from flask_cors import CORS

# ============================
# CONFIG
# ============================
app = Flask(__name__, static_folder="static", template_folder="templates")
app.secret_key = "inwmh_secret_123"
CORS(app)

# Email Configuration
app.config['MAIL_SERVER'] = 'smtp.gmail.com'
app.config['MAIL_PORT'] = 587
app.config['MAIL_USE_TLS'] = True
app.config['MAIL_USERNAME'] = 'sanjanastudys@gmail.com'
app.config['MAIL_PASSWORD'] = 'ynyf lltf zdwu rreu'
app.config['MAIL_DEFAULT_SENDER'] = 'sanjanastudys@gmail.com'

mail = Mail(app)

BASE_DIR = os.path.abspath(os.path.dirname(__file__))
EXCEL_FILE = os.path.join(BASE_DIR, "bookings.xlsx")

# ============================
# EXCEL CONFIG
# ============================
COLUMNS = [
    "Booking ID",
    "Timestamp", "Name", "Email", "Phone", "Company",
    "Setup", "People", "Package",
    "Date", "Time Slot", "Duration",
    "Base Price", "Saturday Surcharge", "Total Price",
    "Payment Status",
    "Payment Marked At",
    "Requirements", "Referral"
]

BOOKINGS_SHEET = "Bookings"
OFFDAYS_SHEET = "OffDays"

# Time slots configuration
TIMESLOTS = [
    "10:00 AM", "11:00 AM",
    "12:00 PM", "1:00 PM", "2:00 PM",
    "3:00 PM", "4:00 PM", "5:00 PM",
    "6:00 PM"
]

# Admin credentials
ADMIN_USERNAME = "admin"
ADMIN_PASSWORD = "1234"

# Thread lock for booking operations
booking_lock = Lock()

# ============================
# HELPER FUNCTIONS
# ============================
def initialize_excel():
    """Initialize Excel file with required sheets if not exists"""
    if not os.path.exists(EXCEL_FILE):
        try:
            bookings_df = pd.DataFrame(columns=COLUMNS)
            offdays_df = pd.DataFrame(columns=["Date"])

            with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl") as writer:
                bookings_df.to_excel(writer, sheet_name=BOOKINGS_SHEET, index=False)
                offdays_df.to_excel(writer, sheet_name=OFFDAYS_SHEET, index=False)
            
            app.logger.info(f"Created new Excel file: {EXCEL_FILE}")
        except Exception as e:
            app.logger.error(f"Failed to create Excel file: {str(e)}")
            raise

def load_excel_sheets():
    """Load bookings and offdays sheets from Excel"""
    initialize_excel()
    
    try:
        xls = pd.ExcelFile(EXCEL_FILE)
        
        # Load bookings sheet
        if BOOKINGS_SHEET in xls.sheet_names:
            bookings_df = pd.read_excel(xls, BOOKINGS_SHEET, dtype=str)
        else:
            # Fallback to first sheet
            bookings_df = pd.read_excel(xls, xls.sheet_names[0], dtype=str)
        
        # Ensure all required columns exist
        for col in COLUMNS:
            if col not in bookings_df.columns:
                bookings_df[col] = ""
        
        # Load offdays sheet
        if OFFDAYS_SHEET in xls.sheet_names:
            offdays_df = pd.read_excel(xls, OFFDAYS_SHEET, dtype=str)
        else:
            offdays_df = pd.DataFrame(columns=["Date"])
        
        # Convert date columns to string for consistency
        if "Date" in bookings_df.columns:
            bookings_df["Date"] = bookings_df["Date"].astype(str)
        if "Date" in offdays_df.columns:
            offdays_df["Date"] = offdays_df["Date"].astype(str)
        
        return bookings_df, offdays_df
        
    except Exception as e:
        app.logger.error(f"Failed to load Excel sheets: {str(e)}")
        traceback.print_exc()
        return pd.DataFrame(columns=COLUMNS), pd.DataFrame(columns=["Date"])

def save_excel_sheets(bookings_df, offdays_df):
    """Save bookings and offdays sheets to Excel"""
    try:
        with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl") as writer:
            bookings_df.to_excel(writer, sheet_name=BOOKINGS_SHEET, index=False)
            offdays_df.to_excel(writer, sheet_name=OFFDAYS_SHEET, index=False)
        return True
    except Exception as e:
        app.logger.error(f"Failed to save Excel sheets: {str(e)}")
        traceback.print_exc()
        return False

def get_time_slots_for_duration(start_time, duration_hours):
    """Get all time slots needed for a given duration"""
    if start_time not in TIMESLOTS:
        return []
    
    start_index = TIMESLOTS.index(start_time)
    duration_slots = int(duration_hours)
    
    if start_index + duration_slots > len(TIMESLOTS):
        return []
    
    return TIMESLOTS[start_index:start_index + duration_slots]

def check_time_slot_overlap(existing_bookings, new_start_time, new_duration):
    """Check if new booking overlaps with existing bookings"""
    new_slots = get_time_slots_for_duration(new_start_time, new_duration)
    if not new_slots:
        return True, []  # Invalid time slot
    
    for booking in existing_bookings:
        existing_start = booking.get("Time Slot", "")
        existing_duration_str = booking.get("Duration", "2")
        
        try:
            existing_duration = int(existing_duration_str)
        except ValueError:
            existing_duration = 2
        
        existing_slots = get_time_slots_for_duration(existing_start, existing_duration)
        
        # Check for overlap
        if set(new_slots) & set(existing_slots):
            return True, existing_slots
    
    return False, []

def generate_booking_id():
    """Generate unique booking ID"""
    return f"INWMH-{datetime.now().strftime('%Y%m%d')}-{uuid.uuid4().hex[:6].upper()}"

def format_date_for_display(date_str):
    """Format date string for display in emails"""
    try:
        date_obj = datetime.strptime(date_str, "%Y-%m-%d")
        return date_obj.strftime("%B %d, %Y")
    except ValueError:
        return date_str

# ============================
# EMAIL FUNCTIONS
# ============================
def send_booking_confirmation(booking_data):
    """Send booking confirmation emails to user and admin"""
    
    formatted_date = format_date_for_display(booking_data.get('Date', ''))
    
    # User email
    user_subject = f"INWMH Studios Booking Confirmation - {formatted_date}"
    user_body = f"""
Dear {booking_data.get('Name', 'Customer')},

Thank you for booking with INWMH Studios!

📅 **Booking Details:**
- Date: {formatted_date}
- Time: {booking_data.get('Time Slot', '')}
- Duration: {booking_data.get('Duration', '2')} hours
- Studio Setup: {booking_data.get('Setup', '')}
- Package: {booking_data.get('Package', '')}
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
    admin_subject = f"📋 New Booking: {booking_data.get('Name', 'Customer')} - {formatted_date}"
    admin_body = f"""
🚨 **NEW BOOKING RECEIVED** 🚨

**Customer Information:**
- Name: {booking_data.get('Name', '')}
- Email: {booking_data.get('Email', '')}
- Phone: {booking_data.get('Phone', '')}
- Company: {booking_data.get('Company', 'Not provided')}

**Booking Details:**
- Date: {formatted_date}
- Time: {booking_data.get('Time Slot', '')}
- Duration: {booking_data.get('Duration', '2')} hours
- Setup: {booking_data.get('Setup', '')}
- Package: {booking_data.get('Package', '')}
- People: {booking_data.get('People', '1')}

**Financial Information:**
- Base Price: ${booking_data.get('Base Price', '0.00')}
- Saturday Surcharge: ${booking_data.get('Saturday Surcharge', '0.00')}
- **Total: ${booking_data.get('Total Price', '0.00')}**

**Additional Information:**
- Special Requirements: {booking_data.get('Requirements', 'None')}
- Referral Source: {booking_data.get('Referral', 'Not specified')}
- Booking Time: {booking_data.get('Timestamp', '')}
- Booking ID: {booking_data.get('Booking ID', '')}

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
        
        # Send to admin
        admin_msg = Message(
            subject=admin_subject,
            recipients=['sanjanastudys@gmail.com'],
            body=admin_body
        )
        mail.send(admin_msg)
        
        app.logger.info(f"Booking confirmation emails sent for {booking_data.get('Email')}")
        return True
        
    except Exception as e:
        app.logger.error(f"Failed to send booking emails: {str(e)}")
        return False

def send_payment_emails(booking_data):
    """Send payment confirmation emails to user and admin"""
    
    booking_id = booking_data.get("Booking ID", "N/A")
    formatted_date = format_date_for_display(booking_data.get('Date', ''))
    amount = booking_data.get('Total Price', '0.00')
    
    # User email
    user_subject = f"Payment Submitted — Verification Pending | Booking ID {booking_id}"
    user_body = f"""
Dear {booking_data.get('Name', 'Customer')},

Thank you for completing your payment for INWMH Studios.

🧾 **Booking ID:** {booking_id}
📅 **Date:** {formatted_date}
⏰ **Time:** {booking_data.get('Time Slot', '')}
⏳ **Duration:** {booking_data.get('Duration', '2')} hours
💰 **Amount Paid:** ${amount}

⚠️ **Important:**
Your payment is currently **under manual verification**.

Please contact our admin team and share:
• Booking ID
• Payment screenshot or UPI reference ID

📞 **Admin Contact**
Email: sanjanastudys@gmail.com
Phone / WhatsApp: +91-XXXXXXXXXX

Once verified, your booking will be confirmed.

Thank you for your patience.

— INWMH Studios
"""
    
    # Admin email
    admin_subject = f"💰 Payment Submitted — Verification Required | {booking_id}"
    admin_body = f"""
🚨 PAYMENT SUBMITTED (MANUAL VERIFICATION REQUIRED)

🧾 Booking ID: {booking_id}

👤 Customer Details:
Name: {booking_data.get('Name', '')}
Email: {booking_data.get('Email', '')}
Phone: {booking_data.get('Phone', '')}

📅 Booking Details:
Date: {formatted_date}
Time: {booking_data.get('Time Slot', '')}
Duration: {booking_data.get('Duration', '2')} hours
Package: {booking_data.get('Package', '')}

💰 Amount Reported: ${amount}

⚠️ ACTION REQUIRED:
• Verify payment manually
• Confirm booking in admin panel
• Update status to CONFIRMED or CANCELLED
"""
    
    try:
        # Send to user
        user_msg = Message(
            subject=user_subject,
            recipients=[booking_data.get("Email")],
            body=user_body
        )
        mail.send(user_msg)
        
        # Send to admin
        admin_msg = Message(
            subject=admin_subject,
            recipients=["sanjanastudys@gmail.com"],
            body=admin_body
        )
        mail.send(admin_msg)
        
        app.logger.info(f"Payment emails sent for booking {booking_id}")
        return True
        
    except Exception as e:
        app.logger.error(f"Failed to send payment emails: {str(e)}")
        return False

def send_admin_request_email(booking_data):
    """Send booking request notification to admin when payment is submitted"""
    
    formatted_date = format_date_for_display(booking_data.get('Date', ''))
    
    # Admin request email
    admin_subject = f"💰 PAYMENT REQUEST - Booking ID: {booking_data.get('Booking ID', 'N/A')}"
    admin_body = f"""
🆕 NEW PAYMENT SUBMITTED - VERIFICATION REQUIRED

📋 **Booking Details:**
- Booking ID: {booking_data.get('Booking ID', 'N/A')}
- Customer: {booking_data.get('Name', '')}
- Date: {formatted_date}
- Time: {booking_data.get('Time Slot', '')}
- Duration: {booking_data.get('Duration', '2')} hours
- Package: {booking_data.get('Package', '')}
- Total Amount: ${booking_data.get('Total Price', '0.00')}

📧 **Customer Contact:**
- Email: {booking_data.get('Email', '')}
- Phone: {booking_data.get('Phone', '')}
- Company: {booking_data.get('Company', 'Not provided')}

💳 **Payment Information:**
- Status: {booking_data.get('Payment Status', 'PAYMENT_SUBMITTED')}
- Base Price: ${booking_data.get('Base Price', '0.00')}
- Surcharge: ${booking_data.get('Saturday Surcharge', '0.00')}
- Total: ${booking_data.get('Total Price', '0.00')}
- Payment Marked At: {booking_data.get('Payment Marked At', '')}

⚠️ **ACTION REQUIRED:**
1. Verify the payment manually
2. Go to Admin Panel: /admin
3. Click "Confirm" button next to this booking
4. User will receive confirmation email automatically

🔗 **Quick Links:**
- Admin Panel: http://127.0.0.1:5000/admin
- View Booking: Check Excel file for details

📞 **Customer Special Requests:**
{booking_data.get('Requirements', 'None')}
"""
    
    try:
        admin_msg = Message(
            subject=admin_subject,
            recipients=['sanjanastudys@gmail.com'],  # Admin email
            body=admin_body
        )
        mail.send(admin_msg)
        
        app.logger.info(f"Admin request email sent for booking {booking_data.get('Booking ID')}")
        return True
        
    except Exception as e:
        app.logger.error(f"Failed to send admin request email: {str(e)}")
        return False

def send_user_confirmation_email(booking_data):
    """Send booking confirmation email to user after admin confirmation"""
    
    formatted_date = format_date_for_display(booking_data.get('Date', ''))
    
    user_subject = f"✅ Booking Confirmed - {booking_data.get('Booking ID', '')} - INWMH Studios"
    user_body = f"""
🎉 YOUR BOOKING IS CONFIRMED! 🎉

Dear {booking_data.get('Name', 'Customer')},

We're delighted to confirm your booking at INWMH Studios!

📋 **CONFIRMATION DETAILS:**
- Confirmation Number: {booking_data.get('Booking ID', '')}
- Status: ✅ CONFIRMED
- Date: {formatted_date}
- Time: {booking_data.get('Time Slot', '')}
- Duration: {booking_data.get('Duration', '2')} hours
- Setup: {booking_data.get('Setup', '')}
- Package: {booking_data.get('Package', '')}
- People: {booking_data.get('People', '1')}
- Total Paid: ${booking_data.get('Total Price', '0.00')}

📍 **STUDIO INFORMATION:**
- Address: [Your Studio Address]
- Arrival Time: Please arrive 15 minutes before your session
- Contact: sanjanastudys@gmail.com

🎯 **PREPARATION CHECKLIST:**
✓ Bring any necessary equipment
✓ Have your booking ID ready: {booking_data.get('Booking ID', '')}
✓ Arrive 15 minutes early for setup
✓ Contact us if you need to reschedule (24+ hours notice required)

❓ **NEED HELP?**
Email: sanjanastudys@gmail.com
Phone: [Your Contact Number]

We look forward to welcoming you to INWMH Studios!

Best regards,
The INWMH Studios Team
www.inwmhstudios.com
"""
    
    try:
        user_msg = Message(
            subject=user_subject,
            recipients=[booking_data.get("Email")],
            body=user_body
        )
        mail.send(user_msg)
        
        app.logger.info(f"Confirmation email sent to {booking_data.get('Email')}")
        return True
        
    except Exception as e:
        app.logger.error(f"Failed to send confirmation email: {str(e)}")
        return False

# ============================
# PUBLIC ROUTES
# ============================
@app.route("/")
def home():
    """Home page"""
    return render_template("index.html")

@app.route("/booked_slots")
def booked_slots():
    """Get booked slots for a specific date"""
    date_str = request.args.get("date")
    if not date_str:
        return jsonify({"booked": [], "off_day": False}), 400
    
    try:
        bookings_df, offdays_df = load_excel_sheets()
        
        # Filter bookings for the selected date
        date_bookings = bookings_df[bookings_df["Date"] == date_str]
        
        # Create list of booked slots with duration info
        booked_slots_with_duration = []
        for _, row in date_bookings.iterrows():
            start_time = str(row.get("Time Slot", ""))
            duration_str = str(row.get("Duration", "2"))
            
            try:
                duration = int(duration_str)
            except ValueError:
                duration = 2
            
            if start_time in TIMESLOTS:
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
        
    except Exception as e:
        app.logger.error(f"Error in booked_slots: {str(e)}")
        return jsonify({"booked": [], "off_day": False, "error": str(e)}), 500

@app.route("/submit", methods=["POST"])
def submit_form():
    """Handle booking form submission"""
    with booking_lock:
        try:
            data = request.get_json()
            if not data:
                return jsonify({"message": "No data received"}), 400
            
            # Validate required fields
            required_fields = ["name", "email", "phone", "setup", "package", "date", "time_slot"]
            missing_fields = [field for field in required_fields if field not in data or not data[field]]
            if missing_fields:
                return jsonify({
                    "message": f"Missing required fields: {', '.join(missing_fields)}"
                }), 400
            
            # Load Excel data
            bookings_df, offdays_df = load_excel_sheets()
            
            # Check Sunday (studio closed)
            try:
                date_obj = datetime.strptime(data["date"], "%Y-%m-%d")
                if date_obj.weekday() == 6:  # Sunday
                    return jsonify({"message": "Studio is closed on Sundays"}), 400
            except ValueError:
                return jsonify({"message": "Invalid date format"}), 400
            
            # Validate time slot
            if data["time_slot"] not in TIMESLOTS:
                return jsonify({"message": "Invalid time slot"}), 400
            
            # Get duration
            duration = int(data.get("duration", 2))
            
            # Check for overlap
            date_str = data["date"]
            existing_bookings = bookings_df[bookings_df["Date"] == date_str]
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
            
            # Calculate prices
            base_price = float(data.get("base_price", 0))
            saturday_surcharge = float(data.get("saturday_surcharge", 0))
            total_price = float(data.get("total_price", 0))
            
            # Generate booking ID
            booking_id = generate_booking_id()
            
            # Create new booking entry
            new_entry = {
                "Booking ID": booking_id,
                "Timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "Name": data.get("name", ""),
                "Email": data.get("email", ""),
                "Phone": data.get("phone", ""),
                "Company": data.get("company", ""),
                "Setup": data.get("setup", ""),
                "People": data.get("people", "1"),
                "Package": data.get("package", ""),
                "Date": data.get("date", ""),
                "Time Slot": data.get("time_slot", ""),
                "Duration": str(duration),
                "Base Price": f"{base_price:.2f}",
                "Saturday Surcharge": f"{saturday_surcharge:.2f}",
                "Total Price": f"{total_price:.2f}",
                "Payment Status": "PENDING_PAYMENT",
                "Payment Marked At": "",
                "Requirements": data.get("requirements", ""),
                "Referral": data.get("referral", "")
            }
            
            # Save booking
            bookings_df = pd.concat(
                [bookings_df, pd.DataFrame([new_entry])],
                ignore_index=True
            )
            
            if not save_excel_sheets(bookings_df, offdays_df):
                return jsonify({"message": "Failed to save booking"}), 500
            
            # Send confirmation email
            email_sent = send_booking_confirmation(new_entry)
            
            return jsonify({
                "message": "Booking saved successfully. Payment pending.",
                "booking_id": booking_id,
                "booking": new_entry,
                "email_sent": email_sent
            }), 200
            
        except ValueError as e:
            app.logger.error(f"Value error in submit_form: {str(e)}")
            return jsonify({"message": "Invalid data format"}), 400
        except Exception as e:
            app.logger.error(f"Error in submit_form: {str(e)}")
            traceback.print_exc()
            return jsonify({"message": "Server error occurred"}), 500

@app.route("/payment/<booking_id>")
def payment_page(booking_id):
    """Display payment page for a booking"""
    try:
        bookings_df, _ = load_excel_sheets()
        booking = bookings_df[bookings_df["Booking ID"] == booking_id]
        
        if booking.empty:
            return render_template("error.html", message="Invalid Booking ID"), 404
        
        booking_dict = booking.iloc[0].to_dict()
        # Replace NaN with empty string for template
        for key, value in booking_dict.items():
            if pd.isna(value):
                booking_dict[key] = ""
        
        return render_template("payment.html", booking=booking_dict)
    except Exception as e:
        app.logger.error(f"Error in payment_page: {str(e)}")
        return render_template("error.html", message="Error loading payment page"), 500

@app.route("/payment_completed/<booking_id>", methods=["POST"])
def payment_completed(booking_id):
    """Handle payment completion"""
    try:
        bookings_df, offdays_df = load_excel_sheets()
        
        # Find booking index
        idx = bookings_df.index[bookings_df["Booking ID"] == booking_id].tolist()
        if not idx:
            return jsonify({"message": "Invalid booking ID"}), 404
        
        # Update payment status
        bookings_df.loc[idx[0], "Payment Status"] = "PAYMENT_SUBMITTED"
        bookings_df.loc[idx[0], "Payment Marked At"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        # Save changes
        if not save_excel_sheets(bookings_df, offdays_df):
            return jsonify({"message": "Failed to update payment status"}), 500
        
        # Get booking data for email
        booking_data = bookings_df.loc[idx[0]].to_dict()
        
        # Send payment emails (existing)
        payment_email_sent = send_payment_emails(booking_data)
        
        # Send admin request email (NEW)
        admin_request_sent = send_admin_request_email(booking_data)
        
        return jsonify({
            "message": "Payment marked for verification",
            "payment_email_sent": payment_email_sent,
            "admin_request_sent": admin_request_sent
        }), 200
        
    except Exception as e:
        app.logger.error(f"Error in payment_completed: {str(e)}")
        return jsonify({"message": "Server error occurred"}), 500

# ============================
# ADMIN ROUTES
# ============================
@app.route("/login", methods=["GET", "POST"])
def login():
    """Admin login page"""
    if request.method == "POST":
        username = request.form.get("username", "").strip()
        password = request.form.get("password", "").strip()
        
        if username == ADMIN_USERNAME and password == ADMIN_PASSWORD:
            session["admin"] = True
            session.permanent = True
            return redirect("/admin")
        
        return render_template("login.html", error="Invalid credentials")
    
    return render_template("login.html")

@app.route("/logout")
def logout():
    """Admin logout"""
    session.pop("admin", None)
    return redirect("/login")

@app.route("/admin")
def admin_dashboard():
    """Admin dashboard"""
    if "admin" not in session:
        return redirect("/login")
    return render_template("admin.html")

@app.route("/api/bookings")
def get_all_bookings():
    """API endpoint to get all bookings (admin only)"""
    if "admin" not in session:
        return jsonify({"message": "Unauthorized"}), 401
    
    try:
        bookings_df, _ = load_excel_sheets()
        
        # Replace NaN with empty string for JSON serialization
        bookings_df = bookings_df.where(pd.notna(bookings_df), "")
        
        # Convert to list of dictionaries
        bookings = bookings_df.to_dict(orient="records")
        
        return jsonify({
            "status": "success",
            "count": len(bookings),
            "bookings": bookings
        })
    except Exception as e:
        app.logger.error(f"Error in get_all_bookings: {str(e)}")
        return jsonify({
            "status": "error",
            "message": str(e)
        }), 500

@app.route("/api/bookings/<booking_id>", methods=["DELETE"])
def delete_booking(booking_id):
    """Delete a booking (admin only)"""
    if "admin" not in session:
        return jsonify({"message": "Unauthorized"}), 401
    
    try:
        bookings_df, offdays_df = load_excel_sheets()
        
        # Find booking by ID
        idx = bookings_df.index[bookings_df["Booking ID"] == booking_id].tolist()
        if not idx:
            return jsonify({"message": "Booking not found"}), 404
        
        # Remove booking
        bookings_df = bookings_df.drop(idx[0]).reset_index(drop=True)
        
        if save_excel_sheets(bookings_df, offdays_df):
            return jsonify({"message": "Booking deleted successfully"}), 200
        else:
            return jsonify({"message": "Failed to save changes"}), 500
            
    except Exception as e:
        app.logger.error(f"Error in delete_booking: {str(e)}")
        return jsonify({"message": str(e)}), 500

@app.route("/api/bookings/<booking_id>/confirm", methods=["POST"])
def confirm_booking(booking_id):
    """Admin confirms booking and sends confirmation email to user"""
    if "admin" not in session:
        return jsonify({"message": "Unauthorized"}), 401
    
    try:
        bookings_df, offdays_df = load_excel_sheets()
        
        # Find booking by ID
        idx = bookings_df.index[bookings_df["Booking ID"] == booking_id].tolist()
        if not idx:
            return jsonify({"message": "Booking not found"}), 404
        
        # Check if booking is in payment submitted status
        current_status = bookings_df.loc[idx[0], "Payment Status"]
        if current_status != "PAYMENT_SUBMITTED":
            return jsonify({
                "message": f"Cannot confirm booking with status: {current_status}"
            }), 400
        
        # Update payment status to CONFIRMED
        bookings_df.loc[idx[0], "Payment Status"] = "CONFIRMED"
        
        # Save changes
        if not save_excel_sheets(bookings_df, offdays_df):
            return jsonify({"message": "Failed to save changes"}), 500
        
        # Get booking data for email
        booking_data = bookings_df.loc[idx[0]].to_dict()
        
        # Send confirmation email to user
        email_sent = send_user_confirmation_email(booking_data)
        
        return jsonify({
            "message": "Booking confirmed successfully",
            "email_sent": email_sent,
            "booking": booking_data
        }), 200
        
    except Exception as e:
        app.logger.error(f"Error in confirm_booking: {str(e)}")
        return jsonify({"message": str(e)}), 500

# ============================
# ERROR HANDLERS
# ============================
@app.errorhandler(404)
def not_found(error):
    """Handle 404 errors"""
    return render_template("error.html", message="Page not found"), 404

@app.errorhandler(500)
def server_error(error):
    """Handle 500 errors"""
    app.logger.error(f"Server error: {str(error)}")
    return render_template("error.html", message="Internal server error"), 500

# ============================
# START SERVER
# ============================
if __name__ == "__main__":
    # Initialize Excel file
    initialize_excel()
    
    print("=" * 50)
    print("🚀 INWMH Studios Booking System")
    print(f"📁 Data file: {EXCEL_FILE}")
    print("📧 Email notifications: ACTIVE")
    print("📋 Admin confirmation system: ACTIVE")
    print("🔒 Admin panel: /admin")
    print("=" * 50)
    print("Server running on http://127.0.0.1:5000")
    print("Press Ctrl+C to stop")
    
    app.run(debug=True, host='0.0.0.0', port=5000)
