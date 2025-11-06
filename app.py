from flask import Flask, render_template, request, jsonify, redirect
import pandas as pd
import os, uuid
from datetime import datetime
import shutil

app = Flask(__name__)

EXCEL_FILE = 'bookings.xlsx'
ADMIN_PASSWORD = "studio123"  # Change this in production

# 🔹 Initialize Excel file
def initialize_excel():
    if not os.path.exists(EXCEL_FILE):
        df = pd.DataFrame(columns=[
            "id", "name", "email", "phone", "company",
            "setup", "people", "experience",
            "package", "addons", "date", "duration", "time_slot",
            "frequency", "requirements", "referral", "status"
        ])
        df.to_excel(EXCEL_FILE, index=False)

initialize_excel()

# 🔹 Data validation
def validate_booking_data(form):
    errors = []
    
    # Required fields
    required_fields = ['name', 'email', 'phone', 'setup', 'date', 'time_slot']
    for field in required_fields:
        if not form.get(field):
            errors.append(f"{field} is required")
    
    # Email validation
    email = form.get('email')
    if email and '@' not in email:
        errors.append("Valid email is required")
    
    # Date validation
    try:
        booking_date = datetime.strptime(form.get('date'), '%Y-%m-%d').date()
        if booking_date < datetime.now().date():
            errors.append("Booking date cannot be in the past")
    except (ValueError, TypeError):
        errors.append("Invalid date format")
    
    return errors

# 🔹 Backup system
def backup_data():
    """Create backup of Excel file"""
    if os.path.exists(EXCEL_FILE):
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        backup_name = f"backups/booking_backup_{timestamp}.xlsx"
        os.makedirs("backups", exist_ok=True)
        shutil.copy2(EXCEL_FILE, backup_name)
        print(f"✅ Backup created: {backup_name}")

# 🔹 Email confirmation (placeholder)
def send_confirmation_email(booking_data):
    """Send booking confirmation email"""
    message = f"""
    Thank you for your booking with INWMH Studios!
    
    Booking Details:
    - Name: {booking_data['name']}
    - Date: {booking_data['date']}
    - Time: {booking_data['time_slot']}
    - Setup: {booking_data['setup']}
    - Package: {booking_data['package']}
    
    We'll contact you shortly to confirm your reservation.
    
    Booking ID: {booking_data['id']}
    """
    
    # In a real application, integrate with an email service
    print("=" * 50)
    print("CONFIRMATION EMAIL WOULD BE SENT:")
    print(message)
    print("=" * 50)

# ------------------------------
# ROUTES
# ------------------------------

# ✅ 1. Booking form (main page)
@app.route('/')
def booking_form():
    return render_template('booking.html')

# ✅ 2. API to get booked slots for a given date
@app.route('/api/booked_slots')
def api_booked_slots():
    date = request.args.get('date')
    if not date:
        return jsonify([])
    
    try:
        df = pd.read_excel(EXCEL_FILE)
        df = df[df['date'] == date]
        df = df[df['status'] == 'booked']

        slots = []
        for _, row in df.iterrows():
            try:
                # Extract start time and duration
                time_slot_str = str(row['time_slot'])
                start = time_slot_str.split(' ')[0]
                duration = int(row['duration']) if pd.notna(row['duration']) else 1
                
                # Calculate end time
                start_hour = int(start.split(':')[0])
                end_hour = start_hour + duration
                end_time = f"{end_hour:02d}:{start.split(':')[1]}"
                
                slots.append({"start": start, "end": end_time})
            except Exception as e:
                print(f"Error processing slot: {e}")
                continue

        return jsonify(slots)
    except Exception as e:
        print(f"Error in booked_slots API: {e}")
        return jsonify([])

# ✅ 3. Submit booking form
@app.route('/submit', methods=['POST'])
def submit_booking():
    try:
        form = request.form.to_dict(flat=True)
        
        # Validate data
        errors = validate_booking_data(form)
        if errors:
            return jsonify({"success": False, "errors": errors})
        
        df = pd.read_excel(EXCEL_FILE)

        # Generate unique ID
        new_id = str(uuid.uuid4())[:8]
        
        # Handle package selection (multiple checkboxes)
        packages = request.form.getlist('package')
        
        new_entry = {
            "id": new_id,
            "name": form.get("name"),
            "email": form.get("email"),
            "phone": form.get("phone"),
            "company": form.get("company", ""),
            "setup": form.get("setup"),
            "people": form.get("people", ""),
            "experience": form.get("experience", ""),
            "package": ', '.join(packages) if packages else "",
            "addons": form.get("addons", ""),
            "date": form.get("date"),
            "duration": form.get("duration", 1),
            "time_slot": form.get("time_slot"),
            "frequency": form.get("frequency", ""),
            "requirements": form.get("requirements", ""),
            "referral": form.get("referral", ""),
            "status": "booked"
        }

        # Add new booking to dataframe
        df = pd.concat([df, pd.DataFrame([new_entry])], ignore_index=True)
        df.to_excel(EXCEL_FILE, index=False)
        
        # Create backup
        backup_data()
        
        # Send confirmation email
        send_confirmation_email(new_entry)
        
        print(f"✅ New booking created: {new_id} - {form.get('name')}")
        return jsonify({"success": True, "booking_id": new_id})
        
    except Exception as e:
        print("❌ Error saving booking:", e)
        return jsonify({"success": False, "error": str(e)})

# ✅ 4. Admin dashboard (view + cancel)
@app.route('/admin')
def admin_dashboard():
    # Simple password protection
    password = request.args.get('password')
    if password != ADMIN_PASSWORD:
        return """
        <h2>Admin Login Required</h2>
        <form method="GET">
            <input type="password" name="password" placeholder="Enter admin password" required>
            <button type="submit">Login</button>
        </form>
        """, 401
    
    try:
        df = pd.read_excel(EXCEL_FILE)
        df = df.fillna("")
        
        # Convert to records and sort by date (newest first)
        bookings = df.to_dict(orient="records")
        bookings.sort(key=lambda x: x.get('date', ''), reverse=True)
        
        # Calculate stats
        total_bookings = len(bookings)
        active_bookings = len([b for b in bookings if b.get('status') == 'booked'])
        cancelled_bookings = len([b for b in bookings if b.get('status') == 'cancelled'])
        
        return render_template("admin.html", 
                             bookings=bookings,
                             total_bookings=total_bookings,
                             active_bookings=active_bookings,
                             cancelled_bookings=cancelled_bookings)
    except Exception as e:
        return f"Error loading admin dashboard: {e}", 500

# ✅ 5. Cancel booking by ID
@app.route('/cancel/<booking_id>', methods=['POST'])
def cancel_booking(booking_id):
    # Password protection for cancel action
    password = request.args.get('password') or request.form.get('password')
    if password != ADMIN_PASSWORD:
        return "Unauthorized", 401
    
    try:
        df = pd.read_excel(EXCEL_FILE)

        if booking_id in df['id'].values:
            df.loc[df['id'] == booking_id, 'status'] = 'cancelled'
            df.to_excel(EXCEL_FILE, index=False)
            
            # Create backup after cancellation
            backup_data()
            
            print(f"✅ Booking {booking_id} cancelled successfully.")
        else:
            print(f"⚠️ Booking {booking_id} not found.")

        return redirect(f'/admin?password={ADMIN_PASSWORD}')
    except Exception as e:
        print(f"Error cancelling booking: {e}")
        return f"Error cancelling booking: {e}", 500

# ✅ 6. Health check endpoint
@app.route('/health')
def health_check():
    return jsonify({
        "status": "healthy",
        "bookings_file_exists": os.path.exists(EXCEL_FILE),
        "total_bookings": len(pd.read_excel(EXCEL_FILE)) if os.path.exists(EXCEL_FILE) else 0
    })

# ✅ 7. Get booking by ID (API endpoint)
@app.route('/api/booking/<booking_id>')
def get_booking(booking_id):
    try:
        df = pd.read_excel(EXCEL_FILE)
        booking = df[df['id'] == booking_id]
        
        if booking.empty:
            return jsonify({"error": "Booking not found"}), 404
            
        return jsonify(booking.iloc[0].to_dict())
    except Exception as e:
        return jsonify({"error": str(e)}), 500

# ------------------------------
# ERROR HANDLERS
# ------------------------------

@app.errorhandler(404)
def not_found(error):
    return jsonify({"error": "Resource not found"}), 404

@app.errorhandler(500)
def internal_error(error):
    return jsonify({"error": "Internal server error"}), 500

# ------------------------------
# RUN APP
# ------------------------------
if __name__ == '__main__':
    print("🚀 Starting INWMH Studios Booking System...")
    print(f"📊 Data file: {EXCEL_FILE}")
    print(f"🔐 Admin access: /admin?password={ADMIN_PASSWORD}")
    print("📍 Booking form: /")
    app.run(debug=True, host='0.0.0.0', port=5000)
