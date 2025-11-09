from flask import Flask, render_template, request, jsonify, redirect, session, send_file
import pandas as pd
import os, uuid, shutil
from datetime import datetime
import io
import json
import traceback

# ----------------------------
# Configuration
# ----------------------------
app = Flask(__name__)
app.secret_key = "inwmh_secret_123"

EXCEL_FILE = 'bookings.xlsx'
ADMIN_PASSWORD = "studio123"

# ----------------------------
# Excel Setup
# ----------------------------
def initialize_excel():
    if not os.path.exists(EXCEL_FILE):
        df = pd.DataFrame(columns=[
            "id", "name", "email", "phone", "company",
            "setup", "people", 
            "package", "addons", "date", "duration", "time_slot",
            "frequency", "requirements", "referral", "status",
            "total_cost", "addons_cost", "package_cost", "created_at"
        ])
        df.to_excel(EXCEL_FILE, index=False)
        print("✅ Excel file initialized")

initialize_excel()

def backup_data():
    if os.path.exists(EXCEL_FILE):
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        backup_name = f"backups/booking_backup_{timestamp}.xlsx"
        os.makedirs("backups", exist_ok=True)
        shutil.copy2(EXCEL_FILE, backup_name)

def generate_booking_id():
    """Generate a unique booking ID with format INWMH-XXXXX"""
    timestamp = datetime.now().strftime("%H%M%S")
    random_part = str(uuid.uuid4())[:3].upper()
    return f"INWMH-{timestamp}-{random_part}"

# ----------------------------
# Package and Add-ons Pricing
# ----------------------------
PACKAGE_PRICES = {
    "Basic Package": 99,
    "Professional Package": 199,
    "Premium Package": 349
}

ADDON_PRICES = {
    "Green Screen Setup": 49,
    "Professional Audio Mixing": 79,
    "4K Recording Upgrade": 99,
    "Live Streaming Setup": 129,
    "Teleprompter Rental": 39,
    "Additional Camera Operator": 89
}

def calculate_costs(packages, addons_data):
    """Calculate total cost, package cost, and add-ons cost"""
    package_cost = sum(PACKAGE_PRICES.get(pkg, 0) for pkg in packages)
    
    addons_cost = 0
    if addons_data:
        try:
            addons_list = json.loads(addons_data)
            addons_cost = sum(ADDON_PRICES.get(addon.get('name', ''), 0) for addon in addons_list)
        except (json.JSONDecodeError, TypeError):
            # If addons_data is not JSON, try to parse as string
            addons_cost = 0
    
    total_cost = package_cost + addons_cost
    return total_cost, package_cost, addons_cost

# ----------------------------
# Validation + Email (Simulated)
# ----------------------------
def validate_booking_data(form):
    errors = []
    required_fields = ['name', 'email', 'phone', 'setup', 'date', 'time_slot', 'duration']
    for field in required_fields:
        if not form.get(field):
            errors.append(f"{field.replace('_', ' ').title()} is required")

    email = form.get('email')
    if email and '@' not in email:
        errors.append("Valid email is required")

    try:
        booking_date = datetime.strptime(form.get('date'), '%Y-%m-%d').date()
        if booking_date < datetime.now().date():
            errors.append("Booking date cannot be in the past")
    except (ValueError, TypeError):
        errors.append("Invalid date format")

    # Check if at least one package is selected
    packages = request.form.getlist('package')
    if not packages:
        errors.append("At least one package must be selected")

    return errors

def send_confirmation_email(booking_data):
    message = f"""
    Booking Confirmed with INWMH Studios!

    Booking ID: {booking_data['id']}
    Name: {booking_data['name']}
    Date: {booking_data['date']}
    Time: {booking_data['time_slot']}
    Setup: {booking_data['setup']}
    Package: {booking_data['package']}
    Total Cost: ${booking_data.get('total_cost', 0)}
    """
    print("=" * 40)
    print("📧 Confirmation email would be sent:")
    print(message)
    print("=" * 40)

# ==========================================================
# ROUTES
# ==========================================================
@app.route('/')
def booking_form():
    return render_template('booking.html')

# --- Hidden Admin Login ---
@app.route('/hidden-login', methods=['GET', 'POST'])
def hidden_login():
    if request.method == 'POST':
        if request.form.get('password') == ADMIN_PASSWORD:
            session['is_admin'] = True
            return redirect('/admin')
        return render_template('admin_login.html', error="Invalid password")
    return render_template('admin_login.html')

# --- Debug File Info ---
@app.route('/debug-file')
def debug_file():
    file_info = {
        "file_exists": os.path.exists(EXCEL_FILE),
        "file_size": os.path.getsize(EXCEL_FILE) if os.path.exists(EXCEL_FILE) else 0,
        "current_directory": os.getcwd(),
        "files_in_directory": os.listdir('.')
    }
    return jsonify(file_info)

# --- Create Test Booking ---
@app.route('/create-test-booking')
def create_test_booking():
    """Create a test booking for debugging"""
    try:
        test_booking = {
            "id": "INWMH-TEST-001",
            "name": "Test User",
            "email": "test@example.com",
            "phone": "123-456-7890",
            "company": "Test Company",
            "setup": "Podcast Recording",
            "people": "2",
            "package": "Basic Package",
            "addons": "",
            "date": datetime.now().strftime("%Y-%m-%d"),
            "duration": 2,
            "time_slot": "10:00 (2 hours)",
            "frequency": "One-time",
            "requirements": "Test requirements",
            "referral": "Google",
            "status": "booked",
            "total_cost": 99,
            "package_cost": 99,
            "addons_cost": 0,
            "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        }
        
        df = pd.read_excel(EXCEL_FILE)
        df = pd.concat([df, pd.DataFrame([test_booking])], ignore_index=True)
        df.to_excel(EXCEL_FILE, index=False)
        
        return "Test booking created! <a href='/admin'>Go to Admin</a>"
    except Exception as e:
        return f"Error creating test booking: {e}"

# --- Admin Dashboard ---
@app.route('/admin')
def admin_dashboard():
    if not session.get('is_admin'):
        return redirect('/hidden-login')
    
    try:
        # Debug: Check if file exists
        if not os.path.exists(EXCEL_FILE):
            print("❌ Excel file not found!")
            return render_template("admin.html",
                                   bookings=[],
                                   total_bookings=0,
                                   active_bookings=0,
                                   cancelled_bookings=0,
                                   total_revenue=0,
                                   datetime=datetime)

        # Read from Excel file
        df = pd.read_excel(EXCEL_FILE)
        
        # Debug print
        print(f"📊 Excel file loaded: {len(df)} rows")
        print(f"📋 Columns: {list(df.columns)}")
        
        # Check if dataframe is empty or has issues
        if df.empty:
            print("⚠️ Excel file is empty - no data rows")
            bookings = []
        else:
            # Fill NaN values with empty strings and handle data types
            df = df.fillna('')
            
            # Convert all columns to string to avoid serialization issues
            for col in df.columns:
                df[col] = df[col].astype(str)
            
            bookings = df.to_dict('records')
            print(f"📖 Sample booking data: {bookings[0] if bookings else 'None'}")
        
        total_bookings = len(bookings)
        active_bookings = len([b for b in bookings if b.get('status') == 'booked'])
        cancelled_bookings = len([b for b in bookings if b.get('status') == 'cancelled'])
        
        # Calculate total revenue safely
        total_revenue = 0
        for b in bookings:
            if b.get('status') == 'booked':
                try:
                    total_revenue += float(b.get('total_cost', 0))
                except (ValueError, TypeError):
                    continue
        
        # Sort by date safely
        try:
            bookings.sort(key=lambda x: x.get('date', ''), reverse=True)
        except:
            # If sorting fails, keep original order
            pass

        print(f"🎯 Final stats - Total: {total_bookings}, Active: {active_bookings}, Cancelled: {cancelled_bookings}")

        return render_template("admin.html",
                               bookings=bookings,
                               total_bookings=total_bookings,
                               active_bookings=active_bookings,
                               cancelled_bookings=cancelled_bookings,
                               total_revenue=total_revenue,
                               datetime=datetime)
    except Exception as e:
        print(f"❌ Error loading admin dashboard: {e}")
        traceback.print_exc()
        return f"Error loading admin dashboard: {e}", 500

# --- Download Excel ---
@app.route('/download-excel')
def download_excel():
    if not session.get('is_admin'):
        return "Unauthorized", 401
    try:
        return send_file(EXCEL_FILE, 
                        as_attachment=True, 
                        download_name=f"INWMH_Bookings_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    except Exception as e:
        return f"Error downloading file: {e}", 500

# --- Reset Excel ---
@app.route('/reset-excel')
def reset_excel():
    if not session.get('is_admin'):
        return "Unauthorized", 401
    
    # Backup current file
    if os.path.exists(EXCEL_FILE):
        backup_data()
        os.remove(EXCEL_FILE)
    
    # Reinitialize
    initialize_excel()
    return "Excel file reset successfully"

# --- Logout ---
@app.route('/logout')
def logout():
    session.pop('is_admin', None)
    return redirect('/')

# --- API: Get Booked Slots ---
@app.route('/api/booked_slots')
def api_booked_slots():
    date = request.args.get('date')
    if not date:
        return jsonify([])
    try:
        df = pd.read_excel(EXCEL_FILE)
        # Convert date to string for comparison (Excel dates might be stored differently)
        df['date_str'] = df['date'].astype(str)
        bookings = df[(df['date_str'] == date) & (df['status'] == 'booked')].to_dict('records')
        
        slots = []
        for b in bookings:
            time_slot = b.get('time_slot', '')
            if time_slot:
                # Extract start time from time_slot format like "10:30 (2 hours)"
                start = time_slot.split(' ')[0]
                duration = int(b.get('duration', 1))
                try:
                    start_hour = int(start.split(':')[0])
                    end_hour = start_hour + duration
                    end_time = f"{end_hour:02d}:{start.split(':')[1]}"
                    slots.append({"start": start, "end": end_time})
                except (ValueError, IndexError):
                    continue
        return jsonify(slots)
    except Exception as e:
        print("Error fetching slots:", e)
        return jsonify([])

# --- Submit Booking ---
@app.route('/submit', methods=['POST'])
def submit_booking():
    try:
        form = request.form.to_dict(flat=True)
        errors = validate_booking_data(form)
        if errors:
            return jsonify({"success": False, "errors": errors})

        # Check for time slot conflicts
        conflict = check_time_slot_conflict(form.get('date'), form.get('time_slot'))
        if conflict:
            return jsonify({"success": False, "errors": ["This time slot is already booked"]})

        new_id = generate_booking_id()
        packages = request.form.getlist('package')
        
        # Calculate costs
        addons_data = form.get("selected_addons", "")
        total_cost, package_cost, addons_cost = calculate_costs(packages, addons_data)

        new_entry = {
            "id": new_id,
            "name": form.get("name"),
            "email": form.get("email"),
            "phone": form.get("phone"),
            "company": form.get("company", ""),
            "setup": form.get("setup"),
            "people": form.get("people", ""),
            "package": ', '.join(packages) if packages else "",
            "addons": addons_data,
            "date": form.get("date"),
            "duration": int(form.get("duration", 1)),
            "time_slot": form.get("time_slot"),
            "frequency": form.get("frequency", ""),
            "requirements": form.get("requirements", ""),
            "referral": form.get("referral", ""),
            "status": "booked",
            "total_cost": total_cost,
            "package_cost": package_cost,
            "addons_cost": addons_cost,
            "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        }

        # Save to Excel
        df = pd.read_excel(EXCEL_FILE)
        df = pd.concat([df, pd.DataFrame([new_entry])], ignore_index=True)
        df.to_excel(EXCEL_FILE, index=False)
        backup_data()

        send_confirmation_email(new_entry)
        print(f"✅ Booking saved to Excel: {new_id}")
        return jsonify({"success": True, "booking_id": new_id})
    except Exception as e:
        print("❌ Error:", e)
        traceback.print_exc()
        return jsonify({"success": False, "error": str(e)})
    
def check_time_slot_conflict(date, time_slot):
    """Check if time slot is already booked"""
    try:
        df = pd.read_excel(EXCEL_FILE)
        existing = df[(df['date'] == date) & 
                     (df['time_slot'] == time_slot) & 
                     (df['status'] == 'booked')]
        return len(existing) > 0
    except:
        return False

# --- Cancel Booking ---
@app.route('/cancel/<booking_id>', methods=['POST'])
def cancel_booking(booking_id):
    if not session.get('is_admin'):
        return "Unauthorized", 401
    try:
        df = pd.read_excel(EXCEL_FILE)
        if booking_id in df['id'].values:
            df.loc[df['id'] == booking_id, 'status'] = 'cancelled'
            df.to_excel(EXCEL_FILE, index=False)
            backup_data()
            print(f"❌ Booking {booking_id} cancelled.")
        return redirect('/admin')
    except Exception as e:
        return f"Error cancelling booking: {e}", 500

# --- Health Check ---
@app.route('/health')
def health_check():
    try:
        df = pd.read_excel(EXCEL_FILE)
        total_bookings = len(df)
        return jsonify({
            "status": "healthy",
            "total_bookings": total_bookings,
            "excel_exists": os.path.exists(EXCEL_FILE)
        })
    except Exception as e:
        return jsonify({"status": "error", "error": str(e)})

# --- API: Get Pricing Information ---
@app.route('/api/pricing')
def get_pricing():
    """API endpoint to get package and add-on pricing"""
    return jsonify({
        "packages": PACKAGE_PRICES,
        "addons": ADDON_PRICES
    })

# --- Debug Data ---
@app.route('/debug-data')
def debug_data():
    try:
        df = pd.read_excel(EXCEL_FILE)
        print("📊 Excel Data:")
        print(df)
        print(f"📈 Total rows: {len(df)}")
        
        # Convert to dict for JSON response
        data_dict = df.to_dict('records')
        
        return jsonify({
            "total_rows": len(df),
            "columns": list(df.columns),
            "data": data_dict
        })
    except Exception as e:
        return f"Error: {e}", 500

import functools

def admin_required(f):
    @functools.wraps(f)
    def decorated_function(*args, **kwargs):
        if not session.get('is_admin'):
            return redirect('/hidden-login')
        return f(*args, **kwargs)
    return decorated_function

# --- Run ---
if __name__ == '__main__':
    print("🚀 INWMH Studio Booking System (Excel Only)")
    print(f"🔐 Hidden Admin Login: /hidden-login")
    print(f"💰 Package Prices: {PACKAGE_PRICES}")
    print(f"🔧 Add-on Prices: {ADDON_PRICES}")
    print(f"📁 Excel File: {EXCEL_FILE}")
    print(f"📊 Debug Routes:")
    print(f"   /debug-file - Check file status")
    print(f"   /debug-data - View raw data")
    print(f"   /create-test-booking - Add test data")
    print(f"   /health - System health check")
    app.run(debug=True, host='0.0.0.0', port=5000)
