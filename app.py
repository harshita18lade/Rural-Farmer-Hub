import openpyxl
import os
from datetime import datetime, timedelta
from flask import Flask, render_template, request, redirect, url_for, flash, session

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY")

# -------------------------------------------------------------------
# Excel Database Helper Function
# -------------------------------------------------------------------
def save_to_excel(filename, sheet_name, data_row):
    """Opens the Excel workbook and appends a new row permanently."""
    file_path = os.path.join(os.path.abspath(os.path.dirname(__file__)), filename)
    if os.path.exists(file_path):
        wb = openpyxl.load_workbook(file_path)
        if sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            ws.append(data_row)
            wb.save(file_path)

# In-memory user database
users_db = {
    "farmer1": {"password": "123", "role": "farmer", "name": "Rajesh Kumar"},
    "chef1": {"password": "123", "role": "restaurant", "name": "Green Bite Bistro"}
}

# Farmer listings database
crop_listings = [
    {"id": 1, "farmer": "Rajesh Kumar", "category": "Vegetables", "name": "Organic Roma Tomatoes", "quantity": 250, "price": 35},
    {"id": 2, "farmer": "Sunita Patil", "category": "Vegetables", "name": "Fresh Spinach", "quantity": 100, "price": 25},
    {"id": 3, "farmer": "Harpreet Singh", "category": "Grains", "name": "Basmati Rice", "quantity": 500, "price": 65},
    {"id": 4, "farmer": "Anil Deshmukh", "category": "Fruits", "name": "Nagpur Oranges", "quantity": 300, "price": 50}
]

# Restaurant demand/buying requests database
buy_requests = [
    {"id": 1, "restaurant": "Green Bite Bistro", "category": "Vegetables", "item_needed": "Fresh Onions", "quantity_needed": 100, "credit_term": "Net-15 Approved"},
    {"id": 2, "restaurant": "Spice Garden Hotel", "category": "Grains", "item_needed": "Wheat Flour (Atta)", "quantity_needed": 250, "credit_term": "Net-15 Approved"},
    {"id": 3, "restaurant": "Urban Dhaba", "category": "Vegetables", "item_needed": "Potatoes", "quantity_needed": 150, "credit_term": "Net-15 Approved"},
    {"id": 4, "restaurant": "Royal Heritage Cafe", "category": "Fruits", "item_needed": "Apples", "quantity_needed": 80, "credit_term": "Net-15 Approved"}
]

@app.route('/')
def home():
    if 'username' not in session:
        return redirect(url_for('login'))
    if session.get('role') == 'farmer':
        return redirect(url_for('farmer_dashboard'))
    return redirect(url_for('restaurant_portal'))

@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        action = request.form.get('action')

        if action == 'register':
            username = request.form.get('username')
            password = request.form.get('password')
            role = request.form.get('role')
            full_name = request.form.get('full_name')

            if username in users_db:
                flash("Username already exists! Please choose another.", "danger")
            else:
                users_db[username] = {
                    "password": password,
                    "role": role,
                    "name": full_name
                }
                flash("Registration successful! You can now log in.", "success")
            return redirect(url_for('login'))

        elif action == 'login':
            username = request.form.get('username')
            password = request.form.get('password')

            user = users_db.get(username)
            if user and user['password'] == password:
                session['username'] = username
                session['role'] = user['role']
                session['name'] = user['name']
                flash(f"Welcome back, {user['name']}!", "success")

                if user['role'] == 'farmer':
                    return redirect(url_for('farmer_dashboard'))
                else:
                    return redirect(url_for('restaurant_portal'))
            else:
                flash("Invalid username or password!", "danger")

    return render_template('login.html')

@app.route('/logout')
def logout():
    session.clear()
    flash("Logged out successfully.", "info")
    return redirect(url_for('login'))

@app.route('/farmer', methods=['GET', 'POST'])
def farmer_dashboard():
    if session.get('role') != 'farmer':
        flash("Please log in as a Farmer to access this page.", "warning")
        return redirect(url_for('login'))

    recommendation = None

    if request.method == 'POST':
        action = request.form.get('action')

        if action == 'add_crop':
            crop_name = request.form.get('crop_name')
            category = request.form.get('category')
            quantity = float(request.form.get('quantity', 0))
            price = float(request.form.get('price', 0))
            farmer_name = session.get('name')

            new_id = len(crop_listings) + 1
            crop_listings.append({
                "id": new_id,
                "farmer": farmer_name,
                "category": category,
                "name": crop_name,
                "quantity": quantity,
                "price": price
            })

            # Automatically save new crop listing into Excel sheet 'Crop Catalog'
            save_to_excel("rural_farmer_hub_database.xlsx", "Crop Catalog", 
                          [new_id, crop_name, farmer_name, quantity, price, str(datetime.utcnow().date())])

            flash("Produce listing posted successfully for restaurants and saved to Excel!", "success")

        elif action == 'get_recommendation':
            water_source = request.form.get('water_source')
            if water_source == 'rainfed':
                recommendation = "Recommended: Pulses/Legumes (Arhar, Chickpeas) to avoid market glut and enrich soil."
            else:
                recommendation = "Recommended: Seasonal Vegetables or Wheat with high restaurant demand."

    farmer_crops = [c for c in crop_listings if c['farmer'] == session.get('name')]
    return render_template('farmer_dashboard.html', 
                           farmer_crops=farmer_crops, 
                           buy_requests=buy_requests,
                           recommendation=recommendation)

@app.route('/restaurant', methods=['GET', 'POST'])
def restaurant_portal():
    if session.get('role') != 'restaurant':
        flash("Please log in as a Restaurant to access this page.", "warning")
        return redirect(url_for('login'))

    category_filter = request.args.get('category', 'All')

    if request.method == 'POST':
        action = request.form.get('action')

        if action == 'post_buy_request':
            category = request.form.get('category')
            item_needed = request.form.get('item_needed')
            quantity_needed = float(request.form.get('quantity_needed', 0))
            restaurant_name = session.get('name')

            new_req_id = len(buy_requests) + 1
            buy_requests.append({
                "id": new_req_id,
                "restaurant": restaurant_name,
                "category": category,
                "item_needed": item_needed,
                "quantity_needed": quantity_needed,
                "credit_term": "Net-15 Approved"
            })

            # Automatically save restaurant credit contract/buy request into Excel sheet 'Net-15 Credit Contracts'
            due_date_calc = datetime.utcnow().date() + timedelta(days=15)
            save_to_excel("rural_farmer_hub_database.xlsx", "Net-15 Credit Contracts", 
                          [new_req_id, restaurant_name, f"27GST{new_req_id}XYZ", str(datetime.utcnow().date()), str(due_date_calc), "Active (Net-15 Approved)"])

            flash("Buy Requirement posted and Net-15 Credit Contract recorded in Excel!", "success")

        elif action == 'order_produce':
            item_name = request.form.get('item_name')
            payment_method = request.form.get('payment_method')
            order_qty = request.form.get('order_qty')
            
            flash(f"Order confirmed! Placed order for {order_qty} kg of '{item_name}' via [{payment_method}].", "success")

    filtered_crops = crop_listings if category_filter == 'All' else [c for c in crop_listings if c['category'] == category_filter]

    return render_template('restaurant_portal.html', 
                           crops=filtered_crops, 
                           current_category=category_filter,
                           my_requests=[r for r in buy_requests if r['restaurant'] == session.get('name')])

if __name__ == '__main__':
    app.run(debug=True)