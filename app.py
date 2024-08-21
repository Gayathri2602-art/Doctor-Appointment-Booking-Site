from flask import Flask, render_template, request, redirect
import pandas as pd
import os

app = Flask(__name__)

# Path to save the Excel file
excel_file_path = 'appointments.xlsx'

# Ensure the Excel file exists
if not os.path.exists(excel_file_path):
    df = pd.DataFrame(columns=['Name', 'Email', 'Phone', 'Reason', 'Preferred Date'])
    df.to_excel(excel_file_path, index=False)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/submit', methods=['POST'])
def submit():
    # Get form data
    name = request.form['name']
    email = request.form['email']
    phone = request.form['phone']
    reason = request.form['reason']
    date = request.form['date']

    # Read existing data
    df = pd.read_excel(excel_file_path)

    # Append new data using pd.concat
    new_data = pd.DataFrame({
        'Name': [name],
        'Email': [email],
        'Phone': [phone],
        'Reason': [reason],
        'Preferred Date': [date]
    })
    df = pd.concat([df, new_data], ignore_index=True)

    # Save back to Excel
    df.to_excel(excel_file_path, index=False)

    return redirect('/')

if __name__ == '__main__':
    app.run(debug=True, port=5001)
