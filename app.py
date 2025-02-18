from flask import Flask, render_template, request, redirect, url_for
import sqlite3
import subprocess
import os

# Flask app
app = Flask(__name__)

# Paths
attendance_file = os.path.join(os.getcwd(), "Attendance.xlsx")
db_path = os.path.join(os.getcwd(), "AttendanceDatabase.db")
face_recognition_script = os.path.join(os.getcwd(), "Automatic Attendance System.py")

# Homepage
@app.route('/')
def index():
    return render_template('index.html')

# View Attendance
@app.route('view_attendance.html')
def view_attendance():
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()
    cursor.execute("SELECT StudentName, MonthYear, TotalPresentDays FROM MonthlyAttendance")
    attendance_data = cursor.fetchall()
    conn.close()
    return render_template('view_attendance.html', attendance_data=attendance_data)

# Run Face Recognition and Mark Attendance
@app.route('/mark', methods=['POST'])
def mark_attendance():
    try:
        # Run the face recognition script
        subprocess.run(['python', face_recognition_script], check=True)
        return render_template('success.html', message="Attendance marked successfully!")
    except subprocess.CalledProcessError as e:
        return render_template('success.html', message=f"Error marking attendance: {e}")

# Update Total Attendance in Database
@app.route('/update')
def update_total():
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()

    # Example logic for updating database
    cursor.execute("UPDATE MonthlyAttendance SET TotalPresentDays = TotalPresentDays + 1 WHERE StudentName = 'John Doe'")
    conn.commit()
    conn.close()

    return render_template('success.html', message="Database updated successfully!")

if __name__ == '__main__':
    app.run(debug=True)
