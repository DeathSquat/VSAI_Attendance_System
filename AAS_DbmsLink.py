import mysql.connector
import numpy as np
import face_recognition
import os
import pickle

# Connect to MySQL database
def create_db_connection():
    return mysql.connector.connect(
        host="localhost",  # replace with your MySQL server host
        user="root",       # replace with your MySQL username
        password="dzhghost",       # replace with your MySQL password
        database="attendancedatabase"
    )

# Create or load the student data
def load_or_insert_students():
    # Establish database connection
    conn = create_db_connection()
    cursor = conn.cursor()

    # Try to fetch students from the database
    cursor.execute("SELECT id, name, face_encoding FROM students")
    rows = cursor.fetchall()

    # If students exist, return their encodings and names
    if rows:
        student_encodings = []
        student_names = []
        for row in rows:
            student_encodings.append(np.frombuffer(row[2], dtype=np.float64))  # Convert binary data back to numpy array
            student_names.append(row[1])
        return student_encodings, student_names

    # If no students, return empty lists
    return [], []

# Insert a new student with face encoding
def insert_student(name, face_encoding):
    conn = create_db_connection()
    cursor = conn.cursor()

    # Convert the face encoding numpy array to binary data
    face_encoding_binary = face_encoding.tobytes()

    # Insert student into the database
    cursor.execute("INSERT INTO students (name, face_encoding) VALUES (%s, %s)", 
                   (name, face_encoding_binary))
    conn.commit()

    print(f"Student {name} added to the database.")
    conn.close()

# Load and encode faces from the given image folder
def encode_faces(image_folder):
    student_encodings = []
    student_names = []

    # Loop through all images in the folder
    for filename in os.listdir(image_folder):
        if filename.endswith('.jpg') or filename.endswith('.png'):
            image_path = os.path.join(image_folder, filename)
            image = face_recognition.load_image_file(image_path)
            face_encodings = face_recognition.face_encodings(image)

            # Ensure there's exactly one face encoding per image
            if face_encodings:
                student_encodings.append(face_encodings[0])
                student_names.append(os.path.splitext(filename)[0])  # Use the filename (without extension) as the name

    return student_encodings, student_names

# Save or update students in the database
def save_or_update_students(image_folder):
    student_encodings, student_names = load_or_insert_students()

    # Get the list of existing student names
    existing_students = set(student_names)

    # Encode faces from the image folder
    new_encodings, new_names = encode_faces(image_folder)

    for name, encoding in zip(new_names, new_encodings):
        # If the student is not already in the database, insert them
        if name not in existing_students:
            insert_student(name, encoding)
            existing_students.add(name)

    print("All students are now up-to-date.")

# Main entry point
if __name__ == "__main__":
    image_folder = "students_images"  # Folder where student images are stored
    save_or_update_students(image_folder)
