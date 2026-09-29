
from flask import Flask, render_template, request, redirect, send_file
import sqlite3
from datetime import date, datetime
import openpyxl
from openpyxl import Workbook

app = Flask(__name__)

DATABASE = "library.db"
FINE_PER_DAY = 5


# ==================================================
# DATABASE CONNECTION
# ==================================================

def get_db_connection():
    conn = sqlite3.connect(DATABASE)
    conn.row_factory = sqlite3.Row
    return conn


# ==================================================
# CREATE DATABASE
# ==================================================

def create_database():

    conn = get_db_connection()

    # BOOKS TABLE
    conn.execute("""
        CREATE TABLE IF NOT EXISTS books (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            book_id TEXT UNIQUE NOT NULL,
            title TEXT NOT NULL,
            author TEXT NOT NULL,
            category TEXT,
            total_copies INTEGER NOT NULL,
            available_copies INTEGER NOT NULL
        )
    """)

    # STUDENTS TABLE
    conn.execute("""
        CREATE TABLE IF NOT EXISTS students (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            roll_no TEXT UNIQUE NOT NULL,
            name TEXT NOT NULL,
            branch TEXT,
            year TEXT,
            phone TEXT
        )
    """)

    # ISSUED BOOKS TABLE
    conn.execute("""
        CREATE TABLE IF NOT EXISTS issued_books (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            book_id TEXT NOT NULL,
            roll_no TEXT NOT NULL,
            issue_date TEXT NOT NULL,
            due_date TEXT NOT NULL,
            return_date TEXT,
            fine REAL DEFAULT 0
        )
    """)

    conn.commit()
    conn.close()


# ==================================================
# DASHBOARD
# ==================================================

@app.route("/")
def home():

    conn = get_db_connection()

    total_books = conn.execute("""
        SELECT COALESCE(SUM(total_copies), 0)
        FROM books
    """).fetchone()[0]

    available_books = conn.execute("""
        SELECT COALESCE(SUM(available_copies), 0)
        FROM books
    """).fetchone()[0]

    issued_books = conn.execute("""
        SELECT COUNT(*)
        FROM issued_books
        WHERE return_date IS NULL
    """).fetchone()[0]

    overdue_books = conn.execute("""
        SELECT COUNT(*)
        FROM issued_books
        WHERE return_date IS NULL
        AND due_date < ?
    """, (date.today().isoformat(),)).fetchone()[0]

    total_students = conn.execute("""
        SELECT COUNT(*)
        FROM students
    """).fetchone()[0]

    conn.close()

    return render_template(
        "dashboard.html",
        total_books=total_books,
        available_books=available_books,
        issued_books=issued_books,
        overdue_books=overdue_books,
        total_students=total_students
    )


# ==================================================
# ADD BOOK
# ==================================================

@app.route("/add-book", methods=["GET", "POST"])
def add_book():

    if request.method == "POST":

        book_id = request.form["book_id"]
        title = request.form["title"]
        author = request.form["author"]
        category = request.form["category"]
        total_copies = int(request.form["total_copies"])

        conn = get_db_connection()

        try:

            conn.execute("""
                INSERT INTO books (
                    book_id,
                    title,
                    author,
                    category,
                    total_copies,
                    available_copies
                )
                VALUES (?, ?, ?, ?, ?, ?)
            """, (
                book_id,
                title,
                author,
                category,
                total_copies,
                total_copies
            ))

            conn.commit()

        except sqlite3.IntegrityError:

            conn.close()

            return """
            <h2>❌ Book ID already exists.</h2>
            <br>
            <a href="/add-book">Try Again</a>
            """

        conn.close()

        return redirect("/books")

    return render_template("add_book.html")


# ==================================================
# VIEW BOOKS
# ==================================================

@app.route("/books")
def books():

    search = request.args.get("search", "")

    conn = get_db_connection()

    if search:

        books = conn.execute("""
            SELECT *
            FROM books
            WHERE book_id LIKE ?
            OR title LIKE ?
            OR author LIKE ?
            OR category LIKE ?
            ORDER BY id DESC
        """, (
            f"%{search}%",
            f"%{search}%",
            f"%{search}%",
            f"%{search}%"
        )).fetchall()

    else:

        books = conn.execute("""
            SELECT *
            FROM books
            ORDER BY id DESC
        """).fetchall()

    conn.close()

    return render_template(
        "books.html",
        books=books,
        search=search
    )


# ==================================================
# DELETE BOOK
# ==================================================

@app.route("/delete-book/<int:book_id>")
def delete_book(book_id):

    conn = get_db_connection()

    conn.execute("""
        DELETE FROM books
        WHERE id = ?
    """, (book_id,))

    conn.commit()
    conn.close()

    return redirect("/books")


# ==================================================
# STUDENTS
# ==================================================

@app.route("/students")
def students():

    search = request.args.get("search", "")

    conn = get_db_connection()

    if search:

        students = conn.execute("""
            SELECT *
            FROM students
            WHERE roll_no LIKE ?
            OR name LIKE ?
            OR branch LIKE ?
            OR year LIKE ?
            ORDER BY id DESC
        """, (
            f"%{search}%",
            f"%{search}%",
            f"%{search}%",
            f"%{search}%"
        )).fetchall()

    else:

        students = conn.execute("""
            SELECT *
            FROM students
            ORDER BY id DESC
        """).fetchall()

    conn.close()

    return render_template(
        "students.html",
        students=students,
        search=search
    )


# ==================================================
# ADD STUDENT
# ==================================================

@app.route("/add-student", methods=["GET", "POST"])
def add_student():

    if request.method == "POST":

        roll_no = request.form["roll_no"]
        name = request.form["name"]
        branch = request.form["branch"]
        year = request.form["year"]
        phone = request.form["phone"]

        conn = get_db_connection()

        try:

            conn.execute("""
                INSERT INTO students (
                    roll_no,
                    name,
                    branch,
                    year,
                    phone
                )
                VALUES (?, ?, ?, ?, ?)
            """, (
                roll_no,
                name,
                branch,
                year,
                phone
            ))

            conn.commit()

        except sqlite3.IntegrityError:

            conn.close()

            return """
            <h2>❌ Student Roll Number already exists.</h2>
            <br>
            <a href="/add-student">Try Again</a>
            """

        conn.close()

        return redirect("/students")

    return render_template("add_student.html")


# ==================================================
# DELETE STUDENT
# ==================================================

@app.route("/delete-student/<int:student_id>")
def delete_student(student_id):

    conn = get_db_connection()

    conn.execute("""
        DELETE FROM students
        WHERE id = ?
    """, (student_id,))

    conn.commit()
    conn.close()

    return redirect("/students")


# ==================================================
# ISSUE BOOK
# ==================================================

@app.route("/issue-book", methods=["GET", "POST"])
def issue_book():

    conn = get_db_connection()

    if request.method == "POST":

        book_id = request.form["book_id"]
        roll_no = request.form["roll_no"]
        issue_date = request.form["issue_date"]
        due_date = request.form["due_date"]

        # Check book
        book = conn.execute("""
            SELECT *
            FROM books
            WHERE book_id = ?
        """, (book_id,)).fetchone()

        if not book:

            conn.close()

            return """
            <h2>❌ Book ID not found.</h2>
            <br>
            <a href="/issue-book">Go Back</a>
            """

        # Check availability
        if book["available_copies"] <= 0:

            conn.close()

            return """
            <h2>❌ No copies of this book are available.</h2>
            <br>
            <a href="/issue-book">Go Back</a>
            """

        # Check student
        student = conn.execute("""
            SELECT *
            FROM students
            WHERE roll_no = ?
        """, (roll_no,)).fetchone()

        if not student:

            conn.close()

            return """
            <h2>❌ Student Roll Number not found.</h2>
            <br>
            <a href="/issue-book">Go Back</a>
            """

        # Check duplicate issue
        existing = conn.execute("""
            SELECT *
            FROM issued_books
            WHERE book_id = ?
            AND roll_no = ?
            AND return_date IS NULL
        """, (book_id, roll_no)).fetchone()

        if existing:

            conn.close()

            return """
            <h2>❌ This book is already issued to this student.</h2>
            <br>
            <a href="/issue-book">Go Back</a>
            """

        # Add issue record
        conn.execute("""
            INSERT INTO issued_books (
                book_id,
                roll_no,
                issue_date,
                due_date
            )
            VALUES (?, ?, ?, ?)
        """, (
            book_id,
            roll_no,
            issue_date,
            due_date
        ))

        # Reduce available copies
        conn.execute("""
            UPDATE books
            SET available_copies = available_copies - 1
            WHERE book_id = ?
        """, (book_id,))

        conn.commit()
        conn.close()

        return redirect("/reports")

    # Available books
    books = conn.execute("""
        SELECT *
        FROM books
        WHERE available_copies > 0
        ORDER BY title
    """).fetchall()

    # Students
    students = conn.execute("""
        SELECT *
        FROM students
        ORDER BY name
    """).fetchall()

    conn.close()

    return render_template(
        "issue_book.html",
        books=books,
        students=students,
        today=date.today().isoformat()
    )


# ==================================================
# RETURN BOOK
# ==================================================

@app.route("/return-book", methods=["GET", "POST"])
def return_book():

    conn = get_db_connection()

    if request.method == "POST":

        issue_id = request.form["issue_id"]

        record = conn.execute("""
            SELECT *
            FROM issued_books
            WHERE id = ?
            AND return_date IS NULL
        """, (issue_id,)).fetchone()

        if not record:

            conn.close()

            return """
            <h2>❌ Issue record not found.</h2>
            <br>
            <a href="/return-book">Go Back</a>
            """

        return_date = date.today()

        due_date = datetime.strptime(
            record["due_date"],
            "%Y-%m-%d"
        ).date()

        late_days = (return_date - due_date).days

        if late_days > 0:
            fine = late_days * FINE_PER_DAY
        else:
            fine = 0

        # Update return
        conn.execute("""
            UPDATE issued_books
            SET return_date = ?,
                fine = ?
            WHERE id = ?
        """, (
            return_date.isoformat(),
            fine,
            issue_id
        ))

        # Increase available copies
        conn.execute("""
            UPDATE books
            SET available_copies = available_copies + 1
            WHERE book_id = ?
        """, (record["book_id"],))

        conn.commit()
        conn.close()

        return redirect("/reports")

    # Currently issued books
    issued = conn.execute("""
        SELECT
            issued_books.id,
            issued_books.book_id,
            books.title,
            issued_books.roll_no,
            students.name,
            issued_books.issue_date,
            issued_books.due_date
        FROM issued_books
        LEFT JOIN books
            ON issued_books.book_id = books.book_id
        LEFT JOIN students
            ON issued_books.roll_no = students.roll_no
        WHERE issued_books.return_date IS NULL
        ORDER BY issued_books.id DESC
    """).fetchall()

    conn.close()

    return render_template(
        "return_book.html",
        issued=issued,
        today=date.today().isoformat()
    )


# ==================================================
# REPORTS
# ==================================================

@app.route("/reports")
def reports():

    conn = get_db_connection()

    records = conn.execute("""
        SELECT
            issued_books.id,
            issued_books.book_id,
            books.title,
            issued_books.roll_no,
            students.name,
            issued_books.issue_date,
            issued_books.due_date,
            issued_books.return_date,
            issued_books.fine
        FROM issued_books
        LEFT JOIN books
            ON issued_books.book_id = books.book_id
        LEFT JOIN students
            ON issued_books.roll_no = students.roll_no
        ORDER BY issued_books.id DESC
    """).fetchall()

    total_issued = conn.execute("""
        SELECT COUNT(*)
        FROM issued_books
    """).fetchone()[0]

    currently_issued = conn.execute("""
        SELECT COUNT(*)
        FROM issued_books
        WHERE return_date IS NULL
    """).fetchone()[0]

    total_fine = conn.execute("""
        SELECT COALESCE(SUM(fine), 0)
        FROM issued_books
    """).fetchone()[0]

    conn.close()

    return render_template(
        "reports.html",
        records=records,
        total_issued=total_issued,
        currently_issued=currently_issued,
        total_fine=total_fine
    )


# ==================================================
# DOWNLOAD EXCEL REPORT
# ==================================================

@app.route("/download-report")
def download_report():

    conn = get_db_connection()

    records = conn.execute("""
        SELECT
            issued_books.book_id,
            books.title,
            issued_books.roll_no,
            students.name,
            issued_books.issue_date,
            issued_books.due_date,
            issued_books.return_date,
            issued_books.fine
        FROM issued_books
        LEFT JOIN books
            ON issued_books.book_id = books.book_id
        LEFT JOIN students
            ON issued_books.roll_no = students.roll_no
        ORDER BY issued_books.id DESC
    """).fetchall()

    conn.close()

    # Create Excel workbook
    workbook = Workbook()

    sheet = workbook.active
    sheet.title = "Library Report"

    # Headers
    headers = [
        "Book ID",
        "Book Title",
        "Roll No",
        "Student Name",
        "Issue Date",
        "Due Date",
        "Return Date",
        "Fine"
    ]

    sheet.append(headers)

    # Add data
    for record in records:

        sheet.append([
            record["book_id"],
            record["title"],
            record["roll_no"],
            record["name"],
            record["issue_date"],
            record["due_date"],
            record["return_date"] or "Not Returned",
            record["fine"] or 0
        ])

    # Column widths
    widths = [
        15,
        30,
        20,
        25,
        15,
        15,
        18,
        12
    ]

    for index, width in enumerate(widths, start=1):

        column_letter = openpyxl.utils.get_column_letter(index)

        sheet.column_dimensions[column_letter].width = width

    # Save Excel file
    file_path = "library_report.xlsx"

    workbook.save(file_path)

    # Send Excel file to browser
    return send_file(
        file_path,
        as_attachment=True,
        download_name="library_report.xlsx"
    )


# ==================================================
# START APPLICATION
# ==================================================

if __name__ == "__main__":

    create_database()

    app.run(
        host="127.0.0.1",
        port=5000,
        debug=True
    )

