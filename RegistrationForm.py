import tkinter as tk
from tkinter import messagebox
from tkinter import ttk
import sqlite3
import pandas as pd
from datetime import datetime, timedelta
from tkcalendar import DateEntry
from PIL import Image, ImageTk
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle
from reportlab.lib import colors
import babel.numbers



def create_db():
    conn = None
    try:
        conn = sqlite3.connect(r'C:\Registration Form Code\user_full_data.db')
        c = conn.cursor()
        

        c.execute('''
            CREATE TABLE IF NOT EXISTS user_data (
                student_name TEXT NOT NULL,
                email_id TEXT NOT NULL,
                mobile_no TEXT NOT NULL,
                course TEXT NOT NULL,
                gender TEXT NOT NULL,
                reference TEXT NOT NULL,
                date TEXT NOT NULL
            )
        ''')

        
        conn.commit()

        
        
        
    except sqlite3.Error as e:
        print(f"An error occurred: {e}")
    finally:
        if conn:
            conn.close()
            


create_db()


def validate_form():
    if not entry_student_name.get():
        raise ValueError("Student Name is required.")
    if not entry_mobile_no.get():
        raise ValueError("Mobile No is required.")
    if not course_combobox.get():
        raise ValueError("Course Name is required.")
    if not gender.get():
        raise ValueError("Gender is required.")
    

def submit_data():
    try:
        validate_form()
        conn = sqlite3.connect(r'C:\Registration Form Code\user_full_data.db')
        c = conn.cursor()
        current_date = datetime.now().strftime("%Y-%m-%d")

        c.execute("INSERT INTO user_data (student_name, email_id, mobile_no, course, gender, reference, date) VALUES (?, ?, ?, ?, ?, ?, ?)",
                (entry_student_name.get(), entry_email_id.get(), entry_mobile_no.get(), course_combobox.get(), gender.get(), reference.get(), current_date))

        conn.commit()

        
        df = pd.read_sql_query("SELECT * FROM user_data", conn)
        df.to_excel(r'C:\Registration Form Code\user_full_data.xlsx', index=False)
        conn.close()

        messagebox.showinfo("Success", "Data submitted and stored successfully!")
    except ValueError as ve:
        messagebox.showerror("Error", str(ve))
    except sqlite3.Error as e:
        messagebox.showerror("Database Error", f"An error occurred while inserting data: {e}")

def export_to_excel():
    date_window = tk.Toplevel(root)
    date_window.title("Export Excel")
    window_width = 1000
    window_height = 750
    date_window['bg']='skyblue'

    logo_image = tk.PhotoImage(file='C:\\syam\\projects\\VCS1.png')

    date_window.iconphoto(False, logo_image)
    screen_width = date_window.winfo_screenwidth()
    screen_height = date_window.winfo_screenheight()
    position_x = int((screen_width / 2) - (window_width / 2))
    position_y = int((screen_height / 2) - (window_height / 2))
    date_window.geometry(f"{window_width}x{window_height}+{position_x}+{position_y}")
    
    today_date = datetime.now().date()

    tk.Label(date_window, text="Select Start Date:", font=("Helvetica", 15, "normal"), background="skyblue").place(x=250,y=20)
    start_date_entry = DateEntry(date_window, width=12, background='darkblue', foreground='white', borderwidth=2, font=("Helvetica", 15, "normal"), maxdate=today_date)
    start_date_entry.place(x=450,y=20)
    start_date=start_date_entry.get_date()
    tk.Label(date_window, text="Select End Date:", font=("Helvetica", 15, "normal"), background="skyblue").place(x=250,y=100)
    end_date_entry = DateEntry(date_window, width=12, background='darkblue',
                               foreground='white', borderwidth=2,
                               font=("Helvetica", 15, "normal"),
                               maxdate=today_date)
    end_date_entry.place(x=450,y=100)
    end_date=end_date_entry.get_date()
    tk.Label(date_window, text="Enter Course Name:",font=("Helvetica", 15, "normal"), background="skyblue").place(x=250,y=200)
    course_combobox = ttk.Combobox(date_window, font=("Helvetica", 15, "normal"), width=12, values=[
    "Python Full Stack",
    "Python",
    "python with Django",
    "selenium with Python",
    "SQL",
    "Python with SQL",
    "Java",
    "Java Full Stack",
    "Java with Spring",
    "Selenium with Java",
    "Data Science",
    "Data Analytics",
    "ServiceNow"
    "DevOps",
    "AWS/Azure/GCP",
    "Azure DevOps",
    "AWS DevOps",
    "PowerBI",
    "ETL Testing",
    "Blockchain",
    "Artificial Intelligence",
    "Data Science with AI",
    "Generative AI",
    "Prompt Engineering",
    "Go Programming",
    "Data Structures and Algorithms",
    "FrontEnd Development",
    "MERN Stack",
    "MEAN Stack",
    ".Net Full Stack",
    "Medical Coding",
    "AR Calling",
    "Medical Billing",
    "Tableau",
    "WordPress",
    "Digital Marketing", 
    ])
    course_combobox.place(x=450,y=200)

    try:
        if end_date>= start_date:
            end_date_entry.config(state="normal")
        else:
            end_date_entry.config(state="disabled")
    except:
        end_date_entry.config(state="disabled")

    def confirm_date(start_date, end_date):
        conn = sqlite3.connect(r'C:\Registration Form Code\user_full_data.db')
        df = pd.read_sql_query("SELECT * FROM user_data", conn)

        filtered_df = df[(df['date'] >= start_date) & (df['date'] <= end_date)]

        
        if not filtered_df.empty:
            filtered_df.to_excel(r'C:\Registration Form Code\user_data.xlsx', index=False)
            messagebox.showinfo("Success", "Data exported to Excel and saved as 'user_data.xlsx'.")
        else:
            messagebox.showwarning("No Data", "No data found in the database for the selected date range.")

        conn.close()
        date_window.destroy()

    def export_today():
        start_date = end_date = datetime.now().strftime("%Y-%m-%d")
        confirm_date(start_date, end_date)

    def export_last_week():
        end_date = datetime.now().strftime("%Y-%m-%d")
        start_date = (datetime.now() - timedelta(weeks=1)).strftime("%Y-%m-%d")
        confirm_date(start_date, end_date)

    def export_last_month():
        now = datetime.now()
        first_day_of_current_month = datetime(now.year, now.month, 1)
        last_day_of_previous_month = first_day_of_current_month - timedelta(days=1)
        first_day_of_previous_month = datetime(last_day_of_previous_month.year, last_day_of_previous_month.month, 1)
        start_date = first_day_of_previous_month.strftime("%Y-%m-%d")
        end_date = last_day_of_previous_month.strftime("%Y-%m-%d")
        confirm_date(start_date, end_date)
    def course():
        start_date = start_date_entry.get_date().strftime("%Y-%m-%d")
        end_date = end_date_entry.get_date().strftime("%Y-%m-%d")
        database_path = r'C:\Registration Form Code\user_full_data.db'
        with sqlite3.connect(database_path) as conn:
            df = pd.read_sql_query("SELECT * FROM user_data", conn)
            filtered_df = df[(df['date'] >= start_date) & (df['date'] <= end_date)]
            if not filtered_df.empty:
                course_filtered_df = filtered_df[filtered_df['course'].str.lower() == course_combobox.get().lower()]
                if not course_filtered_df.empty:
                    excel_path = r'C:\Registration Form Code\user_data.xlsx'
                    course_filtered_df.to_excel(excel_path, index=False)
                    messagebox.showinfo("Success", f"Data exported to Excel and saved as 'user_data.xlsx'.")
                else:
                    messagebox.showwarning("No Data", "No data found for the selected course in the given date range.")
            else:
                messagebox.showwarning("No Data", "No data found in the database for the selected date range.")
        date_window.destroy()



    tk.Button(date_window, text="Last Month", background="yellow", command=export_last_month, font=("Helvetica", 15, "normal")).place(x=100,y=300)
    tk.Button(date_window, text="Last Week", background="orange", command=export_last_week, font=("Helvetica", 15, "normal")).place(x=250,y=300)
    tk.Button(date_window, text="Today", background="red", command=export_today, font=("Helvetica", 15, "normal")).place(x=400,y=300)
    tk.Button(date_window, text="Export", background="green", command=lambda:confirm_date(start_date_entry.get_date().strftime("%Y-%m-%d"), end_date_entry.get_date().strftime("%Y-%m-%d")), font=("Helvetica", 15, "normal")).place(x=550,y=300)
    tk.Button(date_window, text="Get Data With Course", background="lightgreen", command=course, font=("Helvetica", 15, "normal")).place(x=700,y=300) 

def reset_form():
    entry_site_name.delete(0, tk.END)
    entry_email_id.delete(0, tk.END)
    entry_mobile_no.delete(0, tk.END)
    course_combobox.set('')
    gender.set(None)
    role.set(None)

def export_to_pdf():
    date_window = tk.Toplevel(root)
    date_window.title("Export PDF")
    window_width = 1000
    window_height = 750
    date_window['bg']='skyblue'

    logo_image = tk.PhotoImage(file='C:\\syam\\projects\\VCS1.png')

    date_window.iconphoto(False, logo_image)
    screen_width = date_window.winfo_screenwidth()
    screen_height = date_window.winfo_screenheight()
    position_x = int((screen_width / 2) - (window_width / 2))
    position_y = int((screen_height / 2) - (window_height / 2))
    date_window.geometry(f"{window_width}x{window_height}+{position_x}+{position_y}")

    today_date = datetime.now().date()

    tk.Label(date_window, text="Select Start Date:", background="skyblue", font=("Helvetica", 15, "normal")).place(x=250,y=20)
    start_date_entry = DateEntry(date_window, width=12, background='darkblue', foreground='white', borderwidth=2, font=("Helvetica", 15, "normal"), maxdate=today_date)
    start_date_entry.place(x=450,y=20)

    tk.Label(date_window, text="Select End Date:", background="skyblue", font=("Helvetica", 15, "normal")).place(x=250,y=100)
    end_date_entry = DateEntry(date_window, width=12, background='darkblue', foreground='white', borderwidth=2, font=("Helvetica", 15, "normal"), maxdate=today_date)
    end_date_entry.place(x=450,y=100)
    tk.Label(date_window, text="Enter Course Name:",font=("Helvetica", 15, "normal"), background="skyblue").place(x=250,y=200)
    course_combobox = ttk.Combobox(date_window, font=("Helvetica", 15, "normal"), width=12, values=[
    "Python Full Stack",
    "Python",
    "python with Django",
    "selenium with Python",
    "SQL",
    "Python with SQL",
    "Java",
    "Java Full Stack",
    "Java with Spring",
    "Selenium with Java",
    "Data Science",
    "Data Analytics",
    "ServiceNow"
    "DevOps",
    "AWS/Azure/GCP",
    "Azure DevOps",
    "AWS DevOps",
    "PowerBI",
    "ETL Testing",
    "Blockchain",
    "Artificial Intelligence",
    "Data Science with AI",
    "Generative AI",
    "Prompt Engineering",
    "Go Programming",
    "Data Structures and Algorithms",
    "FrontEnd Development",
    "MERN Stack",
    "MEAN Stack",
    ".Net Full Stack",
    "Medical Coding",
    "AR Calling",
    "Medical Billing",
    "Tableau",
    "WordPress",
    "Digital Marketing", 
    ])
    course_combobox.place(x=450,y=200)


    def confirm_date(start_date, end_date):
        conn = sqlite3.connect(r'C:\Registration Form Code\user_full_data.db')
        df = pd.read_sql_query("SELECT * FROM user_data", conn)

        filtered_df = df[(df['date'] >= start_date) & (df['date'] <= end_date)]

        if not filtered_df.empty:
            pdf_path = r'C:\Registration Form Code\user_data.pdf'
            doc = SimpleDocTemplate(pdf_path, pagesize=letter)
            elements = []

            data = [filtered_df.columns.tolist()] + filtered_df.values.tolist()
            table = Table(data)
            table.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (-1, 0), 14),
                ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
                ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ]))

            elements.append(table)
            doc.build(elements)

            messagebox.showinfo("Success", f"Data exported to PDF and saved as 'user_data.pdf'.")
        else:
            messagebox.showwarning("No Data", "No data found in the database for the selected date range.")

        conn.close()
        date_window.destroy()

    def export_today():
        start_date = end_date = datetime.now().strftime("%Y-%m-%d")
        confirm_date(start_date, end_date)

    def export_last_week():
        end_date = datetime.now().strftime("%Y-%m-%d")
        start_date = (datetime.now() - timedelta(weeks=1)).strftime("%Y-%m-%d")
        confirm_date(start_date, end_date)

    def export_last_month():
        now = datetime.now()
        first_day_of_current_month = datetime(now.year, now.month, 1)
        last_day_of_previous_month = first_day_of_current_month - timedelta(days=1)
        first_day_of_previous_month = datetime(last_day_of_previous_month.year, last_day_of_previous_month.month, 1)
        start_date = first_day_of_previous_month.strftime("%Y-%m-%d")
        end_date = last_day_of_previous_month.strftime("%Y-%m-%d")
        confirm_date(start_date, end_date)
    def course():
        course_name = course_combobox.get()
        conn = sqlite3.connect(r'C:\Registration Form Code\user_full_data.db')
        df = pd.read_sql_query("SELECT * FROM user_data WHERE course = ?", conn, params=(course_name,))
        conn.close()

        if df.empty:
            messagebox.showwarning("No Data", "No data found for the selected course.")
            date_window.destroy()
        else:
            pdf_filename = r'C:\Registration Form Code\user_data.pdf'
            pdf = SimpleDocTemplate(pdf_filename, pagesize=letter)

            table_data = [["Student Name", "Email ID", "Mobile No", "Course", "Gender", "Reference", "Date"]] + df.values.tolist()
            table = Table(table_data)
            table.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
                ('GRID', (0, 0), (-1, -1), 1, colors.black)
            ]))

            pdf.build([table])
            messagebox.showinfo("Success", f"Data exported to PDF and saved as 'user_data.pdf'.")

        date_window.destroy()


    tk.Button(date_window, text="Last Month", background="yellow", command=export_last_month, font=("Helvetica", 15, "normal")).place(x=100,y=300)
    tk.Button(date_window, text="Last Week", background="orange", command=export_last_week, font=("Helvetica", 15, "normal")).place(x=250,y=300)
    tk.Button(date_window, text="Today", background="red", command=export_today, font=("Helvetica", 15, "normal")).place(x=400,y=300)
    tk.Button(date_window, text="Export", background="green", command=lambda:confirm_date(start_date_entry.get_date().strftime("%Y-%m-%d"), end_date_entry.get_date().strftime("%Y-%m-%d")), font=("Helvetica", 15, "normal")).place(x=550,y=300)
    tk.Button(date_window, text="Get Data With Course", background="lightgreen", command=course, font=("Helvetica", 15, "normal")).place(x=700,y=300) 

root = tk.Tk()
root.title("Student Registration Form")
window_width = 1000
window_height = 750
root['bg']='skyblue'

logo_image = tk.PhotoImage(file='C:\\syam\\projects\\VCS1.png')

root.iconphoto(False, logo_image)

screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()
position_x = int((screen_width / 2) - (window_width / 2))
position_y = int((screen_height / 2) - (window_height / 2))
root.geometry(f"{window_width}x{window_height}+{position_x}+{position_y}")
root['bg']='skyblue'

title = tk.Label(root, text="VCS IT SOLUTIONS PVT. LTD.", background="skyblue", font=("Helvetica", 30, "bold"), pady=20)
title.pack()
title = tk.Label(root, text="Student Registration Form", background="skyblue", font=("Helvetica", 20, "bold"), pady=5)
title.pack()

frame = tk.Frame(root, padx=20, pady=20, background="skyblue")
frame.pack()

tk.Label(frame, text="Student Name:", font=("Helvetica", 15, "normal"), background="skyblue").grid(row=0, column=0, sticky=tk.W, pady=10)
entry_student_name = tk.Entry(frame, font=("Helvetica", 15, "normal"), width=30)
entry_student_name.grid(row=0, column=1, pady=10)

tk.Label(frame, text="Email ID:", font=("Helvetica", 15, "normal"), background="skyblue").grid(row=1, column=0, sticky=tk.W, pady=10)
entry_email_id = tk.Entry(frame, font=("Helvetica", 15, "normal"), width=30)
entry_email_id.grid(row=1, column=1, pady=10)

tk.Label(frame, text="Mobile No:", font=("Helvetica", 15, "normal"), background="skyblue").grid(row=2, column=0, sticky=tk.W, pady=10)
entry_mobile_no = tk.Entry(frame, font=("Helvetica", 15, "normal"),  width=30)
entry_mobile_no.grid(row=2, column=1, pady=10)

tk.Label(frame, text="Course:", font=("Helvetica", 15, "normal"), background="skyblue").grid(row=3, column=0, sticky=tk.W, pady=10)
course_combobox = ttk.Combobox(frame, font=("Helvetica", 15, "normal"), width=28, values=[
   "Python Full Stack",
   "Python",
   "python with Django",
   "selenium with Python",
   "SQL",
   "Python with SQL",
   "Java",
   "Java Full Stack",
   "Java with Spring",
   "Selenium with Java",
   "Data Science",
   "Data Analytics",
   "ServiceNow"
   "DevOps",
   "AWS/Azure/GCP",
   "Azure DevOps",
   "AWS DevOps",
   "PowerBI",
   "ETL Testing",
   "Blockchain",
   "Artificial Intelligence",
   "Data Science with AI",
   "Generative AI",
   "Prompt Engineering",
   "Go Programming",
   "Data Structures and Algorithms",
   "FrontEnd Development",
   "MERN Stack",
   "MEAN Stack",
   ".Net Full Stack",
   "Medical Coding",
   "AR Calling",
   "Medical Billing",
   "Tableau",
   "WordPress",
   "Digital Marketing",
    
])
course_combobox.grid(row=3, column=1, pady=10)

tk.Label(frame, text="Gender:", font=("Helvetica", 15, "normal"), background="skyblue").grid(row=4, column=0, sticky=tk.W, pady=10)
gender = tk.StringVar()
tk.Radiobutton(frame, text="Male", variable=gender, value="Male", background="skyblue", font=("Helvetica", 15, "normal")).grid(row=4, column=1, sticky=tk.W, pady=10)
tk.Radiobutton(frame, text="Female", variable=gender, value="Female", background="skyblue", font=("Helvetica", 15, "normal")).grid(row=4, column=1, pady=10)

tk.Label(frame, text="Reference:", font=("Helvetica", 15, "normal"), background="skyblue").grid(row=5, column=0, sticky=tk.W, pady=10)
reference = tk.StringVar()
tk.Radiobutton(frame, text="Social Media", background="skyblue", variable=reference, value="Social Media", font=("Helvetica", 15, "normal")).grid(row=5, column=1, sticky=tk.W, pady=10)
tk.Radiobutton(frame, text="Board", background="skyblue", variable=reference, value="Board", font=("Helvetica", 15, "normal")).grid(row=5, column=2, pady=10)
tk.Radiobutton(frame, text="Person", background="skyblue", variable=reference, value="Person", font=("Helvetica", 15, "normal")).grid(row=6, column=1, sticky=tk.W, pady=10)
tk.Radiobutton(frame, text="Web/Ad", background="skyblue", variable=reference, value="Web(or)Add", font=("Helvetica", 15, "normal")).grid(row=6, column=2, pady=10)

tk.Button(root, text="Submit", command=submit_data, font=("Helvetica", 15, "normal"), background='green').place(x=100, y=600)
tk.Button(root, text="Reset", command=reset_form, font=("Helvetica", 15, "normal"), background='red').place(x=300, y=600)
tk.Button(root, text="Export to Excel", command=export_to_excel, font=("Helvetica", 15, "normal"), background='orange').place(x=500, y=600)
tk.Button(root, text="Export to PDF", command=export_to_pdf, font=("Helvetica", 15, "normal"), background='Yellow').place(x=800, y=600)


root.mainloop()
