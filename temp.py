import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter.ttk import Button, Label, Entry, Frame
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

# ---------------------------------------------------- Backend Functions ----------------------------------------------------

def categorize_students_attendance(dataset_path, threshold=75):
    try:
        df = pd.read_csv(dataset_path)
    except FileNotFoundError:
        messagebox.showerror("Error", "File not found. Please provide a valid file path.")
        return None, None, None, None, None
    except pd.errors.EmptyDataError:
        messagebox.showerror("Error", "The file is empty.")
        return None, None, None, None, None

    if 'Attendance Percentage' not in df.columns or 'Student Mail ID' not in df.columns:
        messagebox.showerror("Error", "The dataset is missing required columns.")
        return None, None, None, None, None

    defaulter_count = (df['Attendance Percentage'] < threshold).sum()
    non_defaulter_count = (df['Attendance Percentage'] >= threshold).sum()

    defaulter_students = df[df['Attendance Percentage'] < threshold]
    non_defaulter_students = df[df['Attendance Percentage'] >= threshold]

    return defaulter_count, non_defaulter_count, defaulter_students, non_defaulter_students, df


def generate_graphs_and_save_pdf(dataset_path, class_name="Class", class_teacher="Teacher"):
    try:
        defaulter_count, non_defaulter_count, defaulter_students, _, df = categorize_students_attendance(dataset_path)
        if df is None:
            return
        
        fig, axs = plt.subplots(3, 1, figsize=(10, 15))
        gender_counts = df['Gender'].value_counts()

        # Pie chart of Gender Distribution
        axs[0].pie(gender_counts, labels=gender_counts.index, autopct='%1.1f%%', startangle=140)
        axs[0].set_title('Gender Distribution')

        # Pie chart of Attendance Categories
        labels = ['Defaulter', 'Non-Defaulter']
        axs[1].pie([defaulter_count, non_defaulter_count], labels=labels, autopct='%1.1f%%', startangle=140)
        axs[1].set_title('Attendance Categories')

        # Bar chart of Defaulter vs Non-Defaulter
        axs[2].bar(labels, [defaulter_count, non_defaulter_count], color=['red', 'green'])
        axs[2].set_title('Defaulter vs Non-Defaulter Counts')

        pdf_filename = f"Attendance_Report_{class_name}.pdf"
        with PdfPages(pdf_filename) as pdf:
            pdf.savefig(fig)

        plt.close(fig)
        messagebox.showinfo("Success", f"Graphs saved to {pdf_filename}")
        return pdf_filename

    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")
        return None


def send_report(sender_email, password, receiver_email, pdf_filename):
    try:
        msg = MIMEMultipart()
        msg['From'] = sender_email
        msg['To'] = receiver_email
        msg['Subject'] = "Attendance Report"

        message = """
        <html>
            <body>
                <h1 style="color: #2C3E50;">Attendance Report</h1>
                <p>Please find attached the attendance report.</p>
            </body>
        </html>
        """
        msg.attach(MIMEText(message, 'html'))

        with open(pdf_filename, "rb") as attachment:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(attachment.read())
            encoders.encode_base64(part)
            part.add_header("Content-Disposition", f"attachment; filename={pdf_filename}")
            msg.attach(part)

        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(sender_email, password)
        server.sendmail(sender_email, receiver_email, msg.as_string())
        server.quit()

        messagebox.showinfo("Success", "Email sent successfully!")
    except Exception as e:
        messagebox.showerror("Error", f"Error sending email: {e}")

# ---------------------------------------------------- GUI Application ----------------------------------------------------

def select_file():
    file_path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
    entry_file_path.delete(0, tk.END)
    entry_file_path.insert(0, file_path)


def generate_and_send():
    dataset_path = entry_file_path.get()
    sender_email = entry_sender_email.get()
    password = entry_password.get()
    receiver_email = entry_receiver_email.get()

    if not all([dataset_path, sender_email, password, receiver_email]):
        messagebox.showerror("Error", "Please fill in all fields.")
        return

    pdf_filename = generate_graphs_and_save_pdf(dataset_path)
    if pdf_filename:
        send_report(sender_email, password, receiver_email, pdf_filename)


# Initialize GUI
root = tk.Tk()
root.title("Attendance Report Generator")
root.geometry("500x400")

frame = Frame(root)
frame.pack(pady=10, padx=10)

# File selection
Label(frame, text="Select Attendance CSV File:").grid(row=0, column=0, pady=5, sticky="e")
entry_file_path = Entry(frame, width=30)
entry_file_path.grid(row=0, column=1, padx=5)
Button(frame, text="Browse", command=select_file).grid(row=0, column=2)

# Email fields
Label(frame, text="Sender Email:").grid(row=1, column=0, pady=5, sticky="e")
entry_sender_email = Entry(frame, width=30)
entry_sender_email.grid(row=1, column=1, padx=5)

Label(frame, text="App Password:").grid(row=2, column=0, pady=5, sticky="e")
entry_password = Entry(frame, width=30, show="*")
entry_password.grid(row=2, column=1, padx=5)

Label(frame, text="Receiver Email:").grid(row=3, column=0, pady=5, sticky="e")
entry_receiver_email = Entry(frame, width=30)
entry_receiver_email.grid(row=3, column=1, padx=5)

# Generate and Send Button
Button(frame, text="Generate and Send Report", command=generate_and_send).grid(row=4, columnspan=3, pady=20)

# Run the GUI
root.mainloop()
