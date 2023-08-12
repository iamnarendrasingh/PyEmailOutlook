import tkinter as tk
from tkinter import filedialog, messagebox
from send_email import read_email_details_from_excel

def browse_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
    if file_path:
        entry_file_path.delete(0, tk.END)
        entry_file_path.insert(0, file_path)

def send_emails():
    excel_filename = entry_file_path.get()
    if not excel_filename:
        messagebox.showerror("Error", "Please select an Excel file.")
        return

    try:
        read_email_details_from_excel(excel_filename)
        messagebox.showinfo("Success", "Emails sent and Excel sheet updated.")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")

# Create main application window
app = tk.Tk()
app.title("Email Sender App")

# Create and place widgets
label_instruction = tk.Label(app, text="Select the Excel file with email details:")
label_instruction.pack(pady=10)

entry_file_path = tk.Entry(app, width=40)
entry_file_path.pack()

button_browse = tk.Button(app, text="Browse", command=browse_file)
button_browse.pack(pady=5)

button_send = tk.Button(app, text="Send Emails", command=send_emails)
button_send.pack(pady=10)

# Start the GUI application
app.mainloop()
