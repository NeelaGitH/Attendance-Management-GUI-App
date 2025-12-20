import os
import sys
import glob
from datetime import datetime
import pandas as pd
from tkinter import *
from tkinter import ttk, messagebox
from tkcalendar import DateEntry
import subprocess
import ttkbootstrap as tb
from tkinter import simpledialog
import numpy as np


STUDENTS_DIR = "students"
os.makedirs(STUDENTS_DIR, exist_ok=True)

DEFAULT_COLUMNS = ["class no.", "date", "day", "class timing", "comment"]


def list_student_names():
    files = glob.glob(os.path.join(STUDENTS_DIR, "*.xlsx"))
    return sorted([os.path.splitext(os.path.basename(f))[0] for f in files])

def ensure_student_file(name):
    path = os.path.join(STUDENTS_DIR, f"{name}.xlsx")
    if not os.path.exists(path):
        df = pd.DataFrame(columns = DEFAULT_COLUMNS)
        df.to_excel(path, index = False)

    return path

def get_next_class_number(df):
    if "class no." not in df.columns:
        return 1
    nums = []
    for x in df["class no."]:
        if pd.isna(x):
            continue
        try:
            nums.append(int(x))
        except:
            continue

    if len(nums) == 0:
        return 1
    
    return max(nums) + 1

def append_attendance_to_excel(name, date_str, time_str, comment = ""):
    path = ensure_student_file(name)

    try:
        df = pd.read_excel(path)

    except Exception:
        df = pd.DataFrame(columns = DEFAULT_COLUMNS)

    
    for col in DEFAULT_COLUMNS:
        if col not in df.columns:
            df[col] = None
    
    try:
        d = datetime.strptime(date_str, "%d-%m-%Y")
        day = d.strftime("%A")
        date_fmt = d.strftime("%d-%m-%Y")
    except Exception:
        day = ""
        date_fmt = date_str

    #Compute next class number
    if comment not in ["Class Not Counted", "Exam Leave", "Holiday Leave"]:
        next_class_no = get_next_class_number(df)
    else:
        next_class_no = np.nan
    
    if "date" in df.columns and "class timing" in df.columns:
        dup = df[(df["date"].astype(str) == str(date_fmt)) & (df["class timing"].astype(str) == str(time_str))]
        if not dup.empty:
            df.loc[(df["date"].astype(str) == str(date_fmt)) & (df["class timing"].astype(str) == str(time_str)), "comment"] = comment
            if comment in ["Class Not Counted", "Exam Leave", "Holiday Leave"]:
                df.loc[(df["date"].astype(str) == str(date_fmt)) & (df["class timing"].astype(str) == str(time_str)), "class no."] = np.nan


        else:
            new_row = {
                "class no.": next_class_no,
                "date":date_fmt,
                "day": day,
                "class timing": time_str,
                "comment": comment
            }
            df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)

        df.to_excel(path, index=False)


class AutocompleteCombobox(ttk.Combobox):
    def __init__(self, master = None, **kwargs):
        super().__init__(master, **kwargs)
        self._all_values = []
        self.bind('<KeyRelease>', self._on_keyrelease)

    def set_values(self, values):
        self._all_values = list(values)
        self["values"] = self._all_values

    def _on_keyrelease(self, event):
        typed = self.get().strip().lower()
        if typed == "":
            filtered = self._all_values
        else:
            filtered = [v for v in self._all_values if typed in v.lower()]
            if "Add New Student" not in filtered:
                filtered.append("Add New Student")
        self["values"] = filtered
        if filtered:
            self.event_generate('<Down>')

    
class AttendanceApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Attendance Manager")
        self.root.geometry("1000x1000")
        self._build_ui()
        self.refresh_student_list()

    def _build_ui(self):
        pad = 8
        frame = ttk.Frame(self.root, padding = pad)
        frame.pack(fill = BOTH, expand = True)


        row1 = ttk.Frame(frame)
        row1.pack(fill = X, pady = 4)
        ttk.Label(row1, text = "Student (type):").pack(side = LEFT, padx = (0, 6))

        self.combo = AutocompleteCombobox(row1, width = 40)
        self.combo.pack(side = LEFT, padx = (0, 6))
        ttk.Button(row1, text = "Refresh", command = self.refresh_student_list).pack(side = LEFT)
        ttk.Button(row1, text = "Add->", command = self.add_selected_student).pack(side = LEFT, padx = (6, 0))

        row2 = ttk.Frame(frame)
        row2.pack(fill = BOTH, pady = 8, expand = True)


        left = ttk.Frame(row2)
        left.pack(side = LEFT, fill = BOTH, expand = True, padx = (0,6))
        ttk.Label(left, text = "Students to mark:").pack(anchor = W)
        self.listbox = Listbox(left, height = 12)
        self.listbox.pack(fill = BOTH, expand = True)

        btns = ttk.Frame(row2)
        btns.pack(side = LEFT, fill = Y)
        ttk.Button(btns, text = "Remove selected", command = self.remove_selected).pack(pady = 6)
        ttk.Button(btns, text = "Clear List", command = self.clear_list).pack(pady = 6)

        right = ttk.Frame(frame)
        right.pack(fill = X, pady = (6, 0))

        #Date
        date_row = ttk.Frame(right)
        date_row.pack(fill = X, pady = 6)
        ttk.Label(date_row, text = "Date (DD-MM-YYYY):").pack(side = LEFT, padx = (0, 6))

        self.date_picker = tb.DateEntry(date_row,dateformat='%d-%m-%Y',bootstyle="danger",firstweekday = 0)
        self.date_picker.pack(side = LEFT)

        #Time
        time_row = ttk.Frame(right)
        time_row.pack(fill= X, pady = 6)
        ttk.Label(time_row, text = "Time (HH:MM)").pack(side = LEFT, padx = (0, 6))

        self.time_var = StringVar()
        self.time_var.set(datetime.now().strftime("%H:%M"))

        Button(time_row, text = "Pick Time", command = self.pick_time).pack(side = LEFT, padx = (0, 6))
        self.time_entry = Label(time_row, textvariable = self.time_var).pack(side = LEFT)


        # Comment dropdown
        remark_row = ttk.Frame(right)
        remark_row.pack(fill = X, pady = 6)
        ttk.Label(remark_row, text = "Comment:").pack(side = LEFT, padx = (0, 6))
        self.remark_combo = ttk.Combobox(
            remark_row,
            values = ["Present", "Absent", "Class Not Counted", "Exam Leave", "Holiday Leave"],
            state = "readonly",
            width = 20
        )
        self.remark_combo.current(0)
        self.remark_combo.pack(side = LEFT)

        #Save button
        mark_row = ttk.Frame(right)
        mark_row.pack(fill = X, pady = 12)
        ttk.Button(mark_row, text = "Mark for listed students",command = self.mark_for_list).pack()

        #Log
        status_row2 = ttk.Frame(frame)
        status_row2.pack(fill = BOTH, pady = (8, 0), expand = True)
        ttk.Label(status_row2, text = "Log:").pack(anchor = W)
        self.log_text = Text(status_row2, height = 7, state = DISABLED)
        self.log_text.pack(fill = BOTH, expand = True)

    def pick_time(self):
            result = subprocess.check_output(["python", "kivyTime.py"]).decode().strip()
            self.time_var.set(result)
        
    def log(self, msg):
        self.log_text.configure(state = NORMAL)
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.log_text.insert(END, f"[{now}] {msg}\n")
        self.log_text.see(END)
        self.log_text.configure(state = DISABLED)

    def refresh_student_list(self):
        names = list_student_names()

        names.append("Add New Student")

        self.combo.set_values(names)
        self.log(f"Loaded {len(names) - 1} student files.")
    
    def add_selected_student(self):
        name = self.combo.get().strip()

        if not name:
            messagebox.showwarning("No name", "Type or select a student name.")
            return 
        
        if name == "Add New Student":
            new_name = simpledialog.askstring("New Student", "Enter new student name:")
            if not new_name:
                return
            
            new_name = new_name.strip()
            if not new_name:
                return
            

            #If file exists, just add
            path = os.path.join(STUDENTS_DIR, f"{new_name}.xlsx")
            if not os.path.exists(path):
                ensure_student_file(new_name)
                self.log(f"Created file for {new_name}")

            self.refresh_student_list()

            if new_name not in self.listbox.get(0, END):
                self.listbox.insert(END, new_name)
                self.log(f"Added {new_name}")
            
            return
        

        #Normal flow for existing names
        path = os.path.join(STUDENTS_DIR, f"{name}.xlsx")
        if not os.path.exists(path):
            if messagebox.askyesno("Create file?", f"No file for '{name}'. Create?"):
                ensure_student_file(name)
                self.refresh_student_list()
            else:
                return
            
        if name in self.listbox.get(0, END):
            messagebox.showinfo("Already", f"{name} is already in the list.")
            return
        
        self.listbox.insert(END, name)
        self.log(f"Added {name}")

    
    def remove_selected(self):
        sel = self.listbox.curselection()
        for i in reversed(sel):
            name = self.listbox.get(i)
            self.listbox.delete(i)
            self.log(f"Removed {name}")

    
    def clear_list(self):
        self.listbox.delete(0, END)
        self.log("Cleared list.")

    def mark_for_list(self):
        count = self.listbox.size()
        if count == 0:
            messagebox.showinfo("List is empty", "No students to mark")
            return
        
        date_str = self.date_picker.entry.get()
        time_str = self.time_var.get()
        comment = self.remark_combo.get().strip()

        if not messagebox.askyesno("Confirm", f"Mark {count} students as '{comment}' on {date_str} {time_str}?"):
            return
        
        names = self.listbox.get(0, END)

        for name in names:
            err_student = []
            try:
                append_attendance_to_excel(name, date_str, time_str, comment)
                self.log(f"Marked {name} as {comment}")

            except Exception as e:
                err_student.append(name + ".xlsx")
                self.log(f"Error writing {name}: {e}")
        
        if err_student:
            messagebox.showerror(
            "Excel files open",
            "Please close the following Excel files:\n\n"
            + "\n".join(err_student) + "\n\nMarked attendance for rest if any. Please try again after closing the file(s)."
            )

        else:
            messagebox.showinfo("Done", "Marked attendance for all listed students.")

def get_open_excel_files(folder):
    open_files = []

    for filename in os.listdir(folder):
        if filename.lower().endswith(".xlsx"):
            path = os.path.join(folder, filename)
            try:
                with open(path, "a"):
                    pass
            except PermissionError:
                if not filename.startswith("~$"):
                    open_files.append(filename)
    
    return open_files

def main():
    open_files = get_open_excel_files(STUDENTS_DIR)

    if open_files:
        messagebox.showerror(
            "Excel files open",
            "Please close the following Excel files and restart the application:\n\n"
            + "\n".join(open_files) + "\n\nPlease don't open any excel files while the application is running."
        )

        return

    root = tb.Window(themename="darkly")
    app = AttendanceApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()

            
            








