import tkinter as tk
root=tk.Tk()
root.title('change colour')
root.configure(background='beige')
from tkinter import messagebox, scrolledtext
from datetime import datetime
import pandas as pd

class ReportCardApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Immaculate Heart of Mary Girls Higher Secondary School")

        # Create labels and entries for student info
        tk.Label(root, text="Student Name").grid(row=0, column=0,padx=2,pady=2)
        self.name_entry = tk.Entry(self.root)
        self.name_entry.grid(row=0, column=1,padx=2,pady=2)

        tk.Label(root, text="Class").grid(row=1, column=0,padx=2,pady=2)
        self.class_entry = tk.Entry(self.root)
        self.class_entry.grid(row=1, column=1,padx=2,pady=2)

        tk.Label(root, text="Roll No").grid(row=2, column=0,padx=2,pady=2)
        self.rollno_entry = tk.Entry(self.root)
        self.rollno_entry.grid(row=2, column=1,padx=2,pady=2)

        tk.Label(root, text="Date").grid(row=3, column=0,padx=2,pady=2)
        self.date_entry = tk.Entry(self.root)
        self.date_entry.grid(row=3, column=1,padx=2,pady=2)
        self.date_entry.insert(0, datetime.now().strftime("%d-%m-%y"))

        self.subject_entries = []
        subjects = ["Tamil", "English", "Maths", "Physics", "Chemistry","Computer science"]
        for i, subject in enumerate(subjects):
            tk.Label(root, text=subject).grid(row=i + 6, column=0,padx=2,pady=2)
            entry = tk.Entry(self.root)
            entry.grid(row=i + 6, column=1,padx=2,pady=2)
            self.subject_entries.append(entry)

        # Buttons
        
        tk.Button(root, text="Calculate Total",bg='light pink', command=self.calculate_total).grid(row=12, column=0,)
        tk.Button(root, text="Calculate Average",bg='light blue', command=self.calculate_average).grid(row=13, column=1)
        tk.Button(root, text="Generate Remarks",bg='lavender', command=self.generate_remarks).grid(row=14, column=1)
        tk.Button(root, text="Generate Report Card",bg='light green', command=self.generate_report_card).grid(row=12, column=1)
        tk.Button(root, text="Export to Excel",bg='coral', command=self.export_to_excel).grid(row=13, column=0)
        self.report_area = tk.Entry(self.root)
        self.report_area = scrolledtext.ScrolledText(root, width=40, height=10)
        self.report_area.grid(row=14, column=0, columnspan=2,padx=2,pady=2)

        self.total = 0
        self.average = 0
        self.remarks = ""
        self.passed_subjects = []
        self.failed_subjects = []
        if self.average is not None:
            print(f"Average: {self.average}")
        else:
            print("Average not set!")
            

    def calculate_total(self):
        try:
            marks = [int(entry.get()) for entry in self.subject_entries]
            self.total = sum(marks)
            messagebox.showinfo("Total Marks", f"Total Marks: {self.total}")
        except ValueError:
            messagebox.showerror("Input Error", "Please enter valid marks.")

    def calculate_average(self):
        try:
            marks = [int(entry.get()) for entry in self.subject_entries]
            self.average = sum(marks) / len(marks)
            messagebox.showinfo("Average Marks", f"Average Marks: {self.average:.2f}")
        except ValueError:
            messagebox.showerror("Input Error", "Please enter valid marks.")

    def generate_remarks(self):
        marks = [int(entry.get()) for entry in self.subject_entries]
        overall_remark = "Passed" if self.average >= 50 else "Failed"
        
        self.passed_subjects = []
        self.failed_subjects = []
        
        individual_remarks = []
        for i, mark in enumerate(marks):
            subject_remark = "passed" if mark >= 50 else "failed"
            individual_remarks.append(f"Subject {i + 1}: {subject_remark}")
            if subject_remark == "passed":
                self.passed_subjects.append(f"Subject {i + 1}")
            else:
                self.failed_subjects.append(f"Subject {i + 1}")

        self.remarks = "\n".join(individual_remarks) + f"\nOverall: {overall_remark}"
        messagebox.showinfo("Remarks", self.remarks)

    def generate_report_card(self):
        try:
            name = self.name_entry.get()
            student_class = self.class_entry.get()
            rollno = self.rollno_entry.get()
            date = self.date_entry.get()
            marks = [int(entry.get()) for entry in self.subject_entries]
            total = self.total if self.total else sum(marks)
            average = self.average if self.average else total / len(marks)
            min_marks = min(marks)
            max_marks = max(marks)

            # Clear the report area
            self.report_area.delete(1.0, tk.END)

            # Generate report card in tabular format
            report = f"{'Report Card':^40}\n\n"
            report += f"{'Date:':<20} {date}\n"
            report += f"{'Name:':<20} {name}\n"
            report += f"{'Class:':<20} {student_class}\n"
            report += f"{'Roll No:':<20} {rollno}\n"
            report += f"{'Subject':<15} {'Marks':<10}\n"
            report += '-' * 30 + '\n'
            for i, subject in enumerate(["Tamil", "English", "Maths", "Physics", "Chemistry","Computer science"]):
                report += f"{subject:<15} {marks[i]:<10}\n"
            report += '-' * 30 + '\n'
            report += f"{'Total Marks:':<15} {total}\n"
            report += f"{'Average:':<15} {average:.2f}\n"
            report += f"{'Min Marks:':<15} {min_marks}\n"
            report += f"{'Max Marks:':<15} {max_marks}\n"
            report += f"{'Remarks:':<15} {self.remarks}\n"

            # Display report card
            self.report_area.insert(tk.END, report)

        except ValueError:
            messagebox.showerror("Input Error", "Please enter valid marks.")

    def export_to_excel(self):
        try:
            name = self.name_entry.get()
            student_class = self.class_entry.get()
            rollno = self.rollno_entry.get()
            date = self.date_entry.get()
            marks = [int(entry.get()) for entry in self.subject_entries]
            
            # Prepare data for DataFrame
            data = {
                "Student Name": [name],
                "Class": [student_class],
                "Roll No": [rollno],
                "Date": [date],
                "Total Marks": [sum(marks)],
                "Average Marks": [self.average],
                "Min Marks": [min(marks)],
                "Max Marks": [max(marks)],
                "Passed Subjects": [', '.join(self.passed_subjects)],
                "Failed Subjects": [', '.join(self.failed_subjects)],
                "Overall Remark": ["Passed" if self.average >= 50 else "Failed"],
            }

            df = pd.DataFrame(data)
            df.to_excel("report_card.xlsx", index=False)
            messagebox.showinfo("Export Success", "Report card exported to report_card.xlsx")
        except Exception as e:
            messagebox.showerror("Export Error", str(e))

if __name__ == "__main__":
    app = ReportCardApp(root)
    root.mainloop()

