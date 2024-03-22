import tkinter as tk
from tkinter import filedialog
import openpyxl
import datetime
import os

file_path = ''
def import_file():
    file_path = filedialog.askopenfilename(title="Select Excel file", filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")])
    if file_path:
        try:
            wb = openpyxl.load_workbook(f"{file_path}", data_only=True, keep_vba=True)
            sheetk = wb.active
            created_file_name = f'pay_slip_{datetime.date.today()}-{datetime.datetime.now().strftime("%f")}.html'
            token_file = open(created_file_name,'a')
            token_file.write(
                '''<!DOCTYPE html>
            <html>
            <style>
            table{
              border:1px solid black;
              font-size: 9px;
              display: inline-block;
              margin: 5px;
            }
            p{
              text-align: center;
            }
            </style>
            <body>'''
            )
            for v in range(2,34):
                token_file.write(
                    f'''<table style="width:160px;"><tr><td colspan='2'><p>COMPANY NAME</br>Pay Slip</p></td></tr><tr><td>Month:</td><td>MONTH NAME</td></tr><tr><td>Card No.:</td><td>{sheetk[f"AP{v}"].value}</td></tr><tr><td>Name:</td><td>{sheetk[f"B{v}"].value}</td></tr><tr><td>Designation:</td><td>{sheetk[f"D{v}"].value}</td></tr><tr><td>Section:</td><td>{sheetk[f"F{v}"].value}</td></tr><tr><td>Grade:</td><td>{sheetk[f"C{v}"].value}</td></tr><tr><td colspan='2'><hr></td></tr><tr><td>Absent Days:</td><td>{sheetk[f"S{v}"].value}</td></tr><tr><td>H.A.:</td><td>{int(sheetk[f"H{v}"].value)}</td></tr><tr><td>T.A.:</td><td>450</td></tr><tr><td>F.A.:</td><td>1250</td></tr><tr><td>M.A.:</td><td>750</td></tr><tr><td>Basic Salary:</td><td>{int(sheetk[f"G{v}"].value)}</td></tr><tr><td>Gross Salary:</td><td>{sheetk[f"M{v}"].value}</td></tr><tr><td>Attendance Bonus:</td><td>{sheetk[f"W{v}"].value}</td></tr><tr><td colspan='2'><b>Deductions</b></td></tr><tr><td>Absent Ded:</td><td>{int(sheetk[f"U{v}"].value)}</td></tr><tr><td>ID Card Ded.:</td><td>{sheetk[f"AJ{v}"].value}</td></tr><tr><td>Stamp Ded.:</td><td>{sheetk[f"AI{v}"].value}</td></tr><tr><td>Advanced Ded.:</td><td>{sheetk[f"AM{v}"].value}</td></tr><tr><td colspan='2'><b>Over Time Calculate</b></td></tr><tr><td>OT Hr:</td><td>{sheetk[f"AA{v}"].value}</td></tr><tr><td>OT Rate:</td><td>{round(sheetk[f"AB{v}"].value, 2)}</td></tr><tr><td>OT Taka:</td><td>{round(sheetk[f"AC{v}"].value, 2)}</td></tr><tr><td colspan='2'><hr></td></tr><tr><td><b>Payable:</b></td><td><b>{int(sheetk[f"AO{v}"].value)}</b></td></tr><tr><td>.</td><td><hr></td></tr><tr><td></td><td><img width="60px" src="c:\do_not_delete_this_image.png" alt="Mainul"></td></tr><tr><tr><td>Receiver Sign.</td><td>Authorize Sign.</td></tr></table>'''
                    )
                

            token_file.write(
            '''</body>
            </html>'''
            )
            label = tk.Label(root, text='Success!! Check your downloaded file in the same path this application.',bg='#fff', fg='#008000')
            label.pack()
        except SyntaxWarning:
            print('There is a syntax warning. Please avoid this message')
        except:
            label = tk.Label(root, text='Your File is password protected or Something else error. Please turn off the password and try again.',bg='#fff', fg='#f00')
            label.pack()
            token_file.close()
            os.remove(created_file_name)

root = tk.Tk()
root.title("Import Excel File")

# Create an "Import File" button
import_button = tk.Button(root, text="Import File", command=import_file)
import_button.pack(pady=100)
# run the Tkinter event loop
root.mainloop()




