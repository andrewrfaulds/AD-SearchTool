from tkinter import *
from tkinter import messagebox
import re 

#main window paramters
root = Tk()
root.title("Email Address Extractor")
root.geometry('1500x900+200+200')


# function to search input string for regex matches, store in list and then 
# create unique set of matches in new list. It then inputs these to a textbox on the gui form.

def get_emails(s):
    emails = [re.findall("([a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+)", s)]
    for email in emails:
        newset = set(email)
        txtbox2.insert(1.0, newset)
    
    # event handler for the process button. clears output textbox on each run 
    # and takes input from input texbox and passes it to the get_emails function.
def buttonClick():
    txtbox2.delete(1.0, END)
    line = txtbox.get("1.0", END)
    get_emails(line)
 	 
# function to allow exporting to txt file with try/except checks. it will append or 
# create file foo.txt in the current working directory and write each value in the output
# textbox to the file, opening a message box on completion or generating exception on fail.
def export_txt():
 try:
    outputdata = [txtbox2.get("1.0", 'end -1c')]
    with open("foo.txt", "w+") as fo:
      for val in outputdata:
        fo.write(val)
    messagebox.showinfo("Alert", "Export Complete")
 except:
    print('Export Failed!')




# basic form paramters 
app = Frame(root)
app.grid(padx=40, pady=40)
lbl1 = Label(app, text="Enter text here: ")
lbl1.grid(row = 7, column = 2)
lbl2 = Label(app, text="Addresses Detected: ")
lbl2.grid(row = 7, column = 5)
bttn1 = Button(app, command=lambda:buttonClick(), text = "Process Text")
bttn1.grid(row = 9, column = 2)
bttn2 = Button(app, command=lambda:export_txt(), text = "Export")
bttn2.grid(row = 9, column = 5)
txtbox = Text(app, width = 50, height = 25, bg = 'white', wrap = WORD)
txtbox.grid(row = 6, column = 2)
txtbox2 = Text(app, width = 50, height = 25,bg = 'white', wrap = WORD)
txtbox2.grid(row = 6, column = 5)

# initiate main loop
mainloop()
 