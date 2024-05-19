import ttkbootstrap as ttk
import openpyxl
import os
from ttkbootstrap.dialogs import Messagebox
import re, time, sys

def create_excel_file():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(['Name', 'Contact', 'Gender', 'Age', 'Nationality', 'Obtained Marks', 'Occupation','Email Address', 'Password', 'Salary'])
    wb.save('profile_registration.xlsx')

root = ttk.Window(themename='darkly')
root.geometry("500x800")
root.resizable(0, 0)
root.title("Profile Registration")

def reg():
    nadata = nt1.get()
    cdata = ct.get()
    gdata = r.get()
    adata = asb.get()
    ndata = n.get()
    mdata = m.get()
    odata = o.get()
    edata = ent.get()
    pdata = passnt.get()
    sdata = sbox.get()
    
    if not os.path.isfile('job_registration.xlsx'):
        create_excel_file()
    
    wb = openpyxl.load_workbook('job_registration.xlsx')
    ws = wb.active
    
    ws.append([nadata, cdata, gdata, adata, ndata, mdata, odata, edata, pdata, sdata + "$"])
    wb.save('job_registration.xlsx')
    
    Messagebox.show_info("Registration Successful!", title='Success')
    time.sleep(0.5)
    sys.exit()

def validcheck(input):
    return bool(re.match(r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b', input))

def check(*args):
    flag = False
    i = ent.get()
    g = validcheck(i)
    if len(nt1.get()) < 1:
        flag = True
    if len(ct.get()) < 7 or len(ct.get()) > 15:
        flag = True
    if len(passnt.get()) < 8:
        flag = True
    if c.get() != "Y":
        flag = True
    if not g:
        flag = True
    if not flag:
        rbtn.config(state="normal", command=reg)
    else:
        rbtn.config(state="disabled", command=check)

# Name
nameframe = ttk.Frame(root, width=525)
nameframe.pack(pady=(30, 0))
namelabel = ttk.Label(nameframe, text="Name")
namelabel.pack(side=ttk.LEFT, padx=(10, 65))
na = ttk.StringVar()
nt1 = ttk.Entry(nameframe, bootstyle='info', textvariable=na)
nt1.pack(side=ttk.RIGHT, padx=(15, 30))
na.trace('w', check)

# Contact
conframe = ttk.Frame(root, width=525)
conframe.pack(pady=(15, 0))
conlbl = ttk.Label(conframe, text="Contact")
conlbl.pack(side=ttk.LEFT, padx=(20, 65))

def only_numbers(char):
    return char.isdigit()

validation = root.register(only_numbers)
co = ttk.StringVar()
ct = ttk.Entry(conframe, bootstyle='info', textvariable=co, validate='key', validatecommand=(validation, '%S'))
ct.pack(side=ttk.RIGHT, padx=(5, 40))
co.trace('w', check)

# Gender
gframe = ttk.Frame(root)
gframe.pack(pady=(15, 0))
r = ttk.StringVar()
glbl = ttk.Label(gframe, text="Gender")
glbl.pack(side=ttk.LEFT, padx=(20, 60))
rm = ttk.Radiobutton(gframe, text="Male", value="Male", variable=r, bootstyle='info')
rm.pack(side=ttk.LEFT, padx=(30, 20))
rf = ttk.Radiobutton(gframe, text="Female", value="Female", variable=r, bootstyle='info')
rf.pack(side=ttk.RIGHT, padx=(20, 45))

# Age
aframe = ttk.Frame(root)
aframe.pack(pady=(15, 0))
albl = ttk.Label(aframe, text="Age")
albl.pack(side=ttk.LEFT, padx=(10, 100))
var = ttk.IntVar(aframe)
var.set(18)
asb = ttk.Spinbox(aframe, from_=18, to=35, textvariable=var, bootstyle='info')
asb.pack(side=ttk.RIGHT)

# Nationality
nframe = ttk.Frame(root)
nframe.pack(pady=(15, 0))
nlbl = ttk.Label(nframe, text="Nationality")
nlbl.pack(side=ttk.LEFT)
n = ttk.StringVar()
nbox = ttk.Combobox(nframe, bootstyle='info', textvariable=n, values=["Bangladesh", "India", "Nepal", "Sri Lanka", "UK", "USA"])
nbox.pack(side=ttk.RIGHT, padx=(45, 0))

# Marks
mframe = ttk.Frame(root)
mframe.pack(pady=(15, 0))
m = ttk.StringVar()
mlbl = ttk.Label(mframe, text="Obtained Marks")
mlbl.pack(side=ttk.LEFT)
cb1 = ttk.Combobox(mframe, bootstyle='info', textvariable=m, values=["60%-64%", "65%-69%", "70%-74%", "75%-79%", "80%-84%", "85%-89%", "90%-94%", "95%-100%"])
cb1.pack(side=ttk.RIGHT, padx=(5, 0))

# Occupation
oframe = ttk.Frame(root)
oframe.pack(pady=(15, 0))
o = ttk.StringVar()
olbl = ttk.Label(oframe, text="Occupation")
olbl.pack(side=ttk.LEFT, padx=(0, 40))
opb = ttk.Combobox(oframe, bootstyle='info', textvariable=o, values=[
    'Programmer', 'Software Developer', 'Veterinarian', 'Nurse', 'Lawyer', 'Physician', 'Dentist', 'Engineer', 'Surgeon',
    'Physician assistant', 'Web Developer', 'Electrician', 'Information Security Analysts', 'Physical Therapist',
    'Pharmacist', 'Financial Manager', 'Marketing management', 'Financial Analyst', 'Pilot', 'Engineering',
    'Information technology management', 'Sales', 'Student', 'Freelancer', 'Graphics Designing', 'Logistician',
    'Management Analyst', 'Data Scientist', 'Statistician', 'Officer (Army/Navy/Air Force/Police)', 'Scientist',
    'School/College/University Teacher', 'Computer Network Architect', 'Banker'
])
opb.pack(side=ttk.RIGHT)

# Email
eframe = ttk.Frame(root)
eframe.pack(pady=(15, 0))
e = ttk.StringVar()
elbl = ttk.Label(eframe, text="Email Address")
elbl.pack(side=ttk.LEFT, padx=(15, 0))
ent = ttk.Entry(eframe, bootstyle='info', textvariable=e)
ent.pack(side=ttk.RIGHT, padx=(20, 35))
e.trace('w', check)

# Password
passframe = ttk.Frame(root)
passframe.pack(pady=(15, 0))
passlbl = ttk.Label(passframe, text="Password")
passlbl.pack(side=ttk.LEFT, padx=(15, 35))
p = ttk.StringVar()
passnt = ttk.Entry(passframe, bootstyle='info', show="*", textvariable=p)
passnt.pack(side=ttk.RIGHT, padx=(20, 35))
p.trace('w', check)

# Salary
sframe = ttk.Frame(root)
sframe.pack(pady=(15, 0))
slbl = ttk.Label(sframe, text="Salary ($)/5000")
slbl.pack(side=ttk.LEFT, padx=(40, 5))
s = ttk.IntVar()
s.set(500)
sbox = ttk.Spinbox(sframe, bootstyle='info', from_=500, to=5000, textvariable=s)
sbox.pack(side=ttk.RIGHT, padx=(15, 35))

# Checkbox for agreement
cframe = ttk.Frame(root)
cframe.pack(pady=(25, 0))
c = ttk.StringVar(root)
c.set("")
chk = ttk.Checkbutton(cframe, text="I Agree", variable=c, bootstyle='success-round-toggle', onvalue="Y", offvalue="", command=check)
chk.pack(side=ttk.LEFT, padx=(0, 275))

# Sign Up button
bframe = ttk.Frame(root)
bframe.pack(pady=(20, 0))
rbtn = ttk.Button(bframe, text="Sign Up", bootstyle='success-outline', state="disabled", command=check)
rbtn.pack()

root.mainloop()