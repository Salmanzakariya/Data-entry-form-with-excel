import tkinter
from tkinter import ttk
from tkinter import messagebox
import tkinter.messagebox
import os
import openpyxl
import openpyxl.workbook



def enter_data():
    acc=accept.get()
    if acc=="Accepted":
     gender=gender_box.get()
     one=fentry.get()
     
     three=lentry.get()

     if one and three:
         two=mentry.get()
         age=agespinbox.get()
         nat=nationcombbox.get()
         stat=statecombbox.get()
         ed=study_status.get()
         co=cousecheckbox.get()
         se=semsterspinbox.get()
         bs=bspinbox.get()
         print("prnoun:",gender,  "first name :",one,   "Middle name: ",two,  "Last name :",three)
         print("-----------------------------------------------")
         print("age :",age,"nation:",nat,"state:",stat)
         print("-----------------------------------------------")
         print("status",ed)
         print("-----------------------------------------------")
         print("bsc/btech :",bs,"couse:",co,"semster:",se)
         print("-----------------------------------------------")
#-----------------------------------the part connectig the info into excel sheet-----------------------
        
        
         filepath="C:\mywork\python\projects\Data.xlsx"

         if not os.path.exists(filepath):
            workbook =openpyxl.Workbook()
            sheet =workbook.active
            heading=["Pronun","First name","Middle name","Last name","age","Nationality","State","Bsc/btch","Cousre","Semster"]
            sheet.append(heading)
            workbook.save(filepath)
         workbook=openpyxl.load_workbook(filepath)
         sheet=workbook.active
         sheet.append([gender,one,two,three,age,nat,ed,bs,co,se])
         workbook.save(filepath)









     else:
         tkinter.messagebox.showwarning(title="ERROR",message="first or last name is needed")
    else:
       tkinter.messagebox.showwarning(title="ERROR",message="Not verifed")





window=tkinter.Tk()
window.title("Data Entry Form")

frame=tkinter.Frame(window)
frame.pack()

#diving the frame into 4 including button

#1st user information


user_info=tkinter.LabelFrame(frame,text="User information")
user_info.grid(row=0,column=0,padx=20,pady=10)



#pronoun's
title=tkinter.Label(user_info,text="Tile")
gender_box=ttk.Combobox(user_info,values=["MR","MRS","MS","OTHER"])
title.grid(row=0,column=0)
gender_box.grid(row=1,column=0)

#first name


fname=tkinter.Label(user_info,text="First Name")
fname.grid(row=0,column=1)
fentry=tkinter.Entry(user_info)
fentry.grid(row=1,column=1)


# #middle name
mname=tkinter.Label(user_info,text="Middle Name")
mname.grid(row=0,column=2)
mentry=tkinter.Entry(user_info)
mentry.grid(row=1,column=2)

#last name
lname=tkinter.Label(user_info,text="Last Name")
lname.grid(row=0,column=3)
lentry=tkinter.Entry(user_info)
lentry.grid(row=1,column=3)

#age
age=tkinter.Label(user_info,text="Age")
agespinbox=tkinter.Spinbox(user_info, from_=16,to=150)
age.grid(row=2,column=0)
agespinbox.grid(row=3,column=0)

#nationlity
nation=tkinter.Label(user_info,text="Nation")
nation.grid(row =2,column=1)
nationcombbox=ttk.Combobox(user_info,values=["India","Africa","Antartica","Asia","Europe","North ameria","Oceania","South america","China"])
nationcombbox.grid(row=3,column=1)
 
#state
# if nationcombbox==ttk.Combobox(user_info,values="India"):
state=tkinter.Label(user_info,text="State")
state.grid(row=2,column=2)
statecombbox=ttk.Combobox(user_info,values=["kerala","tamil nadu","up","delhi","karanatak"])
statecombbox.grid(row=3,column=2)

#widget
for widget in user_info.winfo_children():
    widget.grid_configure(padx=5,pady=4)
    
##---------------------------------------------------##

#2nd education


edu=tkinter.LabelFrame(frame)
edu.grid(row=1,column=0,sticky="news",padx=20,pady=10)

#currently studying checking
study_check=tkinter.Label(edu,text="Are you currently studying ")
study_check.grid(row=0,column=0)

study_status=tkinter.StringVar(value="Not studying")
studycheck_button=tkinter.Checkbutton(edu,text="currently studying or not?",variable=study_status,onvalue="studying",offvalue="Not studying")


studycheck_button.grid(row=1,column=0)

#btech/bsc
b=tkinter.Label(edu,text="BSC/BTECH")
bspinbox=tkinter.Spinbox(edu, values=["BSC","BTECH","OTHERS"])
b.grid(row=0,column=1)
bspinbox.grid(row=1,column=1)

# #couses
couse=tkinter.Label(edu,text="couse")
couse.grid(row=0,column=2)
cousecheckbox=ttk.Combobox(edu,values=["computer science","biology","chemistry","english","bca","mathematics","physics"])
cousecheckbox.grid(row=1,column=2)

#sem
semster=tkinter.Label(edu,text="Semster")
semsterspinbox=tkinter.Spinbox(edu, from_=1,to=8)
semster.grid(row=0,column=3)
semsterspinbox.grid(row=1,column=3)


#widget
for widget in user_info.winfo_children():
    widget.grid_configure(padx=5,pady=4)




##-----------------------------------------------##
#3nd verification

verify=tkinter.LabelFrame(frame,text="information")
verify.grid(row=2,column=0,sticky="news",padx=20,pady=10)

ver=tkinter.Label(verify,text="verifition")
ver.grid(row=0,column=0)

accept=tkinter.StringVar(value="Not accepted")
vercheck=ttk.Checkbutton(verify,text="I verified the given information",
                         variable=accept,onvalue="Accepted",offvalue="Not accepted")
vercheck.grid(row=1,column=0)

#button
button=tkinter.Button(frame,text="Button",command=enter_data)
button.grid(row=3,column=0,sticky="news",padx=20,pady=10)



window.mainloop()
