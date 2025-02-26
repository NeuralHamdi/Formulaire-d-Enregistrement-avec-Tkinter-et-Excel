from tkinter import *
from tkinter import ttk
from tkinter import messagebox
import os
import openpyxl
# Fonction appelée lors du clic sur le bouton "Submit"
def click():
    validation=validation_var.get()
    if validation=="accepted":
        name=entry_name.get()
        last_name=entry_prenom.get()
        title=combox_title.get()
        age=age_spinbox.get()
        nationality=combox_nationality.get()
        semestre=spinbox_semestre.get()
        course=spinbox_courses.get()
        register=Var.get()
        print('-------------------------------------------------------------------')
        print('First Name :',name,end="  ")
        print("Last Name :",last_name)
        print('Title :',title,end="  ")
        print("age :",age)
        print('Nationality :',nationality)
        print('semestre :',semestre,end="   ")
        print("Course :",course)
        print("Registration Status ",register)
        print('-----------------------------------------------------------------')
        #Définition du chemin du fichier Excel où les données seront enregistrées
        path="ecrire votre chemin ici n'oublier pas d'ajouter '//'"
        if not os.path.exists(path):
            workbok=openpyxl.Workbook()
            feuile=workbok.active
            headin=["nom","prenom","age","nationality"]
            feuile.append(headin)
            workbok.save(path)
        workbook=openpyxl.load_workbook(path)
        feuiel=workbook.active
        feuiel.append([name,last_name,age,nationality])
        workbook.save(path)


    else:
        messagebox.showwarning(title="attention",message="you have to accepte the term and condition")


window=Tk()

window.title("form information")
info_frame=Frame(window)
info_frame.pack()

#cree un frame label entre le pere info_frame
info_personal=LabelFrame(info_frame,text="User information")
info_personal.grid(row=0,column=0,pady=10,padx=20)


label_name=Label(info_personal,text="nom")
label_name.grid(row=0,column=0)

entry_name=Entry(info_personal)
entry_name.grid(row=1,column=0)

label_prenom=Label(info_personal,text="prénom")
label_prenom.grid(row=0,column=1)

entry_prenom=Entry(info_personal)
entry_prenom.grid(row=1,column=1)

#cree un liste pour choisir le sexe du utlisateur avec Combobox
label_title=Label(info_personal,text="Title")
label_title.grid(row=0,column=2)
combox_title=ttk.Combobox(info_personal,values=["Mrs","Mr","Ms"])
combox_title.grid(row=1,column=2)


label_age=Label(info_personal,text="age")
label_age.grid(row=2,column=0)
age_spinbox=Spinbox(info_personal,from_=18,to=60)
age_spinbox.grid(row=3,column=0)



nationality_label=Label(info_personal,text="Nationality")
nationality_label.grid(row=2,column=1)
combox_nationality=ttk.Combobox(info_personal,values=["Maroc","France","japan","Italy"])
combox_nationality.grid(row=3,column=1)


#ajouter padding to widget in frame info_personal
for widget in info_personal.winfo_children():
    widget.grid_configure(pady=5,padx=10)

course_frame=LabelFrame(info_frame,text="courses")
course_frame.grid(row=1,column=0,sticky="news",pady=10,padx=20)

label_register=Label(course_frame,text="#register course")
label_register.grid(row=0,column=0)
Var=StringVar(value="Not Register")
checkbox_register=Checkbutton(course_frame,text="Currently Register",
                              variable=Var,onvalue="Register",offvalue="Not Register")

checkbox_register.grid(row=1,column=0)


complete_courses=Label(course_frame,text="course complete")
complete_courses.grid(row=0,column=1)
spinbox_courses=Spinbox(course_frame,from_=0,to='infinity')
spinbox_courses.grid(row=1,column=1)


semestre=Label(course_frame,text="# Semestre")
semestre.grid(row=0,column=2)
spinbox_semestre=Spinbox(course_frame,from_=0,to=7)
spinbox_semestre.grid(row=1,column=2)
for widget in course_frame.winfo_children():
    widget.grid_configure(pady=5,padx=10)


validation_var=StringVar(value="Not accepted")
term_condition=LabelFrame(info_frame,text="term ans condition")
term_condition.grid(row=2,column=0,sticky="news",padx=20,pady=10)
checkbox_term=Checkbutton(term_condition,text="i accept the term and condition",
                          variable=validation_var,onvalue="accepted",
                          offvalue="Not accepted")
checkbox_term.grid(row=0,column=0)


butoon=Button(info_frame,text="Submit",command=click)
butoon.grid(row=3,column=0,sticky="news",padx=20,pady=10)




window.mainloop()
