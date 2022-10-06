#===========================================================================================
#                                       Project is completed
#==========================================================================================


from tkinter import *
from tkinter import filedialog
from tkinter import scrolledtext
import docxpy
import os
import docx2txt
import os
#========================================================================================
#                                    Work in progress
#========================================================================================
def open_file(event):
    #os.startfile(event)
    path=entry_browse.get()
    path=path.replace('/','\\',10)
    file_name=Lb1.get(Lb1.curselection())
    file_name=file_name.replace('>','\\')
    Total_path=path + file_name
    Total_path=Total_path.replace('  ','')
    print(Total_path)
    os.startfile(Total_path)
#========================================================================================

def final_search(dir_path,keyword):
    matched_file_list=[]
    all_docx_files=[]
    all_files_list=os.listdir(dir_path)
    for file in all_files_list: 
        if(file.endswith(".docx")): 
            all_docx_files.append(file)
            text=docxpy.process(f"{dir_path}/{file}")  
            if(keyword.lower() in text.lower()):
                matched_file_list.append(file)
    return all_files_list,all_docx_files,matched_file_list


#function for browse button
def browse():
    dir_path=filedialog.askdirectory()
    entry_browse.insert(0,dir_path)

def search():
    #----------->here
     Lb1.place(x=900,y=300)
     dir_path=entry_browse.get()
     keyword=entry_key.get()
     all_files,all_docx,matched_files =final_search(dir_path,keyword)
     #------->here
     Lb1.delete(1,END)    #to del previous result 1.0= from 0th char of ist line
     res_str=""
     res_str=res_str+f"Total files found:{len(all_files)}\n"
     res_str=res_str+f"Total Docx files found:{len(all_docx)}\n"
     res_str=res_str+f"Total Matched files found:{len(matched_files)}\n"
     for res in matched_files:
           res_str=res_str+"  >"+res+"\n"  #print("\t>",res)
     
     print(res_str.split('\n'))
     
     res_str=res_str.split('\n')
     #--------->here
     for i in range(len(res_str)):
         #Lb1.insert(END,res_str)
         Lb1.insert(i,res_str[i])
     

def reset():
    #print('this is reset button') #prints msg on console
    s=entry_browse.get()
    ln=len(s)
    entry_browse.delete(0,ln) #to clear username text field

    s=entry_key.get()
    ln=len(s)
    entry_key.delete(0,ln)
#------>here
    Lb1.delete(0,END)
    

root=Tk()
root.geometry('1200x700+10+0')
root.configure(bg='dark olive green')
#root.resizable(width=False,height=False)
root.title("Welcome window")

lbl_title=Label(root,text='WELCOME TO RESUME FINDER',font=('cambria',25,'bold'),bg='dark olive green',fg='white')
lbl_title.pack(pady=10)

lbl_browse=Label(root,text="Directory Path",font=('cambria',22,'italic'),bg="dark olive green",fg="white")
lbl_browse.place(x=280,y=220)

lbl_key=Label(root,text="Enter Keyword",font=('cambria',22,'italic'),bg="dark olive green",fg="white")
lbl_key.place(x=280,y=300)

entry_browse=Entry(root,font=('cambria',20),bg='dark olive green',fg='white',bd=5)
entry_browse.place(x=550,y=220)

entry_key=Entry(root,font=('cambria',20),bg='dark olive green',fg='white',bd=5)
entry_key.place(x=550,y=300)

btn_browse=Button(root,command=browse,text='Browse',font=('Book Antiqua',20),bg='dark olive green',fg='white')
btn_browse.place(x=900,y=220)

btn_srch=Button(root,command=search,text='Search',font=('Book Antiqua',15),bg='dark olive green',fg='white')
btn_srch.place(x=400,y=380)

btn_rst=Button(root,command=reset,text='Reset',font=('Book Antiqua',15),bg='dark olive green',fg='white')
btn_rst.place(x=550,y=380)

"""Lb1 = Listbox(root,font=('Book antiqua',15,'bold'),fg='black',selectmode='SINGLE',width=30)
Lb1.bind('<Button-1>',open_file)"""

root.mainloop()
