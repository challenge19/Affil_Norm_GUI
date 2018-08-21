# -*- coding: utf-8 -*-
"""
Created on Tue Aug 14 15:37:49 2018

@author: Zhesi Shen(zhesi.shen@live.com)
"""


from tkinter import Menu
from tkinter import *
import tkinter as tk
from tkinter import messagebox
from tkinter.ttk import Progressbar
from tkinter.filedialog import askopenfilename,askdirectory
from tkinter import ttk
import tkinter.font as tkFont
import difflib
import json
import pandas as pd
import numpy as np
import time

window = Tk()
window.title("Welcome to Affiliation Normalization")
window.lift()
#window.geometry('1650x450')


#windowWidth = window.winfo_reqwidth()
#windowHeight = window.winfo_reqheight()
#print("Width",windowWidth,"Height",windowHeight)
 
# Gets both half the screen width/height and window width/height
#positionRight = int(window.winfo_screenwidth()/2 - windowWidth/2)
#positionDown = int(window.winfo_screenheight()/2 - windowHeight/2)
#window.geometry("+{}+{}".format(positionRight, positionDown))

##--------- New or Continue Project -------------------------------------------
def ask_new_or_continue():
    new_window = Toplevel(window)
    new_window.lift(aboveThis=window)
    new_window.title('New or Continue')
#    new_window.geometry('600x450')
    user_lb = Label(new_window, text='User:')
    user_lb.pack()
    user_ety = Entry(new_window, width=20)
    user_ety.pack()
    def select_files():
        dirname = askdirectory() # show an "Open" dialog box and return the path to the selected file
        paras = {'current idx':1,
         'standard table':'./data/standard.xlsx',
         'candidate table':dirname+'/candidate.xlsx',
         'to be cleaned':dirname+'/to_cleaned.xlsx'}
        with open('./data/parameters.json', 'w') as outfile:
            json.dump(paras, outfile)
        return
    
    def click_new_project():
        select_files()
        new_window.destroy()
        window.lift()

        load_data()
        show()
        return 0
    new_btn = Button(new_window, text='New Project', command=click_new_project)
    new_btn.pack()
    
    def click_continue_project():
        new_window.destroy()
        window.lift()

        load_data()
        show()
    continue_btn = Button(new_window, text='Continue Project', command=click_continue_project)
    continue_btn.pack()

    return 0

ask_new_or_continue()
##---------- load parameters --------------------------------------------------
Clean_idx = IntVar()
Txt2Clean = StringVar()
MAX = IntVar()
LastSave = DoubleVar() # store save time
df_cand = pd.DataFrame()
df_toClean = pd.DataFrame()
paras = {}
def load_data():
    global df_cand,df_toClean,paras
    with open('./data/parameters.json') as json_file:  
        paras = json.load(json_file).copy()
    
    Clean_idx.set(max(1,paras['current idx']))
    LastSave.set(value=time.time())
#    df_stand = pd.read_excel(paras['standard table'])
    df_cand = pd.read_excel(paras['candidate table'])
    df_toClean = pd.read_excel(paras['to be cleaned'])
    MAX.set(df_toClean.shape[0])
    if 'uid' not in df_toClean.columns:
        df_toClean['uid']=np.nan
    if 'time' not in df_cand.columns:
        df_cand['time'] = np.nan
    if 'based on' not in df_cand.columns:
        df_cand['based on'] = np.nan
    bar['maximum'] = MAX.get()
    bar['value'] = Clean_idx.get()
    bar.update()

##----------- menu bar --------------------------------------------------------

menu = Menu(window) 

file_item = Menu(menu, tearoff=0)
file_item.add_command(label='New') 
file_item.add_command(label='Open')
file_item.add_command(label='Edit')
menu.add_cascade(label='File', menu=file_item)


def click_help_about():
    messagebox.showinfo('About', 'Affiliation Normalization\n\nZhesi Shen(zhesi.shen@live.com)')

help_item = Menu(menu, tearoff=0)
help_item.add_command(label='Tutorial')
help_item.add_command(label='About', command=click_help_about)
menu.add_cascade(label='Help', menu=help_item)
window.config(menu=menu)

##----------- show compared affiliations --------------------------------------

lbl_curr = Label(window, text="Current item: ")
lbl_curr.grid(column=0, row=0,sticky=tk.E)
lbl_cand = Label(window, text="Candidate item: ")
lbl_cand.grid(column=0, row=1,sticky=tk.E)




txt_curr = Text(window,font="Macro 12", width=120, height=1)
txt_curr.grid(column=1, row=0,columnspan=8,padx=2, pady=2,stick=tk.E)
lb_cand = Listbox(window, font="Macro 12", width=120, height=15)
lb_cand.grid(column=1, row=1,columnspan=8,padx=2, pady=2,stick=tk.E)




lbl_sim_threshold = Label(window, text=f'Similarity Threshold:')
lbl_sim_threshold.grid(column=1,row=2,sticky='e')

SimThreshold = DoubleVar(value=0.5)
txt_sim_threshold = Entry(window, textvariable=SimThreshold, width=8)
txt_sim_threshold.grid(column=2, row=2,sticky='w')

def reshow(event):
    show()
txt_sim_threshold.bind('<Return>', reshow)


def show_list_of_candidates(txt_cand,df_cand,text,threshold):
    lb_cand.delete(0,END)
    cands = []
    for i in range(df_cand.shape[0]):
        aff_cand = df_cand.iloc[i]['英文+省']
        unique_id = df_cand.iloc[i]['所级机构名_直译名']
        seqm = difflib.SequenceMatcher(None, text.lower(), aff_cand.lower())
        sim = seqm.ratio()
        if sim > threshold:
            cands.append((sim,aff_cand,unique_id,i))
    cands = sorted(cands,key=lambda e:-e[0])
    for i in range(len(cands)):
        sim,aff_cand,unique_id,od = cands[i]
#        row_format ="{:<.2f}  {:<90}  {:<},{}"
#        item = [sim,aff_cand,unique_id,od]
#        txt = row_format.format(*item)
        txt = f"{sim:.2f}     {aff_cand:<60}{unique_id},{od}"
#        print(txt)
        lb_cand.insert(END,txt)

def show_tobe_cleaned(text):
    txt_curr.delete('1.0', END)
    txt_curr.insert('1.0','         '+text)

def show():
    text1 = df_toClean.loc[Clean_idx.get()-1]['值-ins2_省份']
    Txt2Clean.set(text1)
    show_tobe_cleaned(Txt2Clean.get())
    show_list_of_candidates(lb_cand,df_cand,Txt2Clean.get(),SimThreshold.get())
#show_diff(text1,text2,txt_curr,txt_cand,txt_sim)
##--------- click botton ------------------------------------------------------

def autosave():
    if time.time() - LastSave.get() > 600:
        paras['current idx'] = bar['value']
        with open('./tmp/parameters.json', 'w') as outfile:
            json.dump(paras, outfile)
            
#        writer = pd.ExcelWriter('./tmp/standard.xlsx')
#        df_stand.to_excel(writer,'Sheet1')
#        writer.save()
    
        writer = pd.ExcelWriter('./tmp/candidate.xlsx')
        df_cand.to_excel(writer,'Sheet1')
        writer.save()
    
        writer = pd.ExcelWriter('./tmp/to_cleaned.xlsx')
        df_toClean.to_excel(writer,'Sheet1')
        writer.save()
        LastSave.set(time.time())
        print('Auto Saved')
    print('---'*20)

def click_previous():
    if 1 < Clean_idx.get():
        Clean_idx.set(Clean_idx.get()-1)
        bar["value"] = Clean_idx.get()
        bar.update()  #this works
        autosave()
        
        text1 = df_toClean.loc[Clean_idx.get()-1]['值-ins2_省份']
        Txt2Clean.set(text1)
        show_tobe_cleaned(Txt2Clean.get())
        show_list_of_candidates(lb_cand,df_cand,Txt2Clean.get(),SimThreshold.get())
    else:
        messagebox.showwarning('Warning', '没有上一个')  #shows warning message    
    
def click_next():
    if Clean_idx.get() < MAX.get():
        Clean_idx.set(Clean_idx.get()+1)
        bar["value"] = Clean_idx.get()
        bar.update()  #this works
        autosave()
        
        text1 = df_toClean.loc[Clean_idx.get()-1]['值-ins2_省份']
        Txt2Clean.set(text1)
        show_tobe_cleaned(Txt2Clean.get())
        show_list_of_candidates(lb_cand,df_cand,Txt2Clean.get(),SimThreshold.get())
    else:
        messagebox.showwarning('Warning', '没有下一个')  #shows warning message    


def append_to_cands(cands,cleans,row,uid):
    new_row = cands.shape[0]
    print(Txt2Clean.get())
    print(cands.loc[row])
    cands.loc[new_row] = cands.loc[row]
    cands.loc[new_row,'英文+省'] = Txt2Clean.get()
    cands.loc[new_row,'机构英文写法'] = Txt2Clean.get().split('_')[0]
    cands.loc[new_row,'省（直译名机构所在省市）'] = Txt2Clean.get().split('_')[-1]
    if cands.loc[new_row,'省（直译名机构所在省市）'] != cands.loc[row,'省（直译名机构所在省市）']:
        cands.loc[new_row,'市'] = np.nan

    cands.loc[new_row,'time'] = time.strftime("%Y/%m/%d")
    cands.loc[new_row,'based on'] = row
    cleans.loc[Clean_idx.get()-1,'uid'] = uid

def click_select():
    cand = lb_cand.get(ACTIVE)
    cand1 = cand.split(' ')
    sim = float(cand1[0])
    
    cand1 = cand1[-1].split(',')
    uid = cand1[0]
    row = cand1[-1]
    
    if sim<0.5:
        res = messagebox.askyesno('Save & Close','Do you really want to This One?')
        if res==False:
            print('Please Select Again')
            return
    
    print(cand)
    append_to_cands(df_cand,df_toClean,int(row),uid)
    click_next()
#    if Clean_idx.get()<MAX:
#        Clean_idx.set(Clean_idx.get()+1)
#        bar["value"] = Clean_idx.get()
#        bar.update()  #this works
#        autosave()
#        
#        text1 = df_toClean.loc[Clean_idx.get()-1]['值-ins2_省份']
#        Txt2Clean.set(text1)
#        show_tobe_cleaned(Txt2Clean.get())
#        show_list_of_candidates(lb_cand,df_cand,Txt2Clean.get(),SimThreshold.get())
#        
#    else:
#        messagebox.showwarning('Congratulations', 'All finished')  #shows warning message    

    return 0


def click_new():
    new_window = Toplevel(window)
    new_window.title('Inser New Entity')
    new_window.geometry('600x450')
    row = 0
    lbs = {}
    ets = {}
    for col in df_cand.columns:
        if col in ['所级机构名_直译名']: 
            lbs[col] = Label(new_window, text=f"{col}*:")
        else:
            lbs[col] = Label(new_window, text=f"{col}:")
        lbs[col].grid(column=0, row=row,sticky=tk.E)

        ets[col] = Entry(new_window, width=50)
        ets[col].grid(column=1, row=row)

        row += 1
    
    ets['英文+省'].insert(0,Txt2Clean.get())
    ets['time'].insert(0,time.strftime("%Y/%m/%d"))
    ets['机构英文写法'].insert(0,Txt2Clean.get().split('_')[0])
    ets['省（直译名机构所在省市）'].insert(0,Txt2Clean.get().split('_')[-1])

    
    def click_comfirm_insert():
        res = messagebox.askyesnocancel('Save & Close','确定插入本条？')
        if res==True:
            new_row = df_cand.shape[0]
            print(Txt2Clean.get())
            df_cand.loc[new_row] = [ets[col].get() for col in df_cand.columns]
            df_toClean.loc[Clean_idx.get()-1,'uid'] = ets['所级机构名_直译名'].get()

            new_window.destroy()
            
            click_next()
            
        elif res==False:
            new_window.destroy()
        else:
            new_window.lift()
    bt = Button(new_window, text="Save", command=click_comfirm_insert)
    bt.grid(column=1, row=row)
    
    autosave()
    return 0

    
    
btn_previous = Button(window, text="Previous", command=click_previous)
btn_previous.grid(column=5, row=2,pady=2,padx=6,sticky='we')

btn_next = Button(window, text="Next", command=click_next)
btn_next.grid(column=6, row=2,pady=2,padx=6,sticky='we')

btn_select = Button(window, text="Select", command=click_select)
btn_select.grid(column=7, row=2,pady=2,padx=6,sticky='we')

btn_new = Button(window, text="New", command=click_new)
btn_new.grid(column=8, row=2,pady=2,padx=6,sticky='we')



##------------- progress bar --------------------------------------------------


lbl_probar = Label(window, text=f'Completed:')
lbl_probar.grid(column=1,row=5,sticky='e')


style = ttk.Style()
style.theme_use('clam')
style.configure("black.Horizontal.TProgressbar", background='black')
bar = Progressbar(window, maximum=MAX.get(), length=200, style='grey.Horizontal.TProgressbar')
bar['value'] = Clean_idx.get()
bar.grid(column=2, row=5,columnspan=2,sticky='w')

## -------------- Save result -------------------------------------------------

def click_save():
    paras['current idx'] = Clean_idx.get()
    with open('./data/parameters.json', 'w') as outfile:
        json.dump(paras, outfile)
        
#    writer = pd.ExcelWriter(paras['standard table'])
#    df_stand.to_excel(writer,'Sheet1')
#    writer.save()

    writer = pd.ExcelWriter(paras['candidate table'])
    df_cand.to_excel(writer,'Sheet1')
    writer.save()

    writer = pd.ExcelWriter(paras['to be cleaned'])
    df_toClean.to_excel(writer,'Sheet1')
    writer.save()

    print('Saved')
    return 0

def click_close():
    res = messagebox.askyesnocancel('Save & Close','Do you want to Save&Close?')
    if res==True:
        click_save()
        window.destroy()
    elif res==False:
        print('close')
        window.destroy()
    else:
        print('nothing')

btn_save = Button(window, text="Save", command=click_save)
btn_save.grid(column=3, row=6,padx=6,sticky='we')

btn_close = Button(window, text="Close", command=click_close)
btn_close.grid(column=4, row=6,padx=6,sticky='we')




window.mainloop()