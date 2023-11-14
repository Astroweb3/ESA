import win32
import win32com
import numpy as np
import pandas as pd
import tkinter
from tkinter import*
import xlrd
from tkinter import filedialog
from tabulate import tabulate
from PIL import ImageTk
from datetime import datetime
import xlwt 
from xlwt import Workbook 
import openpyxl 
import collections
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.lib.pagesizes import letter, inch
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle
from operator import itemgetter
from collections import Counter 
from PIL import ImageTk,Image

   
top=Tk()
top.title("exam seating arragement")
canvas=Canvas(top)
bgimage=ImageTk.PhotoImage(Image.open('F://0.portfolio//all projects//ES//Exam seater//bg//bg.jpg'))
Label(top,image=bgimage).place(relwidth=1,relheight=1)

top.configure(background='black')

top.geometry("800x400")


#heading

head=Label(top,text="EXAM SEATING SOFTWERE",font=('times new roman',12,'bold'))
head.place(x = 300, y = 10)



def STUDENT_details():
    global RESULT
    global final_list
    global lists
    global session
    global SES
    global odd
    global even
    global NO_OF_STUDENTS
    global gride
    global DOE
    global date
    global DATE
    
    result_1=filedialog.askopenfilename(initialdir="/",title="select file",filetypes=(("excel files",".xls"),("all files","*.*")))
    df=pd.read_excel(result_1)
         #the code read the numbber of row and roll numbers
    wb=xlrd.open_workbook(result_1)
    sheet=wb.sheet_by_index(0)
    sheet.cell_value(0,0)
    N=sheet.nrows-1
    #print(N)
    session=[]
    gride=[]
    emp=[]
    lists=[]
    regno=[]
    name=[]
    dep=[]
    subcode=[]
    date=[]
    n=0#variable to count the cell in rows
    for i in range(n,N):
         x=df['REG NO'][n]
         y=df['STUDENT NAME'][n]
         a=df['SUB CODE'][n]
         ses=df['SESS'][n]
         doe=df['DOE'][n]
         z=df['DEP'][n]
         regno.append(x)
         name.append(y)
         gride.append(i)
         session.append(ses)
         
         subcode.append(a)
         dep.append(z)
         lists.append(i)
         
         date.append(doe)
         n+=1
    #heading=['REGNO','NAME','PAPCODE','REPCODE','SEATNO','P/A']     
    #print(lists)
    for i in range (0,len(gride)):
        #lists[i]=int(regno[i]),str(name[i]),str(papcode[i]),str(repcode[i]),int(i),emp
        gride[i]=regno[i],str(subcode[i]),str(dep[i]),str(session[i]),date[i]

    
    #print(len(lists))"""
    DOE=[]
  
    for x in date:        
       if x not in DOE: 
          DOE.append(x)
    
    #print(DOE)
    
    DATE=[]
    variable = StringVar(top)
    variable.set(DOE[0]) # default value

    w = OptionMenu(top, variable, *DOE)
    w.place(x=350,y=119)
    DATE.append(variable.get())
    #print (DATE)
    
    se=[]
    for y in session:
        if y not in se:
            se.append(y)
    
    SES=[]
    variable = StringVar(top)
    variable.set(se[0]) # default value

    w = OptionMenu(top, variable, *se)
    w.place(x=350,y=180)
    SES.append(variable.get())
    #print (SES)    
  

#label

heading1=Label(top,text="ENTER THE STUDENT DETAILS:",font=('times new roman',13,'bold'))
heading1.place(x = 10, y = 90)


button1=Button(top,bd=6,text=" BROWSE ",bg='grey',command=STUDENT_details,fg='white',font=('helvetica',13,'bold'))
button1.place(x=177,y=125)

def hall_details():
     global roomno
     global x
     global roW
     global coluM
     global TOTAL_NO_OF_SEATS
     global N 
     global roomid
     global size_of_room
     global size_of_room2
     global cap
    
     result_2=filedialog.askopenfilename(initialdir="/",title="select file",filetypes=(("excel files",".xls"),("all files","*.*")))
     df=pd.read_excel(result_2)
     #the code read the numbber of row
     wb=xlrd.open_workbook(result_2)
     sheet=wb.sheet_by_index(0)
     sheet.cell_value(0,0)
     N=sheet.nrows-1
     #print(N)
     cap=[]
     size_of_room2=[]
     size_of_room=[]
     roomid=[]
     roW=[]
     coluM=[]
     n=0#variable to count the cell in rows
     for i in range(n,N):
         row=df['ROWS'][n]
         colum=df['COLUMNS'][n]
         roomId=df['ROOMNO'][n]
         sizeofroom=df['CUMUCITY'][n]
         capacity=df['CAPACITY'][n]
         n+=1
         roW.append(row)
         coluM.append(colum)
         roomid.append(roomId) 
         size_of_room.append(sizeofroom)
         size_of_room2.append(sizeofroom)
         cap.append(capacity)
        #print(roW,coluM)
     mul=np.multiply(roW,coluM)
     size_of_room.insert(0,0)
     TOTAL_NO_OF_SEATS=sum(mul)
     
     #print(TOTAL_NO_OF_SEATS)
     
     #print(roomid)
     
     #print(size_of_room)
    
  
heading2=Label(top,text="ENTER THE ROOM DETAILS:",font=('times new roman',13,'bold'))
heading2.place(x = 10, y = 175)
 
button2=Button(top,bd=6,text=" BROWSE ",bg='grey',command=hall_details,fg='white',font=('helvetica',13,'bold'))
button2.place(x=175,y=205)





def seating_result():
    global list_of_student
    
    list_of_student=[]
    final_list=[]
    for n in range(len(gride)):
        for k in DATE:
            if k in list(gride[n]):
                for l in SES:
                    if l in list(gride[n]):
                         final_list.append(gride[n])  
    
    #print(code)
     
    #print(final_list)
    
    
    rooms_needed=[]
    total=0
    for item in range(len(cap)):
        total=total+cap[item]
        if total<len(final_list)+30:
            rooms_needed.append(roomid[item])         
    #print(rooms_needed)
    tables=[]
    seat=[]
    
    
    for it in range(0,len(final_list)):
        seat.append(it)
        tables.append(it)
     
    #print(seat)
    
    odd=[]
    
    for j in range(0,len(seat)):
        if (j% 2!=0):
            odd.append(j)
    
    #print(len(odd))
    even=[]
    
    for k in range(0,len(seat)):
        if( k %2 ==0):
            even.append(k)
    #print(len(even))
                         
                                                                                                                                                       
    
    for q in range(0,len(even)):
        for l in odd:
            #tables[l],lists[q]=lists[q],tables[l]
            seat[l],final_list[q]=final_list[q],seat[l]
            
    for a in range(len(odd),len(final_list)):
        for w in even:
            #tables[w],lists[a]=lists[a],tables[w]
            seat[w],final_list[a]=final_list[a],seat[w]
            
            
    none=[(0),(0),(0),(0),(0)]
    while(len(seat)<=size_of_room[len(rooms_needed)]-1):
             seat.append(none)
    
    result=[]
    da=DATE
    todays_date = "'" +datetime.now().strftime("%Y-%m-%d %H:%M") + '.xlsx' + "'"
    writer = pd.ExcelWriter('seating_{}.xlsx'.format(datetime.today().strftime('%y%m%d-%H%M%S')),) 
    for a in range(len(rooms_needed)):    
         result.append(rooms_needed[a])
         for c in range(size_of_room[a],size_of_room[a+1]):
              
               result.append(seat[c])
               #print(result)
               df=pd.DataFrame(result)
               df.to_excel(writer,sheet_name=rooms_needed[a])
         #print(result)
         result.clear()
    
    
         
    writer.close()
    cl=[]
    st=[]
    end=[]
    done=[]
    code=[]
    que=[]
    count=[]
    notice=[]
    notes=[]
    y=[]  
    writer1 = pd.ExcelWriter('notice_{}.xlsx'.format(datetime.today().strftime('%y%m%d-%H%M%S')),engine='xlsxwriter') 
    writer2 = pd.ExcelWriter('quePaper_{}.xlsx'.format(datetime.today().strftime('%y%m%d-%H%M%S')),engine='xlsxwriter') 

    for l in range(len(rooms_needed)):    
         #que.append(rooms_needed[l])
         done.append(rooms_needed[l])
         
         for f in range(size_of_room[l],size_of_room[l+1]):  
               que.append(seat[f])    
               for w in range(len(que)):
                    if que[w][1] not in code:
                        
                        code.append(que[w][1])
                      
         for g in range(0,len(code)):
             done.append(code[g])
             notice.append(rooms_needed[l])
             for h in range(len(que)):
                 if code[g] in que[h]:
                    
                    count.append(que[h])

             notice.append(count[0][2])
             notice.append(count[0][1])
             notice.append(count[-1][0])
             notice.append(count[0][0])
             
             
             
             note=np.array(notice)
             notes.append(note)
             g=pd.DataFrame(notes)
             g=pd.DataFrame(notes,columns=['ROOMNO','DEP','SUBCODE','FROM','TO'])
             
             g.to_excel(writer1,sheet_name='sheet1')       

             done.append(len(count))
             don=np.array(done)
             y.append(don)
             df1=pd.DataFrame(y,columns=['ROOMNO','CODE','SIZE'])
             df1.to_excel(writer2,sheet_name='sheet1')
             
             notice.clear()
             count.clear()
             done.clear()
         
         que.clear()
         
         code.clear()
    writer1.close()
    writer2.close()
    
    
    
heading3=Label(top,text="SEATING RESULT:",font=('times new roman',13,'bold'))
heading3.place(x = 10, y = 270)
            
button3=Button(top,bd=6,text=" RESULTS ",bg='grey',command=seating_result,fg='white',font=('helvetica',12,'bold'))
button3.place(x=175,y=295)
top.mainloop()
