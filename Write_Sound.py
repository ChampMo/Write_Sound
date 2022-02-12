#pip install pypiwin32
import win32com.client as wincl

def Voice():
    r = R.get()
    rr = int(r)
    v = sound.get()
    vv = int(v)
    s = TI3.get()
    syn = wincl.Dispatch('SAPI.SpVoice')
    syn.Rate = rr
    syn.Volume = vv
    print (s)
    syn.speak(s)
    a = s+'\n'
    v2_result.set(a)

#GUI#

from tkinter import *
from tkinter import ttk

GUI = Tk()
GUI.geometry('400x580')
GUI.title('Voice')

FONT1 = ('Tahoma' ,20,'bold')
FONT2 = ('Tahoma' ,15)
FONT3 = ('' ,15)
FONT4 = ('' ,5)

L = ttk.Label(GUI,text ='',font = FONT4)
L.pack()
L3 = ttk.Label(GUI,text ='แปลงข้อความให้เป็นเสียง\n',font = FONT1)
L3.pack()

v2_result = StringVar()
v2_result.set('กรุณาใส่คำที่ต้องการ\n')
R2 = ttk.Label(textvariable = v2_result,font = FONT3)
R2.pack()

L3 = ttk.Label(GUI,text ='ความเร็ว\n',font = FONT2)
L3.pack()
Splist = ['1', '2', '3', '4']
R = ttk.Combobox(GUI, values = Splist,width = 4 ,font = FONT2)
R.current(0)
R.pack(ipadx=10,ipady=5)
L = ttk.Label(GUI,text ='',font = FONT3)
L.pack()

L3 = ttk.Label(GUI,text ='ระดับเสียง\n',font = FONT2)
L3.pack()
soundlist = ['0', '25', '50', '75', '100']
sound = ttk.Combobox(GUI, values = soundlist,width = 4 ,font = FONT2)
sound.current(4)
sound.pack(ipadx=10,ipady=5)
L = ttk.Label(GUI,text ='',font = FONT3)
L.pack()


L3 = ttk.Label(GUI,text ='ข้อความ\n',font = FONT2)
L3.pack()
TI3 = ttk.Entry(GUI)
TI3.pack(ipadx=60,ipady=6)
L4 = ttk.Label(GUI,text ='',font = FONT3)
L4.pack()

B2 = ttk.Button(GUI,text='Start!',command= Voice)
B2.pack(ipadx=100,ipady=6)
L5 = ttk.Label(GUI,text ='',font = FONT4)
L5.pack()

GUI.mainloop()