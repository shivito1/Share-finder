import os
from tkinter import ttk
import subprocess
import re
from threading import Timer
from win32com.client import Dispatch
from tkinter import Tk,Listbox,W,E,END,ACTIVE

class Shortcut_:
    def __init__(self,root):
        self.root = root
        self.root.wm_title("Johns' Docs Mapper")
        self.shares1 = []
        self.shares2 = []
        self.shares3 = []
        self.listbox = Listbox(self.root,bg="gray",height=8,width=50)
        self.listbox.grid(row=5,rowspan=20,column=0,columnspan=5,sticky="e")
        self.lbl1 = ttk.Label(self.root,text='Loading... (Please Wait)')
        self.lbl1.grid(row=0,column=0,columnspan=3,sticky=W)
        bt2 = ttk.Button(self.root,text='Map Selected',command =self.Map_Selected)
        bt2.grid(row=0,column=4,columnspan=5,sticky=E)
        
        stMapper = Timer(0.01,self.formtable)
        stMapper.start()
        
        self.root.mainloop()
    def Map_Selected(self):

        self.listselect = str(self.listbox.get(ACTIVE))
            
        self.listselect = self.listselect.split(' \\ ')
        path1 = self.listselect[1]
        server1 = self.listselect[0]
        netpath = server1 + '\\' + path1

        target = '\\\\' + netpath
        wDir = ''

        shell = Dispatch('WScript.Shell')
        shortcut = shell.CreateShortCut(path1 +'.lnk')
        shortcut.Targetpath = target
        shortcut.WorkingDirectory = wDir
        
        shortcut.save()
        test = os.system('move /Y ' + '"' + path1 + '.lnk'+ '" ' + '%userprofile%\\desktop')
        self.root.destroy()
    def formtable(self):
        print("If this window displays 'System error 53' Please disregard "
              "\nand continue waiting for program to load Shares")
        self.network = subprocess.Popen('net view', stdout=subprocess.PIPE).communicate()
        self.device = str(self.network)
        self.device1 = []
        self.device2 =[]
        
        self.device = tuple(filter(None, self.device.split('\\n')))
        for item in self.device:        # Removes new line and other unwanted charicters "\n","\r"," "    
            self.device1.append(item.replace('\n', '').replace('\\r', '').replace(' ',''))
    #        self.device1.append(item)
            self.device = self.device1
        self.device2.append("localhost")
        for item in self.device:        # Removes unwanted "\\" this is added to the network location later.
            if item.startswith("\\\\"):    
                self.device2.append(re.sub('\\\\', '', item))
                self.device = self.device2
                
                
                
        for i in self.device:
            try:
                self.netviewshare = subprocess.Popen('net view %s' % i, stdout=subprocess.PIPE).communicate()
                self.shares1.append(self.netviewshare)
            except 1:
                pass
            
        for i in self.shares1:
            if 'There are no entries in the list' not in str(i):
                if 'Print' not in str(i):
                    string = str(i).replace('--','').replace('Disk','').replace('\\r\\nThe command completed successfully.','').replace('\\r\\nUsers','').replace('Shared resources at ','\\r\\n')
                    self.shares2.append(string)
            else:
                pass

        try:    
            for i in self.shares2:
                t = str(i).split('\\r\\n')
                
              
                self.shares3.append(t)
        except IndexError:
            pass
        
        for i in self.shares3:
            t = 8
            try: 
                for item in range(5):
                    if "', None" not in str(i[t]):
                        if str(i[t]) != '':
                            one = str(i[t])
                            one = one.split('  ')
                            
                            self.listbox.insert(END, str(i[1]) + " \\ " + str(one[0]))
                    t += 1
            except IndexError:
                pass
        self.lbl1.config(text=u'Done!  Make Selection Then Click \u2192',font='Times 10')


Shortcut_(Tk())
