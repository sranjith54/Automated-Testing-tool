import sys

try:
   import Tkinter as tk
except ImportError:
    import tkinter as tk

try:
    import ttk
    py3 = False
except ImportError:
    import tkinter.ttk as ttk
    py3 = True

import gui_support
#import module_support
from tkinter import filedialog as fd
from tkinter import messagebox as msgbox
import os
import xlrd
import mymodule
import time
from xlutils.copy import copy


def vp_start_gui():
    '''Starting point when module is the main routine.'''
    global val, w, root
    try:
        root = tk.Tk()
    except:
        root = tk.Toplevel()
    root.iconbitmap('C:/Users/Admin/Desktop/ranjith/v1.0/v4.0/images/pricol_logo.ico')
    top = Automatic_Functional_Testing_Tool (root)
    
    gui_support.init(root, top)
    root.mainloop()

w = None
# =============================================================================
# def create_Automatic_Functional_Testing_Tool(root, *args, **kwargs):
#     '''Starting point when module is imported by another program.'''
#     global w, w_win, rt
#     rt = root
#     w = tk.Toplevel (root)
#     top = Automatic_Functional_Testing_Tool (w)
#     gui_support.init(w, top, *args, **kwargs)
#     return (w, top)
# =============================================================================

def destroy_Automatic_Functional_Testing_Tool():
    global w
    w.destroy()
    w = None


class Automatic_Functional_Testing_Tool():
    def __init__(self, top=None, *args, **kwargs):
        #This is where we lauch the file manager bar.
        if sys.platform == "darwin": CmdKey="Command-"
        else: CmdKey="Ctrl+"
        def OpenFile():
            name = fd.askopenfilename(initialdir="C:/Users/Admin/Desktop/",
                           filetypes =(("Text File", "*.txt"),("Excel File", "*.xlsx"),("All Files","*.*")),
                           title = "Choose a file."
                           )
            print (name)
    #Using try in case user types in unknown file or closes without choosing a file.
            try:
                with open(name,'r') as UseFile:
                    print(UseFile.read())
            except:
                print("No file exists")
        
        '''This class configures and populates the toplevel window.
           top is the toplevel containing window.'''
        _bgcolor = '#d9d9d9'  # X11 color: 'gray85'
        _fgcolor = '#000000'  # X11 color: 'black'
        _compcolor = '#d9d9d9' # X11 color: 'gray85'
        _ana1color = '#d9d9d9' # X11 color: 'gray85' 
        _ana2color = '#d9d9d9' # X11 color: 'gray85' 
        font10 = "-family {Segoe UI} -size 9 -weight bold -slant "  \
            "roman -underline 0 -overstrike 0"
        font11 = "-family {Berlin Sans FB Demi} -size 12 -weight bold "  \
            "-slant roman -underline 0 -overstrike 0"
        font9 = "-family {Segoe UI} -size 9 -weight bold -slant "  \
            "roman -underline 0 -overstrike 0"
        font15 = "-family {Segoe UI} -size 14 -weight bold -slant "  \
            "roman -underline 0 -overstrike 0"
        font16 = "-family {Segoe UI} -size 12 -weight bold -slant "  \
            "roman -underline 0 -overstrike 0"
        self.style = ttk.Style()
        if sys.platform == "win32":
            self.style.theme_use('winnative')
        self.style.configure('.',background=_bgcolor)
        self.style.configure('.',foreground=_fgcolor)
        self.style.configure('.',font="TkDefaultFont")
        self.style.map('.',background=
            [('selected', _compcolor), ('active',_ana2color)])

        top.geometry("800x580")
        top.title("Automatic Functional Testing Tool")
        #top.configure(background="#d5d3d0")
        top.configure(background="#730e8c")
        top.configure(highlightbackground="#d9d9d9")
        top.configure(highlightcolor="black")
        top.resizable(0,0)
        
          
              
       
        self.menubar = tk.Menu(top,font="TkMenuFont",bg=_bgcolor,fg=_fgcolor)
        self.file=tk.Menu(self.menubar, tearoff=0)
        self.file.add_command(
                
                label="New",
                accelerator=CmdKey+'N')
        self.file.add_command(
                
                label="Open",
                command=OpenFile, accelerator=CmdKey+'O')
        self.file.add_command(
                
                label="Save",
                accelerator=CmdKey+'S')
        self.file.add_command(
                
                label="Save As",
                accelerator="F12")
        self.file.add_command(
                
                label="Quit",
                command=quit,
                accelerator=CmdKey+'Q')
        
        
        self.menubar.add_cascade(label = 'File', menu = self.file)
        self.edit=tk.Menu(self.menubar, tearoff=0)
        self.edit.add_command(label="Cut", accelerator=CmdKey+'x')
        self.edit.add_command(label="Copy", accelerator=CmdKey+'c')
        self.edit.add_command(label="Module Configuration", accelerator=CmdKey+'M')
        self.menubar.add_cascade(label = 'Edit', menu = self.edit)
        self.view=tk.Menu(self.menubar, tearoff=0)
        self.view.add_command(label="Full Screen", accelerator=CmdKey+'F11')
        self.menubar.add_cascade(label = 'View', menu = self.view)
        self.tools=tk.Menu(self.menubar, tearoff=0)
        self.tools.add_command(label="Connection Port")
        self.menubar.add_cascade(label = 'Tools', menu = self.tools)
        self.help=tk.Menu(self.menubar, tearoff=0)
        self.help.add_command(label="Document", accelerator=CmdKey+'D')
        self.menubar.add_cascade(label="Help", menu=self.help)
        top.configure(menu = self.menubar)
        
        
        #Project Array for storage
        projectname=[]
        modulename=[]
        parameters={}
        testcasefile=[]
        failedtestcase_id=[]
        passedtestcase_id=[]
        #---------End----------
        
        



#------------------------ Test Case Report Screen Start----------------------
        #def connectionport():
            
        def page5():
            fail=len(failedtestcase_id)
            passed=len(passedtestcase_id)
            progprojectname=projectname[0]
            def ViewDetails():
                report=tk.Tk()
                report.iconbitmap('C:/Users/Admin/Desktop/ranjith/v1.0/v4.0/images/pricol_logo.ico')
                report.title("Test Report")
                report.geometry("800x580")
                report.configure(background="#730e8c")
                report.configure(highlightbackground="#d9d9d9")
                report.configure(highlightcolor="black")
                report.resizable(0,0)
                
                self.Frame_middle = tk.Frame(report)
                self.Frame_middle.tkraise()
                self.Frame_middle.place(relx=0.002, rely=0.09, relheight=0.90
                , relwidth=0.993)
                self.Frame_middle.configure(relief=tk.GROOVE)
                self.Frame_middle.configure(borderwidth="2")
                self.Frame_middle.configure(relief=tk.GROOVE)
                self.Frame_middle.configure(background="#ebffd9")
                self.Frame_middle.configure(highlightbackground="#d9d9d9")
                self.Frame_middle.configure(highlightcolor="black")
                self.Frame_middle.configure(width=565)
                
                self.Pprojname = tk.Label(report)
                self.Pprojname.place(relx=0.20, rely=0.03, height=31, width=500)
                self.Pprojname.configure(activebackground="#f9f9f9")
                self.Pprojname.configure(activeforeground="black")
                #self.Pprojname.configure(background="#ebffd9")
                self.Pprojname.configure(foreground="#730e8c")
                self.Pprojname.configure(font=font16)
                self.Pprojname.configure(disabledforeground="#a3a3a3")
            
                self.Pprojname.configure(highlightbackground="#d9d9d9")
                self.Pprojname.configure(highlightcolor="black")
                self.Pprojname.configure(text=" {}".format(progprojectname))
                
                
                self.TLabel7 = ttk.Label(self.Frame_middle)
                self.TLabel7.place(relx=0.02, rely=0.35, height=25, width=186)
                self.TLabel7.configure(background="#d2d5d3")
                self.TLabel7.configure(foreground="#000000")
                self.TLabel7.configure(font=font16)
                self.TLabel7.configure(relief=tk.FLAT)
                self.TLabel7.configure(text='''No. of Failed Test Cases:''')
                
                            
                self.Tfail = ttk.Label(self.Frame_middle)
                self.Tfail.place(relx=0.36, rely=0.35, height=39, width=126)
                self.Tfail.configure(background="#f9f9f9")
                self.Tfail.configure(foreground="#000000")
                self.Tfail.configure(font=font16)
                self.Tfail.configure(borderwidth="20")
                self.Tfail.configure(relief=tk.GROOVE)
                self.Tfail.configure(text="{}".format(fail))
                
                self.TLblpass = ttk.Label(self.Frame_middle)
                self.TLblpass.place(relx=0.02, rely=0.50, height=25, width=186)
                self.TLblpass.configure(background="#d2d5d3")
                self.TLblpass.configure(foreground="#000000")
                self.TLblpass.configure(font=font16)
                self.TLblpass.configure(relief=tk.FLAT)
                self.TLblpass.configure(text='''No. of  Test Cases Passed:''')
                
                self.Tlblpass = ttk.Label(self.Frame_middle)
                self.Tlblpass.place(relx=0.36, rely=0.50, height=39, width=126)
                self.Tlblpass.configure(background="#f9f9f9")
                self.Tlblpass.configure(foreground="#000000")
                self.Tlblpass.configure(font=font16)
                self.Tlblpass.configure(borderwidth="20")
                self.Tlblpass.configure(relief=tk.GROOVE)
                self.Tlblpass.configure(text="{}".format(passed))
                
                
            self.Frame_middle = tk.Frame(top)
            self.Frame_middle.tkraise()
            self.Frame_middle.place(relx=0.002, rely=0.09, relheight=0.90
                , relwidth=0.993)
            self.Frame_middle.configure(relief=tk.GROOVE)
            self.Frame_middle.configure(borderwidth="2")
            self.Frame_middle.configure(relief=tk.GROOVE)
            self.Frame_middle.configure(background="#ebffd9")
            self.Frame_middle.configure(highlightbackground="#d9d9d9")
            self.Frame_middle.configure(highlightcolor="black")
            self.Frame_middle.configure(width=565)
        
            
            self.Task = ttk.Label(self.Frame_middle)
            self.Task.place(relx=0.34, rely=0.05, height=90, width=181)
            self.Task.configure(background="#f9f9f9")
            self.Task.configure(foreground="#000000")
            self.Task.configure(font=font15)
            self.Task.configure(relief=tk.GROOVE)
            self.Task.configure(anchor=tk.CENTER)
            self.Task.configure(text='''Task Completed''')
            self._img1 = tk.PhotoImage(file="C:/Users/Admin/Desktop/ranjith/v1.0/v4.0/images/images.png")
            self.Task.configure(image=self._img1)
            self.Task.configure(compound="bottom")

            self.TLabel7 = ttk.Label(self.Frame_middle)
            self.TLabel7.place(relx=0.02, rely=0.36, height=25, width=186)
            self.TLabel7.configure(background="#d2d5d3")
            self.TLabel7.configure(foreground="#000000")
            self.TLabel7.configure(font=font16)
            self.TLabel7.configure(relief=tk.FLAT)
            self.TLabel7.configure(text='''No. of Failed Test Cases:''')

            self.TLabel8 = ttk.Label(self.Frame_middle)
            self.TLabel8.place(relx=0.36, rely=0.35, height=39, width=126)
            self.TLabel8.configure(background="#f9f9f9")
            self.TLabel8.configure(foreground="#000000")
            self.TLabel8.configure(font=font16)
            self.TLabel8.configure(borderwidth="20")
            self.TLabel8.configure(relief=tk.GROOVE)
            self.TLabel8.configure(text="{}".format(fail))
        
            self.Tviewdetails = ttk.Button(self.Frame_middle)
            self.Tviewdetails.place(relx=0.23, rely=0.52, height=50, width=300)
            self.Tviewdetails.configure(takefocus="")
            self.Tviewdetails.configure(text='''View Details''', command=ViewDetails)
            self.Tviewdetails.configure(width=295)
        
            self.TButton3 = ttk.Button(self.Frame_middle)
            self.TButton3.place(relx=0.23, rely=0.67, height=50, width=300)
            self.TButton3.configure(takefocus="")
            self.TButton3.configure(text='''Save Report''')
            self.TButton3.configure(width=296)

            self.TButton3 = ttk.Button(self.Frame_middle)
            self.TButton3.place(relx=0.23, rely=0.81, height=50, width=300)
            self.TButton3.configure(takefocus="")
            self.TButton3.configure(text='''Exit''')
            self.TButton3.configure(command=root.destroy)
            
        
        
        
#------------------------ Test Case Report Screen Start----------------------            
# -----------------------Test case progress Screen Start---------------------
        def page4():
            progprojectname=projectname[0]
            def starttest():
                self.btnstarttest.configure(state=tk.DISABLED)
                entry=[]
                filename=testcasefile[0]
                module=modulename[0]
                workbook = xlrd.open_workbook(filename)
                worksheet = workbook.sheet_by_name(module)
                num_rows = worksheet.nrows
                
                    #print(num_rows)
                num_cols = worksheet.ncols
                wb =copy(workbook)
                w_sheet = wb.get_sheet(module)
                for row_idx in range(2, num_rows):
                    progress=(((row_idx-1)/(num_rows-2))*100)
                    #self.TProgressbar1['value']=progress
                    self.TProgressbar1.configure(value="{}".format(progress))
                    self.TProgressbar1.update_idletasks()
                    self.lbltestpercentage.configure(text=" {}%".format(progress))
                    self.lblremainingtestcase.configure(text="{}/{}".format(row_idx-1, num_rows-2))
                    self.lblremainingtestcase.update_idletasks()
                    testcase_id=worksheet.cell(row_idx,2).value
                    self.Ptestname.configure(text="{}".format(testcase_id))
                    self.Ptestname.update_idletasks()
                    
                    x=worksheet.cell(row_idx, 3).value
                    
                    try:
                        lastName = x.split(';\n')
                        #print(lastName)
                        x=len(lastName)
                        for i in range(0,x):
                            key,value=lastName[i].split('=')
                            #print(key,value)
                            if key in mymodule.keyword:
                                mymodule.keyword[key]=value
                                entry.append(key)
                                #print(mymodule.keyword[key])
                                passedtestcase_id.append(testcase_id)
                                time.sleep(1)
                            else:
                                failedtestcase_id.append(testcase_id)
                                w_sheet.write(row_idx, 6, "Failed")
                                msgbox.showinfo("Error", "Invalid Key={} at {}".format(key, testcase_id))
                                
                    except ValueError:
                        print(" Do Not Use Space, New Line, While giving Pre Condition")
                    time.sleep(2)
                msgbox.showinfo("Sucess", "Test Cases Executed Sucessfully")
                wb.save('C:/Users/Admin/Desktop/ranjith/v1.0/testreport.xls')
                self.btnstarttest.configure(state=tk.NORMAL)
                self.Button4.configure(state=tk.NORMAL)
                    #excepted=worksheet.cell(row_idx, 4).value
                    #try:
                     #   lastName = excepted.split(';\n')
                        #print(lastName)
                      #  x=len(lastName)
                       # for i in range(0,x):
                        #    key,value=lastName[i].split('=')
                            #print(key,value)
                        #    if key in mymodule.keyword:
                        #        mymodule.keyword[key]=value
                        #        entry.append(key)
                        #        print(mymodule.keyword[key])
                         #   else:
                         #       print("Invalid Key")
                    #except ValueError:
                        #print(" Do Not Use Space, New Line, While giving Pre Condition")
                #print(entry)
            
            
            self.Frame_middle = tk.Frame(top)
            self.Frame_middle.tkraise()
            self.Frame_middle.place(relx=0.002, rely=0.09, relheight=0.90
                , relwidth=0.993)
            self.Frame_middle.configure(relief=tk.GROOVE)
            self.Frame_middle.configure(borderwidth="2")
            self.Frame_middle.configure(relief=tk.GROOVE)
            self.Frame_middle.configure(background="#ebffd9")
            self.Frame_middle.configure(highlightbackground="#d9d9d9")
            self.Frame_middle.configure(highlightcolor="black")
            self.Frame_middle.configure(width=565)
        
            self.TProgressbar1 = ttk.Progressbar(self.Frame_middle)
            self.TProgressbar1.place(relx=0.04, rely=0.26, relwidth=0.89
                , relheight=0.0, height=30)
            #self.TProgressbar1.configure(value="50")

            self.lblprogprojectname = ttk.Label(self.Frame_middle)
            self.lblprogprojectname.place(relx=0.04, rely=0.05, height=30, width=140)
            self.lblprogprojectname.configure(background="#ebffd9")
            self.lblprogprojectname.configure(foreground="#730e8c")
            self.lblprogprojectname.configure(font=font16)
            self.lblprogprojectname.configure(relief=tk.FLAT)
            self.lblprogprojectname.configure(text='''Project Name:''')
            
            self.Pprojname = tk.Label(self.Frame_middle)
            self.Pprojname.place(relx=0.25, rely=0.05, height=31, width=500)
            self.Pprojname.configure(activebackground="#f9f9f9")
            self.Pprojname.configure(activeforeground="black")
            #self.Pprojname.configure(background="#ebffd9")
            self.Pprojname.configure(foreground="#730e8c")
            self.Pprojname.configure(font=font16)
            self.Pprojname.configure(disabledforeground="#a3a3a3")
            
            self.Pprojname.configure(highlightbackground="#d9d9d9")
            self.Pprojname.configure(highlightcolor="black")
            self.Pprojname.configure(text=" {}".format(progprojectname))

            #self.Pfile = ttk.Label(self.Frame_middle)
            #self.Pfile.place(relx=0.4, rely=0.05, height=30, width=40)
            #self.Pfile.configure(background="#d1cdcb")
            #self.Pfile.configure(foreground="#000000")
            #self.Pfile.configure(font=font11)
            #self.Pfile.configure(relief=tk.FLAT)
            #self.Pfile.configure(text='''File:''')
            
            #self.lbltestfilename = tk.Label(self.Frame_middle)
            #self.lbltestfilename.place(relx=0.45, rely=0.05, height=31, width=114)
            #self.lbltestfilename.configure(activebackground="#f9f9f9")
            #self.lbltestfilename.configure(activeforeground="black")
            #self.lbltestfilename.configure(background="#d1cdcb")
            #self.lbltestfilename.configure(disabledforeground="#a3a3a3")
            #self.lbltestfilename.configure(foreground="#000000")
            #self.lbltestfilename.configure(highlightbackground="#d9d9d9")
            #self.lbltestfilename.configure(highlightcolor="black")
            #self.lbltestfilename.configure(text='''File name''')

            self.PprogressStatus = ttk.Label(self.Frame_middle)
            self.PprogressStatus.place(relx=0.04, rely=0.19, height=30, width=140)
            self.PprogressStatus.configure(background="#ebffd9")
            self.PprogressStatus.configure(foreground="#730e8c")
            self.PprogressStatus.configure(font=font16)
            self.PprogressStatus.configure(relief=tk.FLAT)
            self.PprogressStatus.configure(text='''Progession Status:''')

            self.Pcurrenttestcase = ttk.Label(self.Frame_middle)
            self.Pcurrenttestcase.place(relx=0.04, rely=0.44, height=30, width=140)
            self.Pcurrenttestcase.configure(background="#ebffd9")
            self.Pcurrenttestcase.configure(foreground="#730e8c")
            self.Pcurrenttestcase.configure(font=font16)
            self.Pcurrenttestcase.configure(relief=tk.FLAT)
            self.Pcurrenttestcase.configure(text='''Current Test Case:''')

            self.Ptestname = ttk.Label(self.Frame_middle)
            self.Ptestname.place(relx=0.24, rely=0.44, height=30, width=140)
            self.Ptestname.configure(background="#ebffd9")
            self.Ptestname.configure(foreground="#730e8c")
            self.Ptestname.configure(font=font16)
            self.Ptestname.configure(borderwidth="1")
            self.Ptestname.configure(relief=tk.FLAT)
            #self.Ptestname.configure(text='''TC_01_001''')

            

            

            self.lbltestpercentage = tk.Label(self.Frame_middle)
            self.lbltestpercentage.place(relx=0.75, rely=0.19, height=31, width=65)
            self.lbltestpercentage.configure(activebackground="#f9f9f9")
            self.lbltestpercentage.configure(activeforeground="black")
            self.lbltestpercentage.configure(background="#ebffd9")
            self.lbltestpercentage.configure(font=font16)
            self.lbltestpercentage.configure(disabledforeground="#a3a3a3")
            self.lbltestpercentage.configure(foreground="#00ca0f")
            self.lbltestpercentage.configure(highlightbackground="#d9d9d9")
            self.lbltestpercentage.configure(highlightcolor="black")
            #self.lbltestpercentage.configure(text='''50%''')

            self.lblremainingtestcase = tk.Label(self.Frame_middle)
            self.lblremainingtestcase.place(relx=0.60, rely=0.2, height=21, width=65)
            self.lblremainingtestcase.configure(activebackground="#f9f9f9")
            self.lblremainingtestcase.configure(activeforeground="black")
            self.lblremainingtestcase.configure(background="#ebffd9")
            self.lblremainingtestcase.configure(font=font16)
            self.lblremainingtestcase.configure(disabledforeground="#a3a3a3")
            self.lblremainingtestcase.configure(foreground="#00ca0f")
            self.lblremainingtestcase.configure(highlightbackground="#d9d9d9")
            self.lblremainingtestcase.configure(highlightcolor="black")
            #self.lblremainingtestcase.configure(text='''[1/2]''')
            
            self.btnstarttest = tk.Button(self.Frame_middle)
            self.btnstarttest.place(relx=0.44, rely=0.60, height=30, width=90)
            self.btnstarttest.configure(activebackground="#d9d9d9")
            self.btnstarttest.configure(activeforeground="#000000")
            self.btnstarttest.configure(background="#d9d9d9")
            self.btnstarttest.configure(disabledforeground="#a3a3a3")
            self.btnstarttest.configure(foreground="#000000")
            self.btnstarttest.configure(highlightbackground="#d9d9d9")
            self.btnstarttest.configure(highlightcolor="black")
            self.btnstarttest.configure(pady="0")
            self.btnstarttest.configure(text='''Start Test''', command=starttest)
            
            
            self.Button2 = tk.Button(self.Frame_middle)
            self.Button2.place(relx=0.54, rely=0.86, height=25, width=65)
            self.Button2.configure(activebackground="#d9d9d9")
            self.Button2.configure(activeforeground="#000000")
            self.Button2.configure(background="#d9d9d9")
            self.Button2.configure(disabledforeground="#a3a3a3")
            self.Button2.configure(foreground="#000000")
            self.Button2.configure(highlightbackground="#d9d9d9")
            self.Button2.configure(highlightcolor="black")
            self.Button2.configure(pady="0")
            self.Button2.configure(state=tk.NORMAL)
            self.Button2.configure(text='''<Back''', command=self.Frame_middle.tkraise())
        
            self.Button3 = tk.Button(self.Frame_middle)
            self.Button3.place(relx=0.63, rely=0.86, height=25, width=65)
            self.Button3.configure(activebackground="#d9d9d9")
            self.Button3.configure(activeforeground="#000000")
            self.Button3.configure(background="#d9d9d9")
            self.Button3.configure(disabledforeground="#a3a3a3")
            self.Button3.configure(foreground="#000000")
            self.Button3.configure(highlightbackground="#d9d9d9")
            self.Button3.configure(highlightcolor="black")
            self.Button3.configure(pady="0")
            self.Button3.configure(state=tk.DISABLED)
            self.Button3.configure(text='''Next>''', command=page5)
        
            self.Button4 = tk.Button(self.Frame_middle)
            self.Button4.place(relx=0.72, rely=0.86, height=25, width=65)
            self.Button4.configure(activebackground="#d9d9d9")
            self.Button4.configure(activeforeground="#000000")
            self.Button4.configure(background="#d9d9d9")
            self.Button4.configure(disabledforeground="#a3a3a3")
            self.Button4.configure(foreground="#000000")
            self.Button4.configure(highlightbackground="#d9d9d9")
            self.Button4.configure(highlightcolor="black")
            self.Button4.configure(pady="0")
            self.Button4.configure(state=tk.DISABLED)
            self.Button4.configure(text='''Finish''', command=page5)
        
            self.Button5 = tk.Button(self.Frame_middle)
            self.Button5.place(relx=0.81, rely=0.86, height=25, width=65)
            self.Button5.configure(activebackground="#d9d9d9")
            self.Button5.configure(activeforeground="#000000")
            self.Button5.configure(background="#d9d9d9")
            self.Button5.configure(disabledforeground="#a3a3a3")
            self.Button5.configure(foreground="#000000")
            self.Button5.configure(highlightbackground="#d9d9d9")
            self.Button5.configure(highlightcolor="black")
            self.Button5.configure(pady="0")
            self.Button5.configure(text='''Cancel''')

            
            
# -----------------------Test case progress Screen End-----------            
            
# -----------------Test Case Upload Screen Start------------------------------    
        def page3(pjname):
            global v
            def Openfile():
                filename = fd.askopenfilename(initialdir="C:/Users/Admin/Desktop/",
                           filetypes =(("Excel File", "*.xlsx"),("All Files","*.*")),
                           title = "Choose a file."
                           )
                testcasefile.insert(0, filename)
                
                self.txttestfilename.delete(0,tk.END)
                self.txttestfilename.insert(tk.END, "{}" .format(filename))
                self.Button3.configure(state=tk.NORMAL)
                #v = tk.StringVar(root, value="{}" .format(filename))
                #self.Projectname = tk.Entry(self.Frame_middle, textvariable=v)
                #self.Projectname.configure(text="{}" .format(filename))
                #workbook = xlrd.open_workbook(filename)
                #worksheet = workbook.sheet_by_name(modulename)
            
            self.Frame_middle = tk.Frame(top)
            self.Frame_middle.tkraise()
            self.Frame_middle.place(relx=0.002, rely=0.09, relheight=0.90
                , relwidth=0.993)
            self.Frame_middle.configure(relief=tk.GROOVE)
            self.Frame_middle.configure(borderwidth="2")
            self.Frame_middle.configure(relief=tk.GROOVE)
            self.Frame_middle.configure(background="#ebffd9")
            self.Frame_middle.configure(highlightbackground="#d9d9d9")
            self.Frame_middle.configure(highlightcolor="black")
            self.Frame_middle.configure(width=565)
            
            self.lbltestfilename = ttk.Label(self.Frame_middle)
            self.lbltestfilename.place(relx=0.01, rely=0.26, height=30, width=136)
            self.lbltestfilename.configure(background="#ebffd9")
            self.lbltestfilename.configure(foreground="#730e8c")
            self.lbltestfilename.configure(font=font16)
            self.lbltestfilename.configure(text='''Upload Test Case''')
            
            self.txttestfilename = tk.Entry(self.Frame_middle)
            self.txttestfilename.place(relx=0.19, rely=0.26,height=30, relwidth=0.49)
            self.txttestfilename.configure(background="white")
            self.txttestfilename.configure(borderwidth="2")
            self.txttestfilename.configure(disabledforeground="#a3a3a3")
            self.txttestfilename.configure(font="TkFixedFont")
            self.txttestfilename.configure(foreground="#000000")
            
            
            self.Btnbrowse = tk.Button(self.Frame_middle)
            self.Btnbrowse.place(relx=0.7, rely=0.26, height=30, width=130)
            self.Btnbrowse.configure(activebackground="#d9d9d9")
            self.Btnbrowse.configure(activeforeground="#000000")
            self.Btnbrowse.configure(background="#d9d9d9")
            self.Btnbrowse.configure(disabledforeground="#a3a3a3")
            self.Btnbrowse.configure(foreground="#000000")
            self.Btnbrowse.configure(highlightbackground="#d9d9d9")
            self.Btnbrowse.configure(highlightcolor="black")
            self.Btnbrowse.configure(pady="0")
            self.Btnbrowse.configure(text='''Browse''', command=Openfile)
            
            
            
            self.Button2 = tk.Button(self.Frame_middle)
            self.Button2.place(relx=0.54, rely=0.86, height=25, width=65)
            self.Button2.configure(activebackground="#d9d9d9")
            self.Button2.configure(activeforeground="#000000")
            self.Button2.configure(background="#d9d9d9")
            self.Button2.configure(disabledforeground="#a3a3a3")
            self.Button2.configure(foreground="#000000")
            self.Button2.configure(highlightbackground="#d9d9d9")
            self.Button2.configure(highlightcolor="black")
            self.Button2.configure(pady="0")
            self.Button2.configure(text='''<Back''', command=lambda:page2(pjname))
        
            self.Button3 = tk.Button(self.Frame_middle)
            self.Button3.place(relx=0.63, rely=0.86, height=25, width=65)
            self.Button3.configure(activebackground="#d9d9d9")
            self.Button3.configure(activeforeground="#000000")
            self.Button3.configure(background="#d9d9d9")
            self.Button3.configure(disabledforeground="#a3a3a3")
            self.Button3.configure(foreground="#000000")
            self.Button3.configure(highlightbackground="#d9d9d9")
            self.Button3.configure(highlightcolor="black")
            self.Button3.configure(pady="0")
            self.Button3.configure(state=tk.DISABLED)
            self.Button3.configure(text='''Next>''', command=page4)
        
            self.Button4 = tk.Button(self.Frame_middle)
            self.Button4.place(relx=0.72, rely=0.86, height=25, width=65)
            self.Button4.configure(activebackground="#d9d9d9")
            self.Button4.configure(activeforeground="#000000")
            self.Button4.configure(background="#d9d9d9")
            self.Button4.configure(disabledforeground="#a3a3a3")
            self.Button4.configure(foreground="#000000")
            self.Button4.configure(highlightbackground="#d9d9d9")
            self.Button4.configure(highlightcolor="black")
            self.Button4.configure(pady="0")
            self.Button4.configure(state=tk.DISABLED)
            self.Button4.configure(text='''Finish''')
        
            self.Button5 = tk.Button(self.Frame_middle)
            self.Button5.place(relx=0.81, rely=0.86, height=25, width=65)
            self.Button5.configure(activebackground="#d9d9d9")
            self.Button5.configure(activeforeground="#000000")
            self.Button5.configure(background="#d9d9d9")
            self.Button5.configure(disabledforeground="#a3a3a3")
            self.Button5.configure(foreground="#000000")
            self.Button5.configure(highlightbackground="#d9d9d9")
            self.Button5.configure(highlightcolor="black")
            self.Button5.configure(pady="0")
            self.Button5.configure(text='''Cancel''')
#------------------Test Case Upload Screen End--------------

#------------------Project Module screen Start--------------            
        def page2(pjname):
            
            
            category = {'Speed': ['SPD.FREQ','FUEL.FREQ','ODO.FREQ','TRIP.FREQ','TEMP.FREQ','SPD.DUTY', 'SPD.OFFSET','SPD.VOLTAGE'],
                        'Tack': ['TAC.FREQ','FUEL.FREQ','ODO.FREQ','TRIP.FREQ','TEMP.FREQ','TAC.DUTY', 'TAC.OFFSET','TAC.VOLTAGE']}
            def getUpdateData(event):
                self.TComboparameter['values'] = category[self.TModule.get()]
                self.TComboparameter.current(0)
               
            def module():
                if self.TModule.get():
                    if self.TComboEquipment.get():
                        if self.TComboinputtype.get():
                            if self.txtparametervalue.get():
                                        
                                if self.TModule.get()=='Speed':
                                    if 'Tack' in modulename:
                                        msgbox.showinfo("Error", "Tacko Already choosed")
                    
                                    else:    
                        
                                        if self.TModule.get() not in modulename:
                                            modulename.append(self.TModule.get())
                                            parameters[self.TComboinputtype.get()]=self.txtparametervalue.get()
                                            msgbox.showinfo("Success", "Module Has been Sucessfully updated")
                                            self.Button3.configure(state=tk.NORMAL)
                                        else:
                                            if self.TComboinputtype.get() in parameters:
                                                msgbox.showinfo("Error", "Parameter Already Assigned")
                                            else:
                                                parameters[self.TComboinputtype.get()]=self.txtparametervalue.get()
                                                msgbox.showinfo("Success", "Module Has been Sucessfully updated")
                                                self.Button3.configure(state=tk.NORMAL)
                                elif self.TModule.get()=='Tack':        
                                    if 'Speed' in modulename:
                                        msgbox.showinfo("Error", "Speed Already choosed")
                    
                                    else:    
                        
                                        if self.TModule.get() not in modulename:
                                            modulename.append(self.TModule.get())
                                            parameters[self.TComboinputtype.get()]=self.txtparametervalue.get()
                                            msgbox.showinfo("Success", "Module Has been Sucessfully updated")
                                            self.Button3.configure(state=tk.NORMAL)
                                        else:
                                            if self.TComboinputtype.get() in parameters:
                                                msgbox.showinfo("Error", "Parameter Already Assigned")
                                            else:
                                                parameters[self.TComboinputtype.get()]=self.txtparametervalue.get()
                                                msgbox.showinfo("Success", "Module Has been Sucessfully updated")
                                                self.Button3.configure(state=tk.NORMAL)
                            #print(modulename, parameters) 
                            else:
                                msgbox.showinfo("Error", "Enter the Parameter Value")
                                self.txtparametervalue.focus_set()
                        else:
                            msgbox.showinfo("Error", "Choose the Parameter")
                            self.TComboinputtype.focus_set()
                    else:
                        msgbox.showinfo("Error", "Choose the Eqipment")
                        self.TComboEquipment.focus_set()
                else:
                    msgbox.showinfo("Error", "Choose the Module")
                    self.TModule.focus_set()
               
                    
            self.Frame_middle = tk.Frame(top)
            self.Frame_middle.tkraise()
            self.Frame_middle.place(relx=0.002, rely=0.09, relheight=0.90
                , relwidth=0.993)
            self.Frame_middle.configure(relief=tk.GROOVE)
            self.Frame_middle.configure(borderwidth="2")
            self.Frame_middle.configure(relief=tk.GROOVE)
            self.Frame_middle.configure(background="#ebffd9")
            self.Frame_middle.configure(highlightbackground="#d9d9d9")
            self.Frame_middle.configure(highlightcolor="black")
            self.Frame_middle.configure(width=565)

            self.Module = ttk.Label(self.Frame_middle)
            self.Module.place(relx=0.03, rely=0.1, height=30, width=142)
            self.Module.configure(background="#ebffd9")
            self.Module.configure(foreground="#730e8c")
            self.Module.configure(font=font16)
            self.Module.configure(relief=tk.FLAT)
            self.Module.configure(text='''Module''')

            self.TModule = ttk.Combobox(self.Frame_middle)
            self.TModule.place(relx=0.25, rely=0.11, relheight=0.07, relwidth=0.40)
            #self.value_list = ["Speed","Tacko",]
            
            self.TModule.configure(values = list(category.keys()))
            self.TModule.bind('<<ComboboxSelected>>', getUpdateData)
            #self.TModule.configure(textvariable=module_support.combobox)
            self.TModule.current(0)  
            self.TModule.configure(takefocus="")

            self.TEquipment = ttk.Label(self.Frame_middle)
            self.TEquipment.place(relx=0.03, rely=0.2, height=30, width=144)
            self.TEquipment.configure(background="#ebffd9")
            self.TEquipment.configure(foreground="#730e8c")
            self.TEquipment.configure(font=font16)
            self.TEquipment.configure(relief=tk.FLAT)
            self.TEquipment.configure(text='''Equipment Details''')

            self.TComboEquipment = ttk.Combobox(self.Frame_middle)
            self.TComboEquipment.place(relx=0.25, rely=0.2, relheight=0.07
                , relwidth=0.40)
            #self.TComboEquipment.configure(textvariable=module_support.combobox)
            
            self.TComboEquipment.configure(takefocus="")
            self.TComboEquipment.configure(value="USB0::0x0957::0x179B::MY52497845::INSTR")

            self.Tlblinputtype = ttk.Label(self.Frame_middle)
            self.Tlblinputtype.place(relx=0.03, rely=0.3, height=30, width=142)
            self.Tlblinputtype.configure(background="#ebffd9")
            self.Tlblinputtype.configure(foreground="#730e8c")
            self.Tlblinputtype.configure(font=font16)
            self.Tlblinputtype.configure(relief=tk.FLAT)
            self.Tlblinputtype.configure(text='''Input Type''')

            

            self.TComboinputtype = ttk.Combobox(self.Frame_middle)
            self.TComboinputtype.place(relx=0.25, rely=0.3, relheight=0.07
                , relwidth=0.40)
            #self.TComboEquipment_1.configure(textvariable=module_support.combobox)
            self.TComboinputtype.configure(takefocus="")

            self.Tlableparametervalue = ttk.Label(self.Frame_middle)
            self.Tlableparametervalue.place(relx=0.03, rely=0.41, height=30, width=142)
            self.Tlableparametervalue.configure(background="#ebffd9")
            self.Tlableparametervalue.configure(foreground="#730e8c")
            self.Tlableparametervalue.configure(font=font16)
            self.Tlableparametervalue.configure(relief=tk.FLAT)
            self.Tlableparametervalue.configure(text='''Parameter Value''')

            self.txtparametervalue = tk.Entry(self.Frame_middle)
            self.txtparametervalue.place(relx=0.25, rely=0.41,height=30, relwidth=0.40)
            self.txtparametervalue.configure(background="white")
            self.txtparametervalue.configure(disabledforeground="#a3a3a3")
            self.txtparametervalue.configure(font="TkFixedFont")
            self.txtparametervalue.configure(foreground="#000000")
            self.txtparametervalue.configure(highlightbackground="#d9d9d9")
            self.txtparametervalue.configure(highlightcolor="black")
            self.txtparametervalue.configure(insertbackground="black")
            self.txtparametervalue.configure(selectbackground="#c4c4c4")
            self.txtparametervalue.configure(selectforeground="black")
            
            self.add = ttk.Button(self.Frame_middle)
            self.add.place(relx=0.32, rely=0.55, height=30, width=180)
            self.add.configure(takefocus="")
            self.add.configure(text='''ADD''', command=module)

            self.Label1 = tk.Label(self.Frame_middle)
            self.Label1.place(relx=0.1, rely=0.04, height=28, width=600)
            self.Label1.configure(background="#ebffd9")
            self.Label1.configure(foreground="#730e8c")
            self.Label1.configure(font=font16)
            self.Label1.configure(foreground="#730e8c")
            self.Label1.configure(text=" {}".format(pjname))
            
            self.Button2 = tk.Button(self.Frame_middle)
            self.Button2.place(relx=0.54, rely=0.86, height=25, width=65)
            self.Button2.configure(activebackground="#d9d9d9")
            self.Button2.configure(activeforeground="#000000")
            self.Button2.configure(background="#d9d9d9")
            self.Button2.configure(disabledforeground="#a3a3a3")
            self.Button2.configure(foreground="#000000")
            self.Button2.configure(highlightbackground="#d9d9d9")
            self.Button2.configure(highlightcolor="black")
            self.Button2.configure(pady="0")
            self.Button2.configure(text='''<Back''', command=page1)


        
            self.Button3 = tk.Button(self.Frame_middle)
            self.Button3.place(relx=0.63, rely=0.86, height=25, width=65)
            self.Button3.configure(activebackground="#d9d9d9")
            self.Button3.configure(activeforeground="#000000")
            self.Button3.configure(background="#d9d9d9")
            self.Button3.configure(disabledforeground="#a3a3a3")
            self.Button3.configure(foreground="#000000")
            self.Button3.configure(highlightbackground="#d9d9d9")
            self.Button3.configure(highlightcolor="black")
            self.Button3.configure(pady="0")
            self.Button3.configure(state=tk.DISABLED)
            self.Button3.configure(text='''Next>''', command=lambda:page3(pjname))
        
            self.Button4 = tk.Button(self.Frame_middle)
            self.Button4.place(relx=0.72, rely=0.86, height=25, width=65)
            self.Button4.configure(activebackground="#d9d9d9")
            self.Button4.configure(activeforeground="#000000")
            self.Button4.configure(background="#d9d9d9")
            self.Button4.configure(disabledforeground="#a3a3a3")
            self.Button4.configure(foreground="#000000")
            self.Button4.configure(highlightbackground="#d9d9d9")
            self.Button4.configure(highlightcolor="black")
            self.Button4.configure(pady="0")
            self.Button4.configure(state=tk.DISABLED)
            self.Button4.configure(text='''Finish''')
        
            self.Button5 = tk.Button(self.Frame_middle)
            self.Button5.place(relx=0.81, rely=0.86, height=25, width=65)
            self.Button5.configure(activebackground="#d9d9d9")
            self.Button5.configure(activeforeground="#000000")
            self.Button5.configure(background="#d9d9d9")
            self.Button5.configure(disabledforeground="#a3a3a3")
            self.Button5.configure(foreground="#000000")
            self.Button5.configure(highlightbackground="#d9d9d9")
            self.Button5.configure(highlightcolor="black")
            self.Button5.configure(pady="0")
            self.Button5.configure(text='''Cancel''')

#------------------Project Module screen Start--------------            
            
            
            
#------------------Project Creation screen Start--------------                        
        def page1():
            #Creat project Start
            def create_project():
                if self.txtProjectname.get():
                    if self.txtprodescription.get():
                        if self.txtbuild.get():
                            if self.txtauto.get():
                                name = self.txtProjectname.get()
                                projectname.insert(0, name)
                                des = self.txtprodescription.get()
                                build1 = self.txtbuild.get()
                                auto1 = self.txtauto.get()
                                FileName = str("D:\\" + name + ".AFT")# get text from entry
                                if os.path.isfile(FileName):
                                    msgbox.showinfo("Error", "File Already Exist")
                                       #print("File Already Exist")
                                else:
                    
                                    with open(FileName, "w") as f: # open file
                                        f.write("Project Name: {} \n".format(name))
                                        f.write("Project Description: {} \n".format(des))
                                        f.write("Build Version: {} \n".format(build1))
                                        f.write("Automation: {} \n".format(auto1))
                                        msgbox.showinfo("Save", "Project Created Sucessfully")
                                        self.Button3.configure(state=tk.NORMAL)
                                        self.btnsaveproject.configure(state=tk.DISABLED)
                                        
                            else:
                                msgbox.showinfo("Mantatory", "Fill All the text Box")
                                self.txtauto.focus_set()
                        else:
                            msgbox.showinfo("Mantatory", "Fill All the text Box")
                            self.txtbuild.focus_set()                               
                    else:
                        msgbox.showinfo("Mantatory", "Fill All the text Box")
                        self.txtprodescription.focus_set()
                else:
                    msgbox.showinfo("Mantatory", "Fill All the text Box")
                    self.txtProjectname.focus_set()
              #Create Project End     
#Frame Middle create project screen start
            font16 = "-family Calibri -size 13 -weight bold -slant roman "  \
            "-underline 0 -overstrike 0"
            self.Frame_middle = tk.Frame(top)
            self.Frame_middle.place(relx=0.002, rely=0.09, relheight=0.90
                , relwidth=0.993)
            self.Frame_middle.configure(relief=tk.GROOVE)
            self.Frame_middle.configure(borderwidth="2")
            self.Frame_middle.configure(relief=tk.GROOVE)
            self.Frame_middle.configure(background="#ebffd9")
            self.Frame_middle.configure(highlightbackground="#d9d9d9")
            self.Frame_middle.configure(highlightcolor="black")
            self.Frame_middle.configure(width=785)

            
            self.txtProjectname = tk.Entry(self.Frame_middle)
            self.txtProjectname.place(relx=0.33, rely=0.14, height=25, relwidth=0.42)

            self.txtProjectname.configure(background="white")
            self.txtProjectname.configure(borderwidth="3")
            self.txtProjectname.configure(disabledforeground="#a3a3a3")
            self.txtProjectname.configure(font="TkFixedFont")
            self.txtProjectname.configure(foreground="#000000")
            self.txtProjectname.configure(highlightbackground="#000000")
            self.txtProjectname.configure(highlightcolor="black")
            self.txtProjectname.configure(insertbackground="black")
            self.txtProjectname.configure(selectbackground="#c4c4c4")
            self.txtProjectname.configure(selectforeground="black")
            
            

            self.lblprojectname = ttk.Label(self.Frame_middle)
            self.lblprojectname.place(relx=0.04, rely=0.14, height=30, width=180)
            self.lblprojectname.configure(background="#ebffd9")
            self.lblprojectname.configure(foreground="#730e8c")
            self.lblprojectname.configure(font=font16)
            self.lblprojectname.configure(relief=tk.FLAT)
            self.lblprojectname.configure(text='''Project Name''')
            self.lblprojectname.configure(width=166)

            self.btnsaveproject = tk.Button(self.Frame_middle)
            self.btnsaveproject.place(relx=0.23, rely=0.74, height=30, width=130)
            self.btnsaveproject.configure(activebackground="#d9d9d9")
            self.btnsaveproject.configure(activeforeground="#000000")
            self.btnsaveproject.configure(background="#004000")
            self.btnsaveproject.configure(disabledforeground="#a3a3a3")
            self.btnsaveproject.configure(foreground="#ffffff")
            self.btnsaveproject.configure(highlightbackground="#d9d9d9")
            self.btnsaveproject.configure(highlightcolor="black")
            self.btnsaveproject.configure(pady="0")
            self.btnsaveproject.configure(text='''Save''', command=create_project)

            self.lblprodescription = tk.Label(self.Frame_middle)
            self.lblprodescription.place(relx=0.04, rely=0.28, height=30, width=180)
            self.lblprodescription.configure(anchor=tk.W)
            self.lblprodescription.configure(background="#ebffd9")
            self.lblprodescription.configure(disabledforeground="#a3a3a3")
            self.lblprodescription.configure(font=font16)
            self.lblprodescription.configure(foreground="#730e8c")
            self.lblprodescription.configure(justify=tk.LEFT)
            self.lblprodescription.configure(text='''Project Description''')
            self.lblprodescription.configure(width=201)

            self.projectdesc=tk.StringVar()
            self.txtprodescription = tk.Entry(self.Frame_middle, textvariable=self.projectdesc)
            self.txtprodescription.place(relx=0.33, rely=0.28, height=25
                , relwidth=0.42)
            self.txtprodescription.configure(background="white")
            self.txtprodescription.configure(borderwidth="3")
            self.txtprodescription.configure(disabledforeground="#a3a3a3")
            self.txtprodescription.configure(font="TkFixedFont")
            self.txtprodescription.configure(foreground="#000000")
            self.txtprodescription.configure(insertbackground="black")

            self.lblbuild = tk.Label(self.Frame_middle)
            self.lblbuild.place(relx=0.04, rely=0.41, height=30, width=180)
            self.lblbuild.configure(anchor=tk.W)
            self.lblbuild.configure(background="#ebffd9")
            self.lblbuild.configure(disabledforeground="#a3a3a3")
            self.lblbuild.configure(font=font16)
            self.lblbuild.configure(foreground="#730e8c")
            self.lblbuild.configure(text='''Build Version''')

            self.build=tk.StringVar()
            self.txtbuild = tk.Entry(self.Frame_middle, textvariable=self.build)
            self.txtbuild.place(relx=0.33, rely=0.41,height=25, relwidth=0.42)
            self.txtbuild.configure(background="white")
            self.txtbuild.configure(borderwidth="3")
            self.txtbuild.configure(disabledforeground="#a3a3a3")
            self.txtbuild.configure(font="TkFixedFont")
            self.txtbuild.configure(foreground="#000000")
            self.txtbuild.configure(insertbackground="black")

            self.lblauto = tk.Label(self.Frame_middle)
            self.lblauto.place(relx=0.04, rely=0.55, height=30, width=180)
            self.lblauto.configure(anchor=tk.W)
            self.lblauto.configure(background="#ebffd9")
            self.lblauto.configure(disabledforeground="#a3a3a3")
            self.lblauto.configure(font=font16)
            self.lblauto.configure(foreground="#730e8c")
            self.lblauto.configure(text='''Automation Prepared by''')

            self.autoprepared=tk.StringVar()
            self.txtauto = tk.Entry(self.Frame_middle, textvariable=self.autoprepared)
            self.txtauto.place(relx=0.33, rely=0.55,height=25, relwidth=0.42)
            self.txtauto.configure(background="white")
            self.txtauto.configure(borderwidth="3")
            self.txtauto.configure(disabledforeground="#a3a3a3")
            self.txtauto.configure(font="TkFixedFont")
            self.txtauto.configure(foreground="#000000")
            self.txtauto.configure(insertbackground="black")

            self.btnprojreset = tk.Button(self.Frame_middle)
            self.btnprojreset.place(relx=0.41, rely=0.74, height=30, width=130)
            self.btnprojreset.configure(activebackground="#d9d9d9")
            self.btnprojreset.configure(activeforeground="#000000")
            self.btnprojreset.configure(background="#004000")
            self.btnprojreset.configure(disabledforeground="#a3a3a3")
            self.btnprojreset.configure(foreground="#ffffff")
            self.btnprojreset.configure(highlightbackground="#d9d9d9")
            self.btnprojreset.configure(highlightcolor="black")
            self.btnprojreset.configure(pady="0")
            self.btnprojreset.configure(text='''Reset''')
        
            self.Button2 = tk.Button(self.Frame_middle)
            self.Button2.place(relx=0.54, rely=0.86, height=25, width=65)
            self.Button2.configure(activebackground="#d9d9d9")
            self.Button2.configure(activeforeground="#000000")
            self.Button2.configure(background="#d9d9d9")
            self.Button2.configure(disabledforeground="#a3a3a3")
            self.Button2.configure(foreground="#000000")
            self.Button2.configure(highlightbackground="#d9d9d9")
            self.Button2.configure(highlightcolor="black")
            self.Button2.configure(pady="0")
            self.Button2.configure(state=tk.DISABLED)
            self.Button2.configure(text='''<Back''', command=self.Frame_middle.tkraise())
        
            self.Button3 = tk.Button(self.Frame_middle)
            self.Button3.place(relx=0.63, rely=0.86, height=25, width=65)
            self.Button3.configure(activebackground="#d9d9d9")
            self.Button3.configure(activeforeground="#000000")
            self.Button3.configure(background="#d9d9d9")
            self.Button3.configure(disabledforeground="#a3a3a3")
            self.Button3.configure(foreground="#000000")
            self.Button3.configure(highlightbackground="#d9d9d9")
            self.Button3.configure(highlightcolor="black")
            self.Button3.configure(pady="0")
            self.Button3.configure(state=tk.DISABLED)
            self.Button3.configure(text='''Next>''', command=lambda:page2(self.txtProjectname.get()))
        
            self.Button4 = tk.Button(self.Frame_middle)
            self.Button4.place(relx=0.72, rely=0.86, height=25, width=65)
            self.Button4.configure(activebackground="#d9d9d9")
            self.Button4.configure(activeforeground="#000000")
            self.Button4.configure(background="#d9d9d9")
            self.Button4.configure(disabledforeground="#a3a3a3")
            self.Button4.configure(foreground="#000000")
            self.Button4.configure(highlightbackground="#d9d9d9")
            self.Button4.configure(highlightcolor="black")
            self.Button4.configure(pady="0")
            self.Button4.configure(state=tk.DISABLED)
            self.Button4.configure(text='''Finish''')
        
            self.Button5 = tk.Button(self.Frame_middle)
            self.Button5.place(relx=0.81, rely=0.86, height=25, width=65)
            self.Button5.configure(activebackground="#d9d9d9")
            self.Button5.configure(activeforeground="#000000")
            self.Button5.configure(background="#d9d9d9")
            self.Button5.configure(disabledforeground="#a3a3a3")
            self.Button5.configure(foreground="#000000")
            self.Button5.configure(highlightbackground="#d9d9d9")
            self.Button5.configure(highlightcolor="black")
            self.Button5.configure(pady="0")
            self.Button5.configure(text='''Cancel''')
            #Frame Middle create project screen End
# =========================================================================================================
#------------------Project Creation  Screen End--------------             
        
        # Frame Top Start
        self.frmae_top = ttk.Frame(top)
        self.frmae_top.place(relx=0.002, rely=0.001, relheight=0.08, relwidth=0.993)
        self.frmae_top.configure(relief=tk.GROOVE)
        self.frmae_top.configure(relief=tk.GROOVE)
        self.frmae_top.configure(width=785)
        self.bttopnew = tk.Button(self.frmae_top)
        self.bttopnew.place(relx=0.01, rely=0.11, height=35, width=35)
        self.bttopnew.configure(activebackground="#d9d9d9")
        self.bttopnew.configure(activeforeground="#000000")
        self.bttopnew.configure(background="#d9d9d9")
        self.bttopnew.configure(disabledforeground="#a3a3a3")
        self.bttopnew.configure(foreground="#000000")
        self.bttopnew.configure(highlightbackground="#d9d9d9")
        self.bttopnew.configure(highlightcolor="black")
        self._img5 = tk.PhotoImage(file="C:/Users/Admin/Desktop/ranjith/v1.0/v4.0/images/1.png")
        self.bttopnew.configure(image=self._img5)
        self.bttopnew.configure(pady="0")
        self.bttopnew.configure(text='''New''', command=page1)

        self.bttopsave = tk.Button(self.frmae_top)
        self.bttopsave.place(relx=0.06, rely=0.11, height=35, width=35)
        self.bttopsave.configure(activebackground="#d9d9d9")
        self.bttopsave.configure(activeforeground="#000000")
        self.bttopsave.configure(background="#d9d9d9")
        self.bttopsave.configure(disabledforeground="#a3a3a3")
        self.bttopsave.configure(foreground="#000000")
        self.bttopsave.configure(highlightbackground="#d9d9d9")
        self.bttopsave.configure(highlightcolor="black")
        self._img6 = tk.PhotoImage(file="C:/Users/Admin/Desktop/ranjith/v1.0/v4.0/images/save1.png")
        self.bttopsave.configure(image=self._img6)
        self.bttopsave.configure(pady="0")
        self.bttopsave.configure(text='''Save''')

        self.bttopsaveas = tk.Button(self.frmae_top)
        self.bttopsaveas.place(relx=0.11, rely=0.11, height=35, width=35)
        self.bttopsaveas.configure(activebackground="#d9d9d9")
        self.bttopsaveas.configure(activeforeground="#000000")
        self.bttopsaveas.configure(background="#d9d9d9")
        self.bttopsaveas.configure(disabledforeground="#a3a3a3")
        self.bttopsaveas.configure(foreground="#000000")
        self.bttopsaveas.configure(highlightbackground="#d9d9d9")
        self.bttopsaveas.configure(highlightcolor="black")
        self._img7 = tk.PhotoImage(file="C:/Users/Admin/Desktop/ranjith/v1.0/v4.0/images/save_as.png")
        self.bttopsaveas.configure(image=self._img7)
        self.bttopsaveas.configure(pady="0")
        self.bttopsaveas.configure(text='''Save As''')
        #start Top Ends
        #--------------------------------------------------------------------
        # Frmae Middle Starts
        self.Frame_middle = tk.Frame(top)
        self.Frame_middle.tkraise()
        self.Frame_middle.place(relx=0.002, rely=0.09, relheight=0.90
                , relwidth=0.993)
        self.Frame_middle.configure(relief=tk.GROOVE)
        self.Frame_middle.configure(borderwidth="2")
        self.Frame_middle.configure(relief=tk.GROOVE)
        self.Frame_middle.configure(background="#d9d9d9")
        self.Frame_middle.configure(highlightbackground="#d9d9d9")
        self.Frame_middle.configure(highlightcolor="black")
        self.Frame_middle.configure(width=565)
        
        self.pricollogo = tk.Label(self.Frame_middle)
        self.pricollogo.place(relx=0.0, rely=0.07, height=100, width=778)
        self.pricollogo.configure(activebackground="SystemButtonText")
        self.pricollogo.configure(activeforeground="white")
        self.pricollogo.configure(background="#ffffff")
        self.pricollogo.configure(disabledforeground="#a3a3a3")
        self.pricollogo.configure(font=font9)
        self.pricollogo.configure(foreground="#ffffff")
        self._img1 = tk.PhotoImage(file="C:/Users/Admin/Desktop/ranjith/v1.0/v4.0/images/logo1.png")
        self.pricollogo.configure(image=self._img1)
        self.pricollogo.configure(relief=tk.RIDGE)
        self.pricollogo.configure(text='''Automated functional Testing Tool''')
        self.pricollogo.configure(width=780)

        self.AFT = tk.Label(self.Frame_middle)
        self.AFT.place(relx=0.0, rely=0.3, height=35, width=778)
        self.AFT.configure(background="#730e8c")
        self.AFT.configure(disabledforeground="#a3a3a3")
        self.AFT.configure(font=font9)
        self.AFT.configure(foreground="#ffffff")
        self.AFT.configure(relief=tk.RIDGE)
        self.AFT.configure(text='''Automated Functional Testing Tool''')
        self.AFT.configure(width=780)

        self.btnewfile = tk.Button(self.Frame_middle)
        self.btnewfile.place(relx=0.2, rely=0.46, height=130, width=130)
        self.btnewfile.configure(activebackground="#d9d9d9")
        self.btnewfile.configure(activeforeground="#000000")
        self.btnewfile.configure(background="#ffffff")
        self.btnewfile.configure(disabledforeground="#a3a3a3")
        self.btnewfile.configure(foreground="#000000")
        self.btnewfile.configure(highlightbackground="#d9d9d9")
        self.btnewfile.configure(highlightcolor="black")
        self._img2 = tk.PhotoImage(file="C:/Users/Admin/Desktop/ranjith/v1.0/v4.0/images/Newlcon1.png")
        self.btnewfile.configure(image=self._img2)
        self.btnewfile.configure(pady="0")
        self.btnewfile.configure(text='''New project''', command=page1)
        self.btnewfile.configure(width=130)

        self.btopenfile = tk.Button(self.Frame_middle)
        self.btopenfile.place(relx=0.39, rely=0.46, height=130, width=130)
        self.btopenfile.configure(activebackground="#d9d9d9")
        self.btopenfile.configure(activeforeground="#000000")
        self.btopenfile.configure(background="#ffffff")
        self.btopenfile.configure(disabledforeground="#a3a3a3")
        self.btopenfile.configure(foreground="#000000")
        self.btopenfile.configure(highlightbackground="#d9d9d9")
        self.btopenfile.configure(highlightcolor="black")
        self._img3 = tk.PhotoImage(file="C:/Users/Admin/Desktop/ranjith/v1.0/v4.0/images/open-folder.png")
        self.btopenfile.configure(image=self._img3)
        self.btopenfile.configure(pady="0")
        self.btopenfile.configure(text='''Open Project''', command=OpenFile)

        self.bthelp = tk.Button(self.Frame_middle)
        self.bthelp.place(relx=0.59, rely=0.46, height=130, width=130)
        self.bthelp.configure(activebackground="#d9d9d9")
        self.bthelp.configure(activeforeground="#000000")
        self.bthelp.configure(background="#ffffff")
        self.bthelp.configure(disabledforeground="#a3a3a3")
        self.bthelp.configure(foreground="#000000")
        self.bthelp.configure(highlightbackground="#d9d9d9")
        self.bthelp.configure(highlightcolor="black")
        self._img4 = tk.PhotoImage(file="C:/Users/Admin/Desktop/ranjith/v1.0/v4.0/images/help.png")
        self.bthelp.configure(image=self._img4)
        self.bthelp.configure(pady="0")
        self.bthelp.configure(text='''Help''')
        self.bthelp.configure(width=140)
        # Frame Middle Ends
        

        
        

        


        






if __name__ == '__main__':
    vp_start_gui()