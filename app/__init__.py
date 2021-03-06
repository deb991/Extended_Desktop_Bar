__project__ = 'OSADOL.'
__author__ = 'Jeys_Ozzius.'
__descript__ = "Copyright (c) 2017 Anirban Ghosh(Anirban83314) & Debashis Biswas(deb991)."
__URL__ = 'https://github.com/Anirban83314/OutLook-Automation-All-scopes-with-APP-utility.'
__NB__ = 'For more information, please see github page & all commit details.'
__CipherSig__ = 'This project & all associate files are encrypted under PBEncryption cryptography. Another details will be available at the end of this pgogram.'

from tkinter import *
import tkinter.messagebox
import tkinter.font
import os
import sys
import subprocess


class MasterWidget:
    desktopButtomBar = Tk ( )

    w = 1595  # width for the Tk desktopButtomBar
    h = 40  # height for the Tk desktopButtomBar

    # get screen width and height
    ws = desktopButtomBar.winfo_screenwidth ( )  # width of the screen
    hs = desktopButtomBar.winfo_screenheight ( )  # height of the screen

    # calculate x and y coordinates for the Tk desktopButtomBar window
    x = (ws / 1) - (w / 1)
    y = (hs / 1) - (h / 1)

    desktopButtomBar.geometry ( '%dx%d+%d+%d' % (w , h , x , y) )

    desktopButtomBar.geometry ( "1595x40" )
    desktopButtomBar.resizable ( 0 , 0 )
    desktopButtomBar.configure ( background='grey' )
    desktopButtomBar.wm_attributes ( "-transparentcolor" , "grey" )
    desktopButtomBar.overrideredirect ( 1 )
    desktopButtomBar.lift ( )

    # helv9 = tkFont.Font(family='Helvetica', size=9, weight=tkFont.BOLD)
    helv9 = tkinter.font.Font ( family='Helvetica' , size=7 , weight='bold' )
    tkinter.font.families ( )
    ButtFontColor = '#00A7FB'

    def __init__(self):
        Tk.Tk.__init__ ( self )
        self.title ( "Handling WM_DELETE_WINDOW protocol" )
        self.geometry ( "500x300+500+200" )
        self.make_topmost ( )
        self.protocol ( "WM_DELETE_WINDOW" , self.on_exit )

    def on_Exit(self):
        """When you click to exit, this function is called"""
        if tkinter.messagebox.askyesno ( "Exit" , "Do you want to quit the application?" ):
            self.destroy ( )

    def center(self , desktopButtomBar):
        """Centers this Tk window"""
        self.eval ( 'tk::PlaceWindow %s center' % desktopButtomBar.winfo_pathname ( desktopButtomBar.winfo_id ( ) ) )

    def make_topmost(self):
        """Makes this window the topmost window"""
        self.lift ( )
        self.attributes ( "-topmost" , 1 )
        self.attributes ( "-topmost" , 1 )

    def Search(self):
        print ( 'Search' )

    b = Button ( desktopButtomBar , text="Search" , font=helv9 , command=Search , padx=8 , bg='black' ,
                 fg=ButtFontColor )
    b.place ( anchor=SW )
    b.pack ( side=LEFT )

    def OUTLOOK(self):
        print ( 'OUTLOOK' )
        os.chdir ( 'C:\\Program Files\\Microsoft Office\\root\\Office16\\' )
        os.startfile ( r"C:\\Program Files\\Microsoft Office\\root\\Office16\\OUTLOOK.EXE" )

    b = Button ( desktopButtomBar , text="OutLook" , font=helv9 , command=OUTLOOK , padx=6 , bg="black" ,
                 fg=ButtFontColor )
    b.place ( anchor=SW )
    b.pack ( side=LEFT )

    def MSSequalServer(self):
        print ( 'Initiating SEQUAL MANAGEMENT STUDIO' )
        os.chdir ( "C:\\Program Files\\Microsoft SQL Server\\100\\Tools\\Binn\\VSShell\\Common7\\IDE" )
        os.startfile ( r"C:\Program Files\Microsoft SQL Server\100\Tools\Binn\VSShell\Common7\IDE\\Ssms.exe" )

    b = Button ( desktopButtomBar , text="Sequal Server" , font=helv9 , command=MSSequalServer , padx=6 , bg="black" ,
                 fg=ButtFontColor )
    b.place ( anchor=SW )
    b.pack ( side=LEFT )

    def PLSql(self):
        print ( 'Initiating SQL developer' )
        os.chdir ( "C:\\Users\\Debashis.Biswas\\Software\\sqldeveloper" )
        os.startfile ( r"C:\\Users\\Debashis.Biswas\\Software\\sqldeveloper\\sqldeveloper.exe" )

    b = Button ( desktopButtomBar , text="PL/ SQL" , font=helv9 , command=PLSql , padx=6 , bg="black" ,
                 fg=ButtFontColor )
    b.place ( anchor=SW )
    b.pack ( side=LEFT )

    def BMCControlM(self):
        print ( 'service start DDF' )

    b = Button ( desktopButtomBar , text="Control M" , font=helv9 , command=BMCControlM , padx=6 , bg="black" ,
                 fg=ButtFontColor )
    b.place ( anchor=SW )
    b.pack ( side=LEFT )

    def PyCharmIDE(self):
        print ( 'service start DDF' )

    b = Button ( desktopButtomBar , text="PyCharm" , font=helv9 , command=PyCharmIDE , padx=6 , bg="black" ,
                 fg=ButtFontColor )
    b.place ( anchor=SW )
    b.pack ( side=LEFT )

    def Visual_BasicIDE(self):
        print ( 'service start DDF' )

    b = Button ( desktopButtomBar , text="Visual Basic" , font=helv9 , command=Visual_BasicIDE , padx=6 , bg="black" ,
                 fg=ButtFontColor )
    b.place ( anchor=SW )
    b.pack ( side=LEFT )

    def File_Manager(self):
        print ( 'File manager' )

    b = Button ( desktopButtomBar , text="File Manager" , font=helv9 , command=File_Manager , padx=6 , bg="black" ,
                 fg=ButtFontColor )
    b.place ( anchor=SW )
    b.pack ( side=LEFT )

    def QUERY_LIBRARY(self):
        print ( 'Query Library' )
        os.startfile ( r"C:\Program Files\Microsoft SQL Server\100\Tools\Binn\VSShell\Common7\IDE\\Ssms.exe" )

    b = Button ( desktopButtomBar , text="QUEERY LIBRARY" , font=helv9 , command=QUERY_LIBRARY , padx=6 , bg="black" ,
                 fg=ButtFontColor )
    b.place ( anchor=SW )
    b.pack ( side=LEFT )

    desktopRightTaskBar = Tk ( )

    w = 1595  # width for the Tk desktopTaskBar
    h = 40  # height for the Tk desktopTaskBar

    # get screen width and height
    ws = desktopRightTaskBar.winfo_screenwidth ( )  # width of the screen
    hs = desktopRightTaskBar.winfo_screenheight ( )  # height of the screen

    # calculate x and y coordinates for the Tk desktopTaskBar window
    x = (ws / 5) - (w / 5)
    y = (hs / 5) - (h / 5)

    desktopRightTaskBar.geometry ( '%dx%d+%d+%d' % (w , h , x , y) )

    desktopRightTaskBar.geometry ( "1595x40" )
    desktopRightTaskBar.resizable ( 0 , 0 )
    desktopRightTaskBar.configure ( background='grey' )
    desktopRightTaskBar.wm_attributes ( "-transparentcolor" , "grey" )
    desktopRightTaskBar.overrideredirect ( 1 )
    desktopRightTaskBar.lift ( )

    # helv9 = tkFont.Font(family='Helvetica', size=9, weight=tkFont.BOLD)
    helv9 = tkinter.font.Font ( family='Helvetica' , size=7 , weight='bold' )
    tkinter.font.families ( )
    ButtFontColor = '#00A7FB'

    def __init__(self):
        Tk.Tk.__init__ ( self )
        self.title ( "Handling WM_DELETE_WINDOW protocol" )
        self.geometry ( "500x300+500+200" )
        self.make_topmost ( )
        self.protocol ( "WM_DELETE_WINDOW" , self.on_exit )

    def on_Exit(self):
        """When you click to exit, this function is called"""
        if tkinter.messagebox.askyesno ( "Exit" , "Do you want to quit the application?" ):
            self.destroy ( )

    def Search(self):
        print ( 'Search' )

    b = Button ( desktopRightTaskBar , text="Search" , font=helv9 , command=Search , padx=8 , bg='black' ,
                 fg=ButtFontColor )
    b.place ( anchor=SW )
    b.pack ( side=LEFT )

    def OUTLOOK(self):
        print ( 'OUTLOOK' )
        os.chdir ( 'C:\\Program Files\\Microsoft Office\\root\\Office16\\' )
        os.startfile ( r"C:\\Program Files\\Microsoft Office\\root\\Office16\\OUTLOOK.EXE" )

    b = Button ( desktopRightTaskBar , text="OutLook" , font=helv9 , command=OUTLOOK , padx=6 , bg="black" ,
                 fg=ButtFontColor )
    b.place ( anchor=SW )
    b.pack ( side=LEFT )

    def MSSequalServer(self):
        print ( 'Initiating SEQUAL MANAGEMENT STUDIO' )
        os.chdir ( "C:\\Program Files\\Microsoft SQL Server\\100\\Tools\\Binn\\VSShell\\Common7\\IDE" )
        os.startfile ( r"C:\Program Files\Microsoft SQL Server\100\Tools\Binn\VSShell\Common7\IDE\\Ssms.exe" )

    b = Button ( desktopRightTaskBar , text="Sequal Server" , font=helv9 , command=MSSequalServer , padx=6 ,
                 bg="black" ,
                 fg=ButtFontColor )
    b.place ( anchor=SW )
    b.pack ( side=LEFT )

    def PLSql(self):
        print ( 'Initiating SQL developer' )
        os.chdir ( "C:\\Users\\Debashis.Biswas\\Software\\sqldeveloper" )
        os.startfile ( r"C:\\Users\\Debashis.Biswas\\Software\\sqldeveloper\\sqldeveloper.exe" )

    b = Button ( desktopRightTaskBar , text="PL/ SQL" , font=helv9 , command=PLSql , padx=6 , bg="black" ,
                 fg=ButtFontColor )
    b.place ( anchor=SW )
    b.pack ( side=LEFT )

    def BMCControlM(self):
        print ( 'service start DDF' )

    b = Button ( desktopRightTaskBar , text="Control M" , font=helv9 , command=BMCControlM , padx=6 , bg="black" ,
                 fg=ButtFontColor )
    b.place ( anchor=SW )
    b.pack ( side=LEFT )

    def PyCharmIDE(self):
        print ( 'service start DDF' )

    b = Button ( desktopRightTaskBar , text="PyCharm" , font=helv9 , command=PyCharmIDE , padx=6 , bg="black" ,
                 fg=ButtFontColor )
    b.place ( anchor=SW )
    b.pack ( side=LEFT )

    def Visual_BasicIDE(self):
        print ( 'service start DDF' )

    b = Button ( desktopRightTaskBar , text="Visual Basic" , font=helv9 , command=Visual_BasicIDE , padx=6 ,
                 bg="black" ,
                 fg=ButtFontColor )
    b.place ( anchor=SW )
    b.pack ( side=LEFT )



    def File_Manager(self):
        print ( 'File manager' )

    b = Button ( desktopRightTaskBar , text="File Manager" , font=helv9 , command=File_Manager , padx=6 , bg="black" ,
                 fg=ButtFontColor )
    b.place ( anchor=SW )
    b.pack ( side=LEFT )

    def QUERY_LIBRARY(self):
        print ( 'Query Library' )
        os.startfile ( r"C:\Program Files\Microsoft SQL Server\100\Tools\Binn\VSShell\Common7\IDE\\Ssms.exe" )

    b = Button ( desktopRightTaskBar , text="QUEERY LIBRARY" , font=helv9 , command=QUERY_LIBRARY , padx=6 ,
                 bg="black" ,
                 fg=ButtFontColor )
    b.place ( anchor=SW )
    b.pack ( side=LEFT )


mainloop ( )
