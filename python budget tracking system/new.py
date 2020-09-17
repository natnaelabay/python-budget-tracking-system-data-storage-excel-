from tkinter import *
from tkinter import messagebox
import datetime
import win32com.client
import os,os.path

class Home:
    def __init__(self,root):
        self.root = root
        self.root.title("Budget Tracking System for AAiT")
        self.root.iconbitmap("newlogo.ico")
        self.root.geometry('590x600')
        self.logoImage = PhotoImage(file = "logo.gif")
        #self.Login_frame = Frame(self.root, width = 1000, height = 1000) #Main Frame of the login
        #self.Login_frame.place( x = 100, y = 50)
        self.login_image = Frame(self.root)
        self.login_image.place(x = 240, y = 80)
        self.login_image_place = Label(self.login_image,image = self.logoImage)
        self.login_image_place.image = self.logoImage
        self.login_image_place.grid(row = 0 , column = 2)
        self.login = Frame(self.root,width = 800, height = 200)
        self.login.place(x = 130 , y = 240)
        self.login.label = Label(self.login,text = "AAiT Budget Tracking system ",font = ('Times New Roman',18,'bold'))
        self.login.label.grid(row = 0,columnspan= 6,padx = 0)
        self.user_ask = Label(self.login,text = "Username: ",font = ('Times New Roman',12))
        self.password_ask = Label(self.login,text = "Password: ",font = ('Times New Roman',12))
        self.user_ask.grid(row = 1,column=0, stick = W,pady = 40,padx = 30)
        self.password_ask.grid(row = 2, column=0, stick = W,padx = 30)
        self.user_name_Entry = Entry(self.login)
        self.user_password_Entry = Entry(self.login,show = "*")
        self.user_name_Entry.grid(row = 1,column = 1)
        self.user_password_Entry.grid(row = 2,column = 1)
        self.Check_value = IntVar()
        self.variable = StringVar()
        self.varframe = Label(self.login,textvariable = self.variable)
        self.LoginButton = Button(self.login,text = "Login",command = self.ConfirmPassword)
        self.LoginButton.grid(row = 5,columnspan = 3, pady = 35)
        self.user_name = self.user_name_Entry.get()
        self.user_password = self.user_password_Entry.get()
        self.Check_value = self.Check_value.get()
        self.homepageicon = PhotoImage(file = 'illumnati.png')
        self.homepagelabel = Label (self.root,image = self.homepageicon)
        self.homepagelabel.place(x = 440,y = 450)

    def ConfirmPassword(self):
        #New = Manager("Managing Director")
        loc = ("Budget.xlsm")
        self.workbookxlrd = xlrd.open_workbook("Budget.xlsm")
        self.sheetxlrd = self.workbookxlrd.sheet_by_name("Password")

        self.user_name = self.user_name_Entry.get()
        self.user_password = self.user_password_Entry.get()
        #self.sheet = self.wb[self.Sheetname]
        allow = False
        for i in range(2,8 + 1):
            #self.sheetxlrd.cell_value(i - 1,1)
            if self.sheetxlrd.cell_value(i - 1,1) == self.user_name:
                #print(self.sheet.cell(i,2).value)
                #print(self.sheet.cell(i,3).value)
                #print(self.sheet.cell(i,3).value == eval(self.user_password))
                if self.sheetxlrd.cell_value(i - 1,2) == eval(self.user_password):
                    #print(self.sheet.cell(i,3).value)
                    if self.sheetxlrd.cell_value(i - 1,3) == "Activate":                    
                        allow = True
                        if self.user_name == "IT":
                            Nextpage = ICT("ICT")
                        elif self.user_name == "aaitadmin":
                            Nextpage = Manager("AAiT Admin")
                        elif self.user_name == "aauadmin":
                            Nextpage = AAU("AAU Admin")
                        else:
                            Nextpage = Department(self.sheetxlrd.cell_value(i - 1,0))
                        self.login.destroy()
                        self.login_image.destroy()
                        self.homepagelabel.destroy()
                        self.root.geometry('1000x800')
                        #self.wb.save("Budget.xlsm")
        if not allow:
            self.varframe.grid(row = 4,columnspan = 4)
            self.variable.set("Invalid Credentials")
            #self.wb.save("Budget.xlsm")
            
    def writexcel(self,sheetname,rowinp,columninp,valueinp):
        self.wb = openpyxl.load_workbook("Budget.xlsm", keep_vba = True)
        self.sheet = sheetname
        self.sheet.cell(row = rowinp, column = columninp ).value = valueinp
        self.wb.save("Budget.xlsm")

    def refresh(self):
        file_reader = open("Budget.xlsm","r")
        file_reader.close()
    #def refresh(self):
    #    f = os.getcwd()
    #    xl = win32com.client.DispatchEx("Excel.application")
    #    excel_path = os.path.expanduser("Budget.xlsm")
        #workbook = excel_macro.Workbooks.Open(c + '/' + "Budget.xlsm")
        #wb.visible = True
    #    excel_path.RefreshAll()
    #    excel_path.Save()

class Department(Home):
    def __init__(self,job):
        self.root = root
        self.job = job
        self.jobonly = self.job
        
        if self.job =='itsc':
            self.job = "Department of " + self.job.upper()
        elif self.job == 'IT':
            self.job = "Super User"
            pass
        elif self.job == 'AAiT Admin':
            self.job = "AAiT Admin"
        elif self.job == 'AAU Admin':
            self.job = "AAU Admin"
        else:
             self.job = "Department of " + self.job + " Engineering" 
        
        self.actions = Frame(root, width = 200, height = 800, bg = "#343a40")
        self.actions.pack(side = LEFT)

        self.tops = Frame(root, width = 1350, height = 120)#, bg = "#E0E0E0"
        self.tops.pack(side = TOP)

        self.intro = Label(self.tops,text = self.job, font = ("Times New Roman",25) )
        self.intro.grid(row = 0,column = 1, sticky = "E")

        self.logoImage = PhotoImage(file = "logo.gif")
        self.login_image_place = Label(self.tops,image =self.logoImage)
        self.login_image_place.image = self.logoImage
        self.login_image_place.grid(row = 0 ,column = 0)
    
        self.main = Frame(root, width = 800, height = 630, bg = "#00CCCC")
        self.main.place(relx = 0.2,rely =0.22)

        self.HomeButton = Button(self.actions, text = "Signout" , command = self.ReturnHome)
        self.Purchase = Button(self.actions, text = "Withdrawal", command = self.purchase)
        self.Distribute = Button(self.actions, text = "Distribute", command = self.distribute)
        self.Related = Button(self.actions, text = "Additional Request", command = self.Budgetrelated)
        self.Report = Button(self.actions, text = "Expenditure report",command = self.ReportDisplay)
        self.BudgetCode = Button(self.actions,text = "Budget Code Help", command = self.BudgetcodeFinder)#,command = pass)
        #self.AnnualReport = Button(self.actions,text = "AnnualReport", command = self.ReportDisplay)#,command = pass))


        self.HomeButton.place(x = 40 , y = 30, width = "130", height = "40")
        self.Related.place(x =40, y = 220,width = '130',height = '40') 
        self.Distribute.place(x = 40 , y = 280, width = "130", height = "40")
        self.Purchase.place(x = 40, y = 340, width = "130", height = "40")
        self.Report.place(x = 40, y = 400, width = "130", height = "40")
        self.BudgetCode.place(x = 40, y = 460, width = "130", height = "40")
        #self.AnnualReport.place(x = 40, y = 520, width = "130", height = "40")

        self.SearchMain = Frame(self.main,width = 800, height = 800,bg = "#CCFFFF")
        self.SearchMain.place(y = 0)

        self.welcometext = Label (self.root,text = "Welcome",font=("Courier", 44),bg = "#CCFFFF")
        self.welcometext.place(x = 500 ,y = 400)
      

    def ReturnHome(self):
        self.actions.destroy()
        try:
            self.SearchMain.destroy()
        except:
            pass

        try:
            self.main.destroy()
        except:
            pass
        try:
            self.title.destroy()
        except:
            pass
        try:
            self.welcometext.destroy()
        except:
            pass
        #self.SearchMain.destroy()
        self.main.destroy()
        self.tops.destroy()
        home = Home(self.root)

    def distribute(self):
        try:
            self.SubmitButton.destroy()
            self.submit.destroy()
            self.SearchMain.destroy()
        except AttributeError:
            self.SearchMain.destroy()
        try:
            self.welcometext.destroy()
        except:
            pass
        try:
            self.AnnualBudget.destroy()
            self.AdjustedBudget.destroy()
            self.Transfer.destroy()
            self.AdjustedBudget.destroy()
            self.Reason.destroy()
            self.Transfer.destroy()
            self.TransferEntry.destroy()
            self.Reasonraise.destroy()
            self.SubmitButton.destroy()
        except:
            pass
        try:
            self.Budget1.destroy()
            self.Budget2.destroy()
            self.Budget3.destroy()
            self.Budget4.destroy()
            self.Budget5.destroy()
            self.Budget6.destroy()
            self.Budget7.destroy()
            self.Budget8.destroy()
            self.Budget9.destroy()
            self.Budget10.destroy()
            self.Budget1Entry.destroy()
            self.Budget2Entry.destroy()
            self.Budget3Entry.destroy()
            self.Budget4Entry.destroy()
            self.Budget5Entry.destroy()
            self.Budget6Entry.destroy()
            self.Budget7Entry.destroy()
            self.Budget8Entry.destroy()
            self.Budget9Entry.destroy()
            self.Budget10Entry.destroy()
            self.SubmitButton.destroy()
            self.title2.destroy()
        except:
            pass
        try:
            self.moneyiconlabel.destroy()
        except:
            pass
        try:
            self.title1.destroy()
        except:
            pass
        try:
            self.title1.destroy()
        except:
            pass
        try:
            self.title2.destroy()
        except:
            pass
        try:
            self.title3.destroy()
        except:
            pass
        try:
            self.title4.destroy()
        except:
            pass
        
        self.SearchMain = Frame(self.main,width = 800, height = 800,bg = "#CCFFFF")
        self.SearchMain.place( y = 50)

        self.title2 = Label(self.main, text = "Distrubute the Annual Budget for each Budget code in percent(%)",font=("Times New Roman", 18) ,bg = "#00CCCC")
        self.title2.place(x = 35, y = 10)

        self.Annual_Budget = 0

        loc = ("Budget.xlsm")
        self.workbookxlrd = xlrd.open_workbook(loc)
        self.sheetxlrd = self.workbookxlrd.sheet_by_name(self.jobonly)

        self.AnnualBudget = Label(self.SearchMain,text = "Adjusted Budget: ",font = ('Times New Roman',12),bg = "#CCFFFF")
        self.AnnualBudgetVariable = StringVar(self.SearchMain)
        self.AnnualBudgetVariable = self.sheetxlrd.cell_value(18,1)
        self.AnnualBudgetValue = Label(self.SearchMain, text = self.AnnualBudgetVariable,font = ('Times New Roman',12),bg = "#CCFFFF")
        self.AnnualBudget.place(x = 450, y = 50)
        self.AnnualBudgetValue.place(x = 600, y = 50)

        self.wb = openpyxl.load_workbook("Budget.xlsm", keep_vba = True)

        
        #self.fill_out = Frame(self.SearchMain, )
        self.Budget1 = Label(self.SearchMain,text = "6111: ",font = ('Times New Roman',13),bg = "#CCFFFF")
        self.Budget1Entry = Entry(self.SearchMain)
        self.Budget1.place(x = 90, y = 50)
        self.Budget1Entry.place(x = 250, y = 50)
        
        self.Budget2 = Label(self.SearchMain,text = "6113: ",font = ('Times New Roman',13),bg = "#CCFFFF")
        self.Budget2Entry = Entry(self.SearchMain)
        self.Budget2.place(x = 90, y = 80)
        self.Budget2Entry.place(x = 250, y = 80)

        self.Budget3 = Label(self.SearchMain,text = "6121: ",font = ('Times New Roman',13),bg = "#CCFFFF")
        self.Budget3Entry = Entry(self.SearchMain)
        self.Budget3.place(x = 90, y = 110)
        self.Budget3Entry.place(x = 250, y = 110)

        self.Budget4 = Label(self.SearchMain,text = "6212: ",font = ('Times New Roman',13),bg = "#CCFFFF")
        self.Budget4Entry = Entry(self.SearchMain)
        self.Budget4.place(x = 90, y = 140)
        self.Budget4Entry.place(x = 250, y = 140)
        

        self.Budget5 = Label(self.SearchMain,text = "6211: ",font = ('Times New Roman',13),bg = "#CCFFFF")
        self.Budget5Entry = Entry(self.SearchMain)
        self.Budget5.place(x = 90, y = 170)
        self.Budget5Entry.place(x = 250, y = 170)

        self.Budget6 = Label(self.SearchMain,text = "6218: ",font = ('Times New Roman',13),bg = "#CCFFFF")
        self.Budget6Entry = Entry(self.SearchMain)
        self.Budget6.place(x = 90, y = 200)
        self.Budget6Entry.place(x = 250, y = 200)

        self.Budget7 = Label(self.SearchMain,text = "6233: ",font = ('Times New Roman',13),bg = "#CCFFFF")
        self.Budget7Entry = Entry(self.SearchMain)
        self.Budget7.place(x = 90, y = 230)
        self.Budget7Entry.place(x = 250, y = 230)

        self.Budget8 = Label(self.SearchMain,text = "6240: ",font = ('Times New Roman',13),bg = "#CCFFFF")
        self.Budget8Entry = Entry(self.SearchMain)
        self.Budget8.place(x = 90, y = 260)
        self.Budget8Entry.place(x = 250, y = 260)

        self.Budget9 = Label(self.SearchMain,text = "6256 ",font = ('Times New Roman',13),bg = "#CCFFFF")
        self.Budget9Entry = Entry(self.SearchMain)
        self.Budget9.place(x = 90, y = 290)
        self.Budget9Entry.place(x = 250, y = 290)

        self.Budget10 = Label(self.SearchMain,text = "6271: ",font = ('Times New Roman',13),bg = "#CCFFFF")
        self.Budget10Entry = Entry(self.SearchMain)
        self.Budget10.place(x = 90, y = 320)
        self.Budget10Entry.place(x = 250, y = 320)

        def distributes():
            self.sheet = self.wb[self.jobonly]
            self.sheet.cell(3,2).value = eval(self.Budget1Entry.get())
            self.sheet.cell(4,2).value = eval(self.Budget2Entry.get())
            self.sheet.cell(5,2).value = eval(self.Budget3Entry.get())
            self.sheet.cell(6,2).value = eval(self.Budget4Entry.get())
            self.sheet.cell(7,2).value = eval(self.Budget5Entry.get())
            self.sheet.cell(8,2).value = eval(self.Budget6Entry.get())
            self.sheet.cell(9,2).value = eval(self.Budget7Entry.get())
            self.sheet.cell(10,2).value = eval(self.Budget8Entry.get())
            self.sheet.cell(11,2).value = eval(self.Budget9Entry.get())
            self.sheet.cell(12,2).value = eval(self.Budget10Entry.get())
            self.wb.save("Budget.xlsm")                

        self.SubmitButton = Button(text = "Submit",command = distributes)
        self.SubmitButton.place(x = 700, y = 600,width = "130", height = "40")

    def ReportDisplay(self):
        try:
            self.welcometext.destroy()
        except:
            pass
        try:
            self.SubmitButton.destroy()
            self.submit.destroy()
            self.SearchMain.destroy()
        except AttributeError:
            self.SearchMain.destroy()
        try:
            self.AnnualBudget.destroy()
            self.AdjustedBudget.destroy()
            self.Transfer.destroy()
            self.AdjustedBudget.destroy()
            self.Reason.destroy()
            self.Transfer.destroy()
            self.TransferEntry.destroy()
            self.Reasonraise.destroy()
            self.SubmitButton.destroy()
        except:
            pass
        try:
            self.Budget1.destroy()
            self.Budget2.destroy()
            self.Budget3.destroy()
            self.Budget4.destroy()
            self.Budget5.destroy()
            self.Budget6.destroy()
            self.Budget7.destroy()
            self.Budget8.destroy()
            self.Budget9.destroy()
            self.Budget10.destroy()
            self.Budget1Entry.destroy()
            self.Budget2Entry.destroy()
            self.Budget3Entry.destroy()
            self.Budget4Entry.destroy()
            self.Budget5Entry.destroy()
            self.Budget6Entry.destroy()
            self.Budget7Entry.destroy()
            self.Budget8Entry.destroy()
            self.Budget9Entry.destroy()
            self.Budget10Entry.destroy()
            self.SubmitButton.destroy()
            self.title2.destroy()
        except:
            pass
        try:
            self.title1.destroy()
        except:
            pass
        try:
            self.title2.destroy()
        except:
            pass
        try:
            self.title3.destroy()
        except:
            pass
        try:
            self.title4.destroy()
        except:
            pass
                
        self.title1 = Label(self.main, text = "Expenditure report",font=("Times New Roman", 18) ,bg = "#00CCCC")
        self.title1.place(x = 250, y = 10)
        self.SearchMain = Frame(self.main,width = 800, height = 800,bg = "#CCFFFF")
        self.SearchMain.place( x = 50 ,y = 50)
        self.wb = openpyxl.load_workbook("Budget.xlsm", keep_vba = True)
        self.sheet = self.wb[self.jobonly + "Report"]
        for i in range(1,self.sheet.max_row + 1):
            #print(self.sheet.cell(row = i, column = 4).value)
            for j in range(1,5):
                Label(self.SearchMain,text = self.sheet.cell(row = i, column = j).value,font=("Times New Roman", 13) ,bg = "#CCFFFF").grid(row = i + 2,column = j + 10, pady = 20, padx = 25)
        #self.wb.save("Budget.xlsm")                
        self.Scrollbar = Scrollbar(self.SearchMain) 
      
      
    def Budgetrelated(self):
        try:
            self.welcometext.destroy()
        except:
            pass
        try:
            self.SubmitButton.destroy()
            self.submit.destroy()
            self.SearchMain.destroy()
        except AttributeError:
            self.SearchMain.destroy()
        try:
            self.AnnualBudget.destroy()
            self.AdjustedBudget.destroy()
            self.Transfer.destroy()
            self.AdjustedBudget.destroy()
            self.Reason.destroy()
            self.Transfer.destroy()
            self.TransferEntry.destroy()
            self.Reasonraise.destroy()
            self.SubmitButton.destroy()
        except:
            pass
        try:
            self.Budget1.destroy()
            self.Budget2.destroy()
            self.Budget3.destroy()
            self.Budget4.destroy()
            self.Budget5.destroy()
            self.Budget6.destroy()
            self.Budget7.destroy()
            self.Budget8.destroy()
            self.Budget9.destroy()
            self.Budget10.destroy()
            self.Budget1Entry.destroy()
            self.Budget2Entry.destroy()
            self.Budget3Entry.destroy()
            self.Budget4Entry.destroy()
            self.Budget5Entry.destroy()
            self.Budget6Entry.destroy()
            self.Budget7Entry.destroy()
            self.Budget8Entry.destroy()
            self.Budget9Entry.destroy()
            self.Budget10Entry.destroy()
            self.SubmitButton.destroy()
            self.title2.destroy()
        except:
            pass
        try:
            self.title1.destroy()
        except:
            pass
        try:
            self.title2.destroy()
        except:
            pass
        try:
            self.title3.destroy()
        except:
            pass
        try:
            self.title4.destroy()
        except:
            pass

        #self.refresh()


        
            
        self.title3 = Label(self.main, text = "Additional Budget form",font=("Times New Roman", 18) ,bg = "#00CCCC")
        self.title3.place(x = 290, y = 10)
        
        self.SearchMain = Frame(self.main,width = 800, height = 800,bg = "#CCFFFF")
        self.SearchMain.place( y = 50)

        self.Annual_Budget = 0

        loc = ("Budget.xlsm")
        self.workbookxlrd = xlrd.open_workbook(loc)
        self.sheetxlrd = self.workbookxlrd.sheet_by_name(self.jobonly)

        self.AnnualBudget = Label(self.SearchMain,text = "Annual Budget: ",bg = "#CCFFFF",font = ('Roboto',16))
        #self.AnnualBudgetVariable = StringVar(self.SearchMain)
        self.AnnualBudgetVariable = self.sheetxlrd.cell_value(15,1)
        self.AnnualBudgetValue = Label(self.SearchMain, text = self.AnnualBudgetVariable,bg = "#CCFFFF",font = ('Times New Roman',15))
        self.AnnualBudget.place(x = 50, y = 60)
        self.AnnualBudgetValue.place(x = 220, y = 62)

        self.AdjustedBudget = Label(self.SearchMain,text = "Adjusted Budget: ",bg = "#CCFFFF",font = ('Roboto',16))
        #self.AdjustedBudgetVariable = StringVar(self.SearchMain)
        self.AdjustedBudgetVariable = self.sheetxlrd.cell_value(18,1)
        self.AdjustedBudgetValue = Label(self.SearchMain, text = self.AdjustedBudgetVariable,bg = "#CCFFFF",font = ('Times New Roman',15))
        self.AdjustedBudget.place(x = 350, y = 60)
        self.AdjustedBudgetValue.place(x = 540, y = 62)

##        self.Report = Label(self.SearchMain, text = "Ask for an increase of Annual Budget",font=("Courier", 15), bg = "#00CCCC" )
##        self.Report.place(x = 150, y = 200)

        self.Transfer = Label(self.SearchMain, text = "Amount of Transfer: ",font = ('Times New Roman',13),bg = "#CCFFFF")
        self.TransferEntry = Entry(self.SearchMain)
        self.Transfer.place(x = 50 , y = 150)
        self.TransferEntry.place(x = 220 , y = 150)

        self.Reason = Label(self.SearchMain, text = "Reason for enquiry: ",font = ('Times New Roman',13),bg = "#CCFFFF")
        self.Reason.place(x = 50, y = 200)
        self.Reasonraise = Text(self.SearchMain,height = 7, width = 40)
        self.Reasonraise.place(x = 220, y = 200)

        def AskAddBudget():
            self.wb = openpyxl.load_workbook("Budget.xlsm", keep_vba = True)
            self.sheet = self.wb[self.jobonly]
            self.Update = self.TransferEntry.get()
            self.sheet.cell(17,2).value = eval(self.Update)

            self.sheet.cell(18,2).value = self.Reasonraise.get("1.0",END)
            self.wb.save("Budget.xlsm")
            file_reader = open("Budget.xlsm","r")
            file_reader.close() 

                
        self.SubmitButton = Button(self.SearchMain, text = "Submit",command = AskAddBudget)
        self.SubmitButton.place(x = 250, y = 400,width = "150", height = "50")


    def BudgetcodeFinder(self):
        try:
            self.welcometext.destroy()
        except:
            pass
        try:
            self.SubmitButton.destroy()
            self.submit.destroy()
            self.SearchMain.destroy()
        except AttributeError:
            self.SearchMain.destroy()
        try:
            self.AnnualBudget.destroy()
            self.AdjustedBudget.destroy()
            self.Transfer.destroy()
            self.AdjustedBudget.destroy()
            self.Reason.destroy()
            self.Transfer.destroy()
            self.TransferEntry.destroy()
            self.Reasonraise.destroy()
            self.SubmitButton.destroy()
        except:
            pass
        try:
            self.Budget1.destroy()
            self.Budget2.destroy()
            self.Budget3.destroy()
            self.Budget4.destroy()
            self.Budget5.destroy()
            self.Budget6.destroy()
            self.Budget7.destroy()
            self.Budget8.destroy()
            self.Budget9.destroy()
            self.Budget10.destroy()
            self.Budget1Entry.destroy()
            self.Budget2Entry.destroy()
            self.Budget3Entry.destroy()
            self.Budget4Entry.destroy()
            self.Budget5Entry.destroy()
            self.Budget6Entry.destroy()
            self.Budget7Entry.destroy()
            self.Budget8Entry.destroy()
            self.Budget9Entry.destroy()
            self.Budget10Entry.destroy()
            self.SubmitButton.destroy()
            self.title2.destroy()
        except:
            pass
        try:
            self.title1.destroy()
        except:
            pass
        try:
            self.title2.destroy()
        except:
            pass
        try:
            self.title3.destroy()
        except:
            pass
        try:
            self.title4.destroy()
        except:
            pass
##        try:
##            self.SearchMain.destroy()
##        except:
##            pass
        
        self.title4 = Label(self.main, text = "Budget code help",font=("Times New Roman", 18) ,bg = "#00CCCC")
        self.title4.place(x = 290, y = 10)
        
        self.SearchMain = Frame(self.main,width = 800, height = 800,bg = "#CCFFFF")
        self.SearchMain.place( y = 50)
        
        self.BudgetCodeName = StringVar()
        self.SearchLabel = Label(self.SearchMain,text = "Write your budget code",font=("Times New Roman", 13) ,bg = "#CCFFFF")
        self.SearchLabel.place(x =92 , y = 60 )

        self.BudgetCodeEntry = Entry(self.SearchMain)
        self.BudgetCodeEntry.place(x = 280, y = 60)

        self.BudetCodeDescription = Text(self.main,height = 10, width = 70)
        self.BudetCodeDescription.place(x = 60, y = 150)

        self.sheet = self.wb["BudgetCode"]

        def BudgetCodeName():
            self.BudetCodeDescription.delete("1.0",END)
            code_hit = False
            for i in range(3,self.sheet.max_row + 1):
                if self.sheet.cell(i,1).value == eval(self.BudgetCodeEntry.get()):
                    code_hit = True
                    self.BudetCodeDescription.insert(INSERT, self.sheet.cell(i,2).value)
            if not code_hit:
                self.BudetCodeDescription.insert(INSERT, "No budget with this code!")
                    
        self.BudgetCodeButton = Button(self.SearchMain, text = "Search", width = 8, command = BudgetCodeName)
        self.BudgetCodeButton.place(x = 450, y = 60)        

    def purchase(self):
        try:
            self.welcometext.destroy()
        except:
            pass
        
        try:
            self.SubmitButton.destroy()
            self.submit.destroy()
            self.SearchMain.destroy()
        except AttributeError:
            self.SearchMain.destroy()
        try:
            self.AnnualBudget.destroy()
            self.AdjustedBudget.destroy()
            self.Transfer.destroy()
            self.AdjustedBudget.destroy()
            self.Reason.destroy()
            self.Transfer.destroy()
            self.TransferEntry.destroy()
            self.Reasonraise.destroy()
            self.SubmitButton.destroy()
        except:
            pass
        try:
            self.Budget1.destroy()
            self.Budget2.destroy()
            self.Budget3.destroy()
            self.Budget4.destroy()
            self.Budget5.destroy()
            self.Budget6.destroy()
            self.Budget7.destroy()
            self.Budget8.destroy()
            self.Budget9.destroy()
            self.Budget10.destroy()
            self.Budget1Entry.destroy()
            self.Budget2Entry.destroy()
            self.Budget3Entry.destroy()
            self.Budget4Entry.destroy()
            self.Budget5Entry.destroy()
            self.Budget6Entry.destroy()
            self.Budget7Entry.destroy()
            self.Budget8Entry.destroy()
            self.Budget9Entry.destroy()
            self.Budget10Entry.destroy()
            self.SubmitButton.destroy()
            self.title2.destroy()
        except:
            pass
        try:
            self.moneyiconlabel.destroy()
        except:
            pass
        try:
            self.title1.destroy()
        except:
            pass
        try:
            self.title1.destroy()
        except:
            pass
        try:
            self.title2.destroy()
        except:
            pass
        try:
            self.title3.destroy()
        except:
            pass
   
        try:
            self.title4.destroy()
        except:
            pass
        
       
        try:
            self.moneyiconlabel.destroy()
        except:
            pass
        try:
            dropdown.destroy()
        except:
            pass
        
        self.SearchMain = Frame(self.main,width = 800, height = 800,bg = "#CCFFFF")
        self.SearchMain.place( y = 50)
        self.title1 = Label(self.main, text = "Money withdrawal form",font=("Times New Roman", 18) ,bg = "#00CCCC")
        self.title1.place(x = 305, y = 10)
        
        #self.fill_out = Frame(self.SearchMain, )
        self.BudgetCode = Label(self.SearchMain,text = "BudgetCode: ",bg = "#CCFFFF")
        #self.BudgetCodeEntry = Entry(self.SearchMain)
        self.BudgetCode.place(x = 90, y = 50)
        #self.BudgetCodeEntry.place(x = 280, y = 150)

        self.variable = StringVar(self.SearchMain)

        loc = ("Budget.xlsm")
        self.workbookxlrd = xlrd.open_workbook(loc)
        self.sheetxlrd = self.workbookxlrd.sheet_by_name(self.jobonly)
        self.varlist = [int(self.sheetxlrd.cell_value(i,0)) for i in range(2,11)]
        self.variable.set("Select here")

        Dropdown = OptionMenu(self.SearchMain, self.variable, *self.varlist)
        Dropdown.place(x = 200 , y = 50)


        self.time = datetime.datetime.now()
        self.time = self.time.strftime("%Y-%m-%d %H:%M:%S")

        self.Date = Label(self.SearchMain, text = "Date",bg = "#CCFFFF")
        self.DateEntry = Label(self.SearchMain,text = self.time)
        self.Date.place(x = 90, y = 120)
        self.DateEntry.place(x = 200, y = 120)

        self.Bought = Label(self.SearchMain, text = "Reason",bg = "#CCFFFF")
        self.Bought.place(x = 90 , y = 180)

        self.BudetCodeDescription = Text(self.SearchMain,height = 5, width = 25)
        self.BudetCodeDescription.place(x = 200, y = 180)

        self.TotalExpense = Label(self.SearchMain,text = "Total Expense: ",bg = "#CCFFFF")
        self.TotalExpenseEntry = Entry(self.SearchMain)
        self.TotalExpense.place(x = 90, y = 280)
        self.TotalExpenseEntry.place(x = 200, y = 280)


        def PurchaseThings():
            self.wb = openpyxl.load_workbook("Budget.xlsm", keep_vba = True)
            self.sheet = self.wb[self.jobonly + "Report"]
            self.sheetxlrdreport = self.workbookxlrd.sheet_by_name(self.jobonly + "Report")
            self.sheetxlrd = self.workbookxlrd.sheet_by_name(self.jobonly)
            Empty_row = self.sheet.max_row + 1

            row = 3
            #try:
            #    self.var = self.variable.get()
            #    self.var = eval(self.var)
            #except:
            #    messagebox.showinfo("Error","Try opening and closing the excel file")

            self.var = self.variable.get()
            self.var = eval(self.var)




            if self.variable == 6111:
                row = 3
            elif self.var ==6113:
                row = 4
            elif self.var == 6121:
                row = 5
            elif self.var == 6211:
                row = 6
            elif self.var == 6212:
                row = 7
            elif self.var == 6218:
                row = 8
            elif self.var == 6233:
                row = 9
            elif self.var ==6240:
                row = 10
            elif self.var == 6256:
                row = 11
            elif self.var == 6271:
                row = 12
            else:
                pass

            #print(self.sheetxlrd.cell_value(row - 1,3))
            #print(row)
            self.wb = openpyxl.load_workbook("Budget.xlsm", keep_vba = True)
            self.sheet = self.wb[self.jobonly + "Report"]
            print(row)
            print(self.sheetxlrd.cell_value(row - 1 ,3))
            self.undisbursed = self.sheetxlrd.cell_value(row - 1 ,3)
            #print(self.undisbursed)
            #print(type(self.undisbursed))

            for i in range(3,100):
                print(self.sheet.cell(row = i,column = 1).value )
                if self.sheet.cell(row = i,column = 1).value == None:
                    Empty_row = i
                    #print(Empty_row,self.sheet.cell(row = i,column = 10).value,i)
                    break
                else:
                    continue
            #print(Empty_row)
            try:
                if self.undisbursed >= eval(self.TotalExpenseEntry.get()):
                        self.sheet.cell(Empty_row,1).value = self.DateEntry["text"]
                        self.sheet.cell(Empty_row,2).value = self.var
                        self.sheet.cell(Empty_row,3).value = self.BudetCodeDescription.get("1.0",END)
                        self.sheet.cell(Empty_row,4).value = eval(self.TotalExpenseEntry.get())
                        print(self.sheet.cell(Empty_row,4).value)
                        self.wb.save("Budget.xlsm")
                        messagebox.showinfo("","Purchased")
                        #self.wb.close() 
                else:
                    messagebox.showinfo("Error","Above Budget")
            except:
                messagebox.showinfo("Error","The excel file seems to be open try to open and close it")

        self.PurchaseButton = Button(self.SearchMain,text = "Purchase", width = 10, height = 2,command = PurchaseThings)
        self.PurchaseButton.place(x = 230, y = 350)

        

class Manager(Department):
    def __init__(self,job):
        super().__init__(job)        
        self.Purchase.destroy()
        self.Report.destroy()
        self.BudgetCode.destroy()
        self.Distribute.destroy()
        self.Related.destroy()
        self.welcometext.destroy()

        self.main = Frame(root, width = 800, height = 630,bg = "#CCFFFF")
        self.main.place(relx = 0.2,rely =0.22)
        
        self.SearchMain = Frame(self.main,width = 800, height = 500,bg = "#CCFFFF")
        self.SearchMain.place(y = 50)

        self.title = Label(self.SearchMain, text = "Welcome",font=("Courier", 44),bg = "#CCFFFF")
        self.title.place(x = 290, y = 150)

        self.BudgetSet = Button(self.actions, text = "Set Budget", command = self.SetBudget)
        self.BudgetSet.place(x = 40, y = 300, width = "130", height = "40")
    

        self.Related = Button(self.actions, text = "Requests", command = self.Budgetupdate)
        self.Related.place(x = 40, y = 380, width = "130", height = "40")

        self.Askupdate = Button(self.actions, text = "Additional Request", command = self.Budgetrelated)
        self.Askupdate.place(x = 40, y = 460, width = "130", height = "40")

    def Budgetupdate(self):
        self.title.destroy()
        self.SearchMain.destroy()
        #except AttributeError:
            #pass
        try:
            self.title1.destroy()
        except:
            pass
        try:
            self.Report.destroy()
        except:
            pass
        self.title = Label(self.main,text = "Approval Form",font = ("Roboto",19),bg = "#CCFFFF")
        self.title.place(y=10,x = 300)
        self.SearchMain = Frame(self.main,width = 800, height = 800,bg = "#CCFFFF")
        self.SearchMain.place(y = 50)

        self.wb = openpyxl.load_workbook("Budget.xlsm", keep_vba = True)
        self.sheet = self.wb["AAiTFinance"]

        loc = ("Budget.xlsm")
        self.workbookxlrd = xlrd.open_workbook(loc)
        self.sheetxlrd = self.workbookxlrd.sheet_by_name("AAiTFinance")

        def approvefun(vary_row,button1,button2):
            self.sheet.cell(row = vary_row, column = 5).value  = "Approved"
            self.wb.save("Budget.xlsm")
            #self.wb.close()
            try:
                messagebox.showinfo(" ","Accepted")
            except:
                pass

        def rejectfun(vary_row,button1,button2):
            self.sheet.cell(row = vary_row, column = 5).value  = "Rejected"
            self.wb.save("Budget.xlsm")
            #self.wb.close()
            try:
                messagebox.showinfo(" ","Rejected")
            except:
                pass



        self.Approve1 = Button(self.SearchMain, text = "Approve", command = lambda: approvefun(3,self.Approve1,self.Reject1))
        self.Approve2 = Button(self.SearchMain, text = "Approve",command = lambda: approvefun(4,self.Approve1,self.Reject1))
        self.Approve3 = Button(self.SearchMain, text = "Approve",command = lambda: approvefun(5,self.Approve1,self.Reject1))
        self.Approve4 = Button(self.SearchMain, text = "Approve",command = lambda: approvefun(6,self.Approve1,self.Reject1))
        self.Approve5 = Button(self.SearchMain, text = "Approve",command = lambda: approvefun(7,self.Approve1,self.Reject1))
        self.Approve6 = Button(self.SearchMain, text = "Approve",command = lambda: approvefun(8,self.Approve1,self.Reject1))

        self.Reject1 = Button(self.SearchMain, text = "Reject", command = lambda: rejectfun(3,self.Approve1,self.Reject1))
        self.Reject2 = Button(self.SearchMain, text = "Reject", command = lambda: rejectfun(4,self.Approve1,self.Reject1))
        self.Reject3 = Button(self.SearchMain, text = "Reject", command = lambda: rejectfun(5,self.Approve1,self.Reject1))
        self.Reject4 = Button(self.SearchMain, text = "Reject", command = lambda: rejectfun(6,self.Approve1,self.Reject1))
        self.Reject5 = Button(self.SearchMain, text = "Reject", command = lambda: rejectfun(7,self.Approve1,self.Reject1))
        self.Reject6 = Button(self.SearchMain, text = "Reject", command = lambda: rejectfun(8,self.Approve1,self.Reject1))

        try:
            for i in range(2,9):
                    Label(self.SearchMain,bg = "#CCFFFF",text = self.sheetxlrd.cell_value(i - 1,1)).grid(row = i + 2,column = 10, pady = 20, padx = 40)
                    Label(self.SearchMain,bg = "#CCFFFF",text = self.sheetxlrd.cell_value(i - 1,2)).grid(row = i + 2,column = 25, pady = 20, padx = 40)
                    Label(self.SearchMain,bg = "#CCFFFF",text = self.sheetxlrd.cell_value(i - 1,3)).grid(row = i + 2,column = 40, pady = 20, padx = 40)
                    #self.workbookxlrd = xlrd.open_workbook(loc)
                    #self.sheetxlrd = self.workbookxlrd.sheet_by_name("AAiTFinance")
                    #print(self.sheetxlrd.cell_value(2,2) )
                    self.sheet = self.wb["AAiTFinance"]

                    loc = ("Budget.xlsm")
                    self.workbookxlrd = xlrd.open_workbook(loc)
                    self.sheetxlrd = self.workbookxlrd.sheet_by_name("AAiTFinance")

                    if self.sheetxlrd.cell_value(2,4) != "Approved" and  self.sheetxlrd.cell_value(2,4) != "Rejected":
                        if self.sheetxlrd.cell_value(2,2) > 0 and (self.sheetxlrd.cell_value(2,4) != "Approved" and  self.sheetxlrd.cell_value(2,3) != "Rejected"):
                            self.Approve1.grid(row = 5 ,column = 60,padx = 10)
                            self.Reject1.grid(row = 5 ,column = 80)
                    
                    if self.sheetxlrd.cell_value(3,4) != "Approved" and  self.sheetxlrd.cell_value(3,4) != "Rejected":
                        if self.sheetxlrd.cell_value(3,2) > 0 and (self.sheetxlrd.cell_value(3,4) != "Approved" and  self.sheetxlrd.cell_value(3,3) != "Rejected"):
                            self.Approve2.grid(row = 6 ,column = 60,padx = 10)
                            self.Reject2.grid(row = 6 ,column = 80)

                    if self.sheetxlrd.cell_value(4,4) != "Approved" and  self.sheetxlrd.cell_value(4,4) != "Rejected":
                        if self.sheetxlrd.cell_value(4,2) > 0 and (self.sheetxlrd.cell_value(4,4) != "Approved" and  self.sheetxlrd.cell_value(4,3) != "Rejected"):
                            self.Approve3.grid(row = 7 ,column = 60,padx = 10)
                            self.Reject3.grid(row = 7 ,column = 80)

                    if self.sheetxlrd.cell_value(5,4) != "Approved" and  self.sheetxlrd.cell_value(2,4) != "Rejected":
                        if self.sheetxlrd.cell_value(5,2) > 0 and (self.sheetxlrd.cell_value(5,4) != "Approved" and  self.sheetxlrd.cell_value(5,3) != "Rejected"):
                            self.Approve4.grid(row = 8 ,column = 60,padx = 10)
                            self.Reject4.grid(row = 8 ,column = 80)

                    if self.sheetxlrd.cell_value(6,4) != "Approved" and  self.sheetxlrd.cell_value(2,4) != "Rejected":
                        if self.sheetxlrd.cell_value(6,2) > 0 and (self.sheetxlrd.cell_value(6,3) != "Approved" and  self.sheetxlrd.cell_value(6,3) != "Rejected"):
                            self.Approve5.grid(row = 9 ,column = 60,padx = 10)
                            self.Reject5.grid(row = 9 ,column = 80)

                        if self.sheetxlrd.cell_value(7,2) > 0 and (self.sheetxlrd.cell_value(7,3) != "Approved" and self.sheetxlrd.cell_value(7,3) != "Rejected"):
                            self.Approve6.grid(row = 10 ,column = 60,padx = 10)
                            self.Reject6.grid(row = 10 ,column = 80)
        except:
            messagebox.showinfo(" ","Looks like the excel file is open, try closing it and login again")

        
    def Budgetrelated(self):
        try:
            self.SubmitButton.destroy()
            self.submit.destroy()
            self.SearchMain.destroy()
        except AttributeError:
            self.SearchMain.destroy()
        try:
            self.title1.destroy()
        except:
            pass
        try:
            self.title.destroy()
        except:
            pass
        try:
            self.Report.destroy()
        except:
            pass
        self.SearchMain = Frame(self.main,width = 800, height = 500,bg = "#CCFFFF")
        self.SearchMain.place(x = 80, y = 50)

        self.Annual_Budget = 0

        loc = ("Budget.xlsm")
        self.workbookxlrd = xlrd.open_workbook("Budget.xlsm")
        self.sheetxlrd = self.workbookxlrd.sheet_by_name("AAiTFinance")

        self.AnnualBudget = Label(self.SearchMain,text = "Annual Budget: ",bg = "#CCFFFF",font = ('Roboto',14))
        self.AnnualBudgetVariable = StringVar(self.SearchMain)
        self.AnnualBudgetVariable = self.sheetxlrd.cell_value(11,1)
        self.AnnualBudgetValue = Label(self.SearchMain, text = self.AnnualBudgetVariable,font = ('Roboto',14),bg = "#CCFFFF")
        self.AnnualBudget.place(x = 50, y = 60)
        self.AnnualBudgetValue.place(x = 220, y = 62)

        self.AdjustedBudget = Label(self.SearchMain,text = "Adjusted Budget: ",bg = "#CCFFFF",font = ('Roboto',14))
        self.AdjustedBudgetVariable = StringVar(self.SearchMain)
        self.AdjustedBudgetVariable = self.sheetxlrd.cell_value(14,1)
        self.AdjustedBudgetValue = Label(self.SearchMain, text = self.AdjustedBudgetVariable,font = ('Roboto',14),bg = "#CCFFFF")
        self.AdjustedBudget.place(x = 350, y = 60)
        self.AdjustedBudgetValue.place(x = 540, y = 62)
        
        self.Report = Label(self.main, text = "Ask for an increase of Annual Budget",font=("Times New Roman", 18),bg = "#CCFFFF")
        self.Report.place(x = 250, y =10)

        self.Transfer = Label(self.SearchMain, text = "Amount of Transfer: ",font = ('Roboto',13),bg = "#CCFFFF")
        self.TransferEntry = Entry(self.SearchMain)
        self.Transfer.place(x = 50 , y = 150)
        self.TransferEntry.place(x = 220 , y = 150)

        self.Reason = Label(self.SearchMain, text = "Reason for enquiry: ",font = ('Times New Roman',13),bg = "#CCFFFF")
        self.Reason.place(x = 50, y = 200)
        self.Reasonraise = Text(self.SearchMain,height = 7, width = 40)
        self.Reasonraise.place(x = 220, y = 200)

        def AskAddBudget():
            self.wb = openpyxl.load_workbook("Budget.xlsm", keep_vba = True)
            self.sheet = self.wb["AAiTFinance"]
            self.Update = self.TransferEntry.get()
            self.sheet.cell(13,2).value = eval(self.Update)
            self.sheet.cell(14,2).value = self.Reasonraise.get("1.0",END)
            messagebox.showinfo(" ","Additional budget asked")
            self.wb.save("Budget.xlsm") 
                
        self.SubmitButton = Button(self.SearchMain, text = "Submit",command = AskAddBudget)
        self.SubmitButton.place(x = 250, y = 400,width = "150", height = "50")

        #for i in range()

    def SetBudget(self):
        try:
            self.submit.destroy()
        except:
            pass
        try:
            self.title1.destroy()
        except:
            pass
        try:
            self.title.destroy()
        except:
            pass
        try:
            self.Report.destroy()
        except:
            pass
        
        self.SearchMain.destroy()
        self.SearchMain = Frame(self.main,width = 800, height = 800,bg = "#CCFFFF")
        self.SearchMain.place( y = 50)

        self.title1 = Label(self.main, text = "Set The Annual Budget For Each Department",font=("Times New Roman", 18) ,bg = "#CCFFFF")
        self.title1.place(x = 180, y = 10)
        
        #self.fill_out = Frame(self.SearchMain, )
        self.Biomedical = Label(self.SearchMain,text = "Biomedical Engineering: ",bg = "#CCFFFF")
        self.BiomedicalEntry = Entry(self.SearchMain)
        self.Biomedical.place(x = 230, y = 50)
        self.BiomedicalEntry.place(x = 380, y = 50)
        
        self.Chemical = Label(self.SearchMain,text = "Chemical Engineering: ",bg = "#CCFFFF")
        self.ChemicalEntry = Entry(self.SearchMain)
        self.Chemical.place(x =230, y = 100)
        self.ChemicalEntry.place(x = 380, y = 100)

        self.ITSC = Label(self.SearchMain,text = "ITSC: ",bg = "#CCFFFF")
        self.ITSCEntry = Entry(self.SearchMain)
        self.ITSC.place(x = 230, y = 150)
        self.ITSCEntry.place(x = 380, y = 150)

        self.Civil = Label(self.SearchMain,text = "Civil Engineering: ",bg = "#CCFFFF")
        self.CivilEntry = Entry(self.SearchMain)
        self.Civil.place(x = 230, y = 200)
        self.CivilEntry.place(x = 380, y = 200)

        self.Electrical = Label(self.SearchMain,text = "Electrical Engineering: ",bg = "#CCFFFF")
        self.ElectricalEntry = Entry(self.SearchMain)
        self.Electrical.place(x = 230, y = 250)
        self.ElectricalEntry.place(x = 380, y = 250)

        self.Mechanical = Label(self.SearchMain,text = "Mechanical Engineering: ",bg = "#CCFFFF")
        self.MechanicalEntry = Entry(self.SearchMain)
        self.Mechanical.place(x = 230, y = 300)
        self.MechanicalEntry.place(x = 380, y = 300)

        #self.deps = ["Biomedical","Civil","ITSC","Mechanical","Electrical" , "Chemical"]

        def submits():
            self.wb = openpyxl.load_workbook("Budget.xlsm", keep_vba = True)
            self.sheet = self.wb["AAiTFinance"]
            print(self.sheet.cell(3,2).value)
            self.sheet.cell(3,2).value = eval(self.BiomedicalEntry.get())
            self.sheet.cell(4,2).value = eval(self.ChemicalEntry.get())
            self.sheet.cell(5,2).value = eval(self.ITSCEntry.get())
            self.sheet.cell(6,2).value = eval(self.CivilEntry.get())
            self.sheet.cell(7,2).value = eval(self.ElectricalEntry.get())
            self.sheet.cell(8,2).value = eval(self.MechanicalEntry.get())
            self.wb.save("Budget.xlsm")
            messagebox.showinfo(" ","Budget distributed")
##


        self.submit = Button(self.SearchMain, text = "Submit", command = submits)
        self.submit.place(x = 340, y = 380)
        #self.BudgetCode.place(x = 40, y = 420, width = "130", height = "40" 
        #self.actions.destroy()
        #self.Purchase = Button(self.actions, text = "Set Annual Budget")
        #self.AnnualBudget = Button(self.actions, text = "Budget transfer")
        #self.Report = Button(self.actions, text = " Total Expenditure report")
        #self.BudgetCode = Button(self.actions,text = "Budget Code Help")
class AAU(Manager):
    def __init__(self,job):
        super().__init__(job)
        self.BudgetSet.destroy()
        self.Related.destroy()
        self.job = "AAU Finance"
        self.title.destroy()
        self.Askupdate.destroy()


        self.Askupdate = Button(self.actions, text = "Requests", command = self.Budgetupdate)
        self.Askupdate.place(x = 40, y = 400, width = "130", height = "40")

        self.SearchMain = Frame(self.main,width = 800, height = 800,bg = "#CCFFFF")
        self.SearchMain.place(y = 0)
        self.BudgetSet = Button(self.actions, text = "Set Budget", command = self.SetBudget)
        self.BudgetSet.place(x = 40, y = 300, width = "130", height = "40")

        self.title = Label(self.SearchMain, text = "Welcome",font=("Courier", 44),bg = "#CCFFFF")
        self.title.place(x = 290, y = 150)

        loc = ("Budget.xlsm")                                                       
        self.workbookxlrd = xlrd.open_workbook(loc)
        self.sheetxlrd = self.workbookxlrd.sheet_by_name("AAUFinance")

    def SetBudget(self):
        self.SearchMain.destroy()
        self.SearchMain = Frame(self.main,width = 800, height = 800,bg = "#CCFFFF")
        self.SearchMain.place( y = 50)

        self.title = Label(self.main, text = "Set The Annual Budget For AAiT",font=("Times New Roman", 18) ,bg = "#CCFFFF")
        self.title.place(x = 200, y = 10)

        #self.fill_out = Frame(self.SearchMain, )
        self.aaitbudget = Label(self.SearchMain,text = "Amount of Budget: ",bg = "#CCFFFF" ,font = (5))
        self.aaitbudgetEntry = Entry(self.SearchMain)
        self.aaitbudget.place(x = 170, y = 150)
        self.aaitbudgetEntry.place(x = 330, y = 150)

        self.wb = openpyxl.load_workbook("Budget.xlsm", keep_vba = True)
        self.sheet = self.wb["AAUFinance"]

        def totalbudget():
            self.sheet.cell(row = 3, column = 2).value = eval(self.aaitbudgetEntry.get())
            self.wb.save("Budget.xlsm")
            #self.wb.forget()
            messagebox.showinfo(" ","Budget set!")
            print(self.sheet.cell(row = 3, column = 2).value)

        self.submit = Button(self.SearchMain, text = "Submit",command = totalbudget)
        self.submit.place(x = 320, y = 200)

    def Budgetupdate(self):
        try: 
            self.title.destroy()
            self.SearchMain.destroy()
        except AttributeError:
            pass

        
        self.SearchMain = Frame(self.main,width = 800, height = 500,bg = "#CCFFFF")
        self.SearchMain.place(x = 80, y = 50)

        self.wb = openpyxl.load_workbook("Budget.xlsm", keep_vba = True)
        self.sheet = self.wb["AAUFinance"]

        loc = ("Budget.xlsm")
        self.workbookxlrd = xlrd.open_workbook(loc)
        self.sheetxlrd = self.workbookxlrd.sheet_by_name("AAUFinance")

        
        try:
            self.title.destroy()
            self.SearchMain.destroy()
        except AttributeError:
            pass

        self.SearchMain = Frame(self.main,width = 800, height = 500,bg = "#CCFFFF")
        self.SearchMain.place(x = 80, y = 50)

        def approvefun(vary_row,button1,button2):
            self.sheet.cell(row = 3, column = 5).value  = "Approved"
            self.wb.save("Budget.xlsm")
            #self.wb.close()
            try:
                messagebox.showinfo(" ","Accepted")
            except:
                pass

        def rejectfun(vary_row,button1,button2):
            self.sheet.cell(row = 3, column = 5).value  = "Rejected"
            self.wb.save("Budget.xlsm")
            s#elf.wb.close()
            try:
                messagebox.showinfo(" ","Rejected")
            except:
                pass

        self.Approve1 = Button(self.SearchMain, text = "Approve", command = lambda: approvefun(3,self.Approve1,self.Reject1))
        self.Reject1 = Button(self.SearchMain, text = "Reject", command = lambda: rejectfun(3,self.Approve1,self.Reject1))

        self.Askedlabel = Label(self.SearchMain, text = "Budget transfer asked",bg = "#CCFFFF",font = (4))
        self.Askedlabel.place(x = 250, y = 80)

        self.AAiTasked = Label(self.SearchMain, text = "AAiT",bg = "#CCFFFF",font = (6))#.grid(row = 2,column = 10, pady = 20, padx = 40)
        self.AAiTaskedamount = Label(self.SearchMain,text = self.sheetxlrd.cell_value(2,2),bg = "#CCFFFF",font = (4) )#.grid(row =  2,column = 40, pady = 20, padx = 40)
        self.text = Label(self.SearchMain, text = "Reason",bg = "#CCFFFF",font = (4))
        self.text.place(x = 250, y = 120)
        self.AAiTasked.place(x  = 90, y = 150)        
             
        
        #print(self.sheet.cell(row = 3, column = 5))
        try:
            if self.sheet.cell(row = 3, column = 5).value != "Approved" and self.sheet.cell(row = 3, column = 5).value != "Rejected":
                if self.sheetxlrd.cell_value(2,2) > 0 and (self.sheetxlrd.cell_value(2,3) != "Approved" and  self.sheetxlrd.cell_value(2,3) != "Rejected"):
                    self.Reasontext = Label(self.SearchMain,text = self.sheetxlrd.cell_value(2,3))
                    self.Reasontext.place(x = 250, y = 150 )
                    self.AAiTaskedamount.place(x = 150, y = 150)
                    self.Approve1.place(x  = 400, y = 150)
                    self.Reject1.place(x  = 500, y = 150)
            else:
                self.resolved = Label(self.SearchMain, text = "No new transfers",bg = "#CCFFFF",font = (4))
                self.resolved.place(x = 300, y = 150)
        except:
            messagebox.showinfo("Error","Try opening and closing the excel file and signin again")


class ICT(Department):
    def __init__(self,job):
        super().__init__(job)
        self.job = "ICT"
        self.Purchase.destroy()
        self.Report.destroy()
        self.BudgetCode.destroy()
        self.Distribute.destroy()
        self.Related.destroy()
        self.welcometext.destroy()

        
        self.SearchMain = Frame(self.main,width = 800, height = 800,bg = "#CCFFFF")
        self.SearchMain.place( y = 0)
        
        self.title = Label(self.SearchMain, text = "Welcome",font=("Courier", 44),bg = "#CCFFFF")
        self.title.place(x = 300, y = 220)

        self.Password_change = Button(self.actions, text = "Change Password", command = self.ChangePassword)#, command = self.SetBudget)
        self.AddUser = Button(self.actions, text = "Add User",command = self.adduser)
        self.Activate = Button(self.actions,text = "Activate account", command = self.activate)#, command = self.BudgetcodeFinder)#,command = pass)

        self.Password_change.place(x = 40, y = 220, width = "130", height = "40")
        self.AddUser.place(x = 40, y = 280, width = "130", height = "40")
        self.Activate.place(x = 40, y = 340, width = "130", height = "40")

    def adduser(self):
    

        try:
            self.title1.destroy()
        except:
            pass
        try:
            self.moneyiconlabel.destroy()
        except:
                pass
        try:
            self.title2.destroy()
        except:
            pass
        try:
            self.title3.destroy()
        except:
            pass
        try:
            self.title.destroy()
        except:
            pass
        try:
            self.SearchMain.destroy()
        except:
            pass          
##        self.SearchMain = Frame(self.main,width = 800, height = 500,bg = "#CCFFFF")
##        self.SearchMain.place(x = 80, y = 50)
        self.SearchMain = Frame(self.main,width = 800, height = 800,bg = "#CCFFFF")
        self.SearchMain.place( y = 50)
        
        self.title1 = Label(self.main, text = "New User form",font=("Roboto", 18) ,bg = "#CCFFFF")
        self.title1.place(x = 290, y = 10)

        self.User = Label(self.SearchMain,text = "Role: ",bg = "#CCFFFF")
        self.User.place(x = 230, y = 70)
        self.UserEntry = Entry(self.SearchMain)
        self.UserEntry.place(x = 330, y = 70)

        #self.NewName = Label(self.SearchMain,text = "Name: ",bg = "#CCFFFF")
        #self.NewNameEntry = Entry(self.SearchMain)
        #self.NewName.place(x = 230, y = 120)
        #self.NewNameEntry.place(x = 330, y = 120)

        self.NewSecurity = Label(self.SearchMain,text = "UserName: ",bg = "#CCFFFF")
        self.NewSecurityEntry = Entry(self.SearchMain)
        self.NewSecurity.place(x = 230, y = 120)
        self.NewSecurityEntry.place(x = 330, y = 120)

    
        self.NewPassword = Label(self.SearchMain,text = "New Password: ",bg = "#CCFFFF")
        self.NewPasswordEntry = Entry(self.SearchMain)
        self.NewPassword.place(x = 230, y = 170)
        self.NewPasswordEntry.place(x = 330, y = 170)

        self.variable = StringVar(self.SearchMain)
        var_list = ["Activate", "Deactivate"]
        self.variable.set("Deactivate")
        
        self.moneyicon = PhotoImage (file = 'iconss.png')
        self.moneyiconlabel = Label(self.SearchMain,image = self.moneyicon,bg = "#CCFFFF")
        self.moneyiconlabel.place(x = 320,y = 10)

        self.activates = Label(self.SearchMain,text = "Activate: ",bg = "#CCFFFF")
        self.activates.place(x = 230, y = 210)

        Dropdown = OptionMenu(self.SearchMain, self.variable, *var_list)
        Dropdown.place(x = 330 , y = 210)

        def adds():
            self.sheet = self.wb["Password"]
            Empty_row = self.sheet.max_row + 1

            try:    
                self.sheet.cell(Empty_row,1).value = self.UserEntry.get()
                #self.sheet.cell(Empty_row,2).value = self.NewNameEntry.get()
                self.sheet.cell(Empty_row,2).value = self.NewSecurityEntry.get()
                self.sheet.cell(Empty_row,3).value = eval(self.NewPasswordEntry.get())
                self.sheet.cell(Empty_row,4).value = self.variable.get()
                self.wb.save("Budget.xlsm")
                messagebox.showinfo("","User added")
            except:
                messagebox.showinfo("Error","Try using numbers")

        self.change = Button(self.SearchMain, text = "Add", command = adds)
        self.change.place(x = 336, y = 320,width = "90", height = "40")

    def ChangePassword(self):
        try:
            self.title1.destroy()
        except:
            pass
        try:
            self.moneyiconlabel.destroy()
        except:
                pass
        try:
            self.title2.destroy()
        except:
            pass
        try:
            self.title3.destroy()
        except:
            pass
        try:
            self.title.destroy()
        except:
            pass
        try:
            self.SearchMain.destroy()
        except:
            pass
               
        self.SearchMain = Frame(self.main,width = 800, height = 800,bg = "#CCFFFF")
        self.SearchMain.place( y = 50)
        
        self.title2 = Label(self.main, text = "change password",font=("Roboto", 18) ,bg = "#CCFFFF")
        self.title2.place(x = 290, y = 30)
        self.sheet = self.wb["Password"]
        User_list = [self.sheet.cell(i,2).value for i in range(3,self.sheet.max_row + 1)]

        self.variable = StringVar(self.SearchMain)
        self.variable.set(User_list[0]) # default value

        Dropdown = OptionMenu(self.SearchMain, self.variable, *User_list)
        Dropdown.place(x = 410 , y = 100)

        self.User = Label(self.SearchMain,text = "Choose User: ",font = ('Roboto',11),bg = "#CCFFFF")
        self.User.place(x = 280, y = 100)

        self.NewPassword = Label(self.SearchMain,text = "New Password: ",font = ('Roboto',11),bg = "#CCFFFF")
        self.NewPasswordEntry = Entry(self.SearchMain)
        self.NewPassword.place(x = 280, y =160)
        self.NewPasswordEntry.place(x = 410, y = 160)

    
        def change():
            self.sheet = self.wb["Password"]
            row = 2
            for i in range(2,self.sheet.max_row + 1):
                #print(self.sheet.cell(i,2).value,self.variable.get())
                if self.sheet.cell(i,2).value == self.variable.get():
                    row = i
            try:        
                self.sheet.cell(row,3).value = eval(self.NewPasswordEntry.get())
                self.wb.save("Budget.xlsm")
                messagebox.showinfo("","Password changed")
            except:
                messagebox.showinfo("Invalid password type","Should use only numbers")



        self.change = Button(self.SearchMain, text = "Change", command = change)
        self.change.place(x = 380, y = 200)

    def activate(self):
        try:
            self.title1.destroy()
        except:
            pass
        try:
            self.moneyiconlabel.destroy()
        except:
                pass
        try:
            self.title2.destroy()
        except:
            pass
        try:
            self.title3.destroy()
        except:
            pass
        try:
            self.title.destroy()
        except:
            pass
        try:
            self.SearchMain.destroy()
        except:
            pass
        
        self.SearchMain = Frame(self.main,width = 800, height = 800,bg = "#CCFFFF")
        self.SearchMain.place( y = 50)
        
        self.title3 = Label(self.main, text = "User account issues",font=("Roboto", 18) ,bg = "#CCFFFF")
        self.title3.place(x = 220, y = 30)

        self.name = Label(self.SearchMain,text = "Username: ",font = ('Times', 13),bg = "#CCFFFF")
        self.name.place(x = 225, y = 70)

        self.sheet = self.wb["Password"]
        User_list = [self.sheet.cell(i,2).value for i in range(3,self.sheet.max_row + 1)]

        self.variable = StringVar(self.SearchMain)
        self.variable.set(User_list[0]) # default value

        Dropdown = OptionMenu(self.SearchMain, self.variable, *User_list)
        Dropdown.place(x = 400 , y = 70)

        self.row = 2
        for i in range(2,self.sheet.max_row + 1):
            if self.sheet.cell(i,2).value == self.variable.get():
                self.row = i

        print(self.sheet.cell(self.row,4).value)
        print(self.row)

        def release():
            self.sheet = self.wb["Password"]
            self.sheet.cell(self.row,4).value = "Activate"
            self.wb.save("Budget.xlsm")
            messagebox.showinfo("Done","Activated")
            self.Release = Button(self.SearchMain, text = "Deactivate Account", command = deactivate)
            self.Release.place( x = 400, y = 200, width = "130", height = "40")

        def deactivate():
            self.sheet = self.wb["Password"]
            self.sheet.cell(self.row,4).value = "Deactivate"
            self.wb.save("Budget.xlsm")
            messagebox.showinfo("Done","Deactivated")            
            self.Release = Button(self.SearchMain, text = "Activate Account", command = release)
            self.Release.place( x = 400, y = 200, width = "130", height = "40")


        if self.sheet.cell(self.row,4).value == "Activate":
            self.Release = Button(self.SearchMain, text = "Deactivate Account", command = deactivate)
            self.Release.place( x = 400, y = 200, width = "130", height = "40")
        elif self.sheet.cell(self.row,4).value == "Deactivate":
            self.Release = Button(self.SearchMain, text = "Activate Account", command = release)
            self.Release.place( x = 400, y = 200, width = "130", height = "40")
        else:
            self.sheet.cell(self.row,4).value = "Activate"

        


root= Tk()
New = Home(root)
root.mainloop()
