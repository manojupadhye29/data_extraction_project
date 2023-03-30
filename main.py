from subprocess import check_call
from tkinter import messagebox, Scrollbar, Menu
from os.path import abspath, exists
from customtkinter import CTk, set_appearance_mode, set_default_color_theme, CTkFrame, CTkButton, CTkLabel, \
    CTkOptionMenu, CTkTextbox, set_widget_scaling, CTkInputDialog
from os import mkdir, startfile, getcwd, path
from datetime import datetime
from glob import glob
from openpyxl import Workbook
from openpyxl.styles import Side, Border
from win32com.client import Dispatch
from tabulate import tabulate
from re import match

set_appearance_mode("Dark")  # Modes: "System" (standard), "Dark", "Light"
set_default_color_theme("blue")  # Themes: "blue" (standard), "green", "dark-blue"

class App(CTk):
    def __init__(self):
        super().__init__()
        outdir = 'TerumoBCT Pending Training Details'
        if not exists(outdir):
            mkdir(outdir)

        self.default_mail()

        # configure window
        self.state('zoomed')
        self.title("TerumoBCT Pending Training Data Extraction Tool")

        # configure grid layout (4x4)
        self.grid_columnconfigure(1, weight=1)
        self.grid_columnconfigure((2, 6), weight=0)
        self.grid_rowconfigure((0, 1, 2), weight=1)

        # create sidebar frame with widgets
        self.sidebar_frame = CTkFrame(self, width=140, corner_radius=0)
        self.sidebar_frame.grid(row=0, column=3, rowspan=4, sticky="nsew")
        self.sidebar_frame.grid_rowconfigure(6, weight=1)

        # create submit button
        self.string_input_button = CTkButton(self.sidebar_frame, text="Extract and Mail",
                                             command=self.data_extraction)
        self.string_input_button.grid(row=0, column=0, padx=0, pady=(25, 10))

        # create open folder button
        self.string_input_button = CTkButton(self.sidebar_frame, text="Open Folder",
                                             command=self.openFolder)
        self.string_input_button.grid(row=1, column=0, padx=0, pady=(10, 10))

        # create open latest file button
        self.string_input_button = CTkButton(self.sidebar_frame, text="Open Latest File",
                                             command=self.openFile)
        self.string_input_button.grid(row=2, column=0, padx=0, pady=(10, 10))

        # change mail
        self.string_input_button = CTkButton(self.sidebar_frame, text="Modify Mail",
                                             command=self.open_input_dialog_event)
        self.string_input_button.grid(row=3, column=0, padx=0, pady=(10, 10))

        # create clear button
        self.string_input_button = CTkButton(self.sidebar_frame, text="Clear", command=self.delete_text)
        self.string_input_button.grid(row=4, column=0, padx=0, pady=(10, 10))

        # create close button
        self.string_input_button = CTkButton(self.sidebar_frame, text="Close", command=self.close)
        self.string_input_button.grid(row=5, column=0, padx=0, pady=(10, 10))

        # System UI label and button
        self.appearance_mode_label = CTkLabel(self.sidebar_frame, text="Appearance Mode:", anchor="w")
        self.appearance_mode_label.grid(row=10, column=0, padx=20, pady=(10, 0))
        self.appearance_mode_optionemenu = CTkOptionMenu(self.sidebar_frame,
                                                         values=["Light", "Dark", "System"],
                                                         command=self.change_appearance_mode_event)

        # UI scaling Label and button
        self.appearance_mode_optionemenu.grid(row=11, column=0, padx=20, pady=(10, 10))
        self.scaling_label = CTkLabel(self.sidebar_frame, text="UI Scaling:", anchor="w")
        self.scaling_label.grid(row=12, column=0, padx=20, pady=(10, 0))
        self.scaling_optionemenu = CTkOptionMenu(self.sidebar_frame,
                                                 values=["80%", "90%", "100%", "110%", "120%"],
                                                 command=self.change_scaling_event)
        self.scaling_optionemenu.grid(row=13, column=0, padx=20, pady=(10, 25))

        # create horizontal scrollbar
        h = Scrollbar(self, orient='horizontal')

        # create textbox
        self.textbox = CTkTextbox(self, width=150, height=1200, wrap=None, xscrollcommand=h.set, undo=True)
        self.textbox.grid(row=0, column=1, padx=(20, 20), pady=(20, 20), sticky="nsew")

        # adding right click Button
        m = Menu(self.textbox, tearoff=0)
        m.add_command(label="Cut", command=self.cut)
        m.add_command(label="Copy", command=self.copy)
        m.add_command(label="Paste", command=self.paste)
        m.add_command(label="SelectAll", command=self.select_all)
        m.add_command(label="clear", command=self.delete_text)
        m.add_separator()

        # set the default values
        self.appearance_mode_optionemenu.set("System")
        self.scaling_optionemenu.set("100%")

    # data extraction main program
    def data_extraction(self):
        try:
            lines = self.textbox.get(1.0, "end-1c").split('.')
            self.rows = []
            names = []
            # splitting the data and storing the data i.e. name, course and due date, and storing in different lists.
            for items in lines:
                if '--' in items:
                    split_data = items.split('--')
                    if split_data[0] != '\n\n' and split_data[0] != '\n':
                        # temporary variable to append all the contents to the list
                        temp = []

                        # append the names to list
                        temp.append(split_data[0].split("\n")[-2].strip())
                        names.append(split_data[0].split("\n")[-2].strip())

                        # extract the
                        s = split_data[1].split('"')[1]
                        f = s.split(" ")
                        if len(f[1]) == 1:
                            del f[:2]
                        listToStr = ' '.join([str(elem) for elem in f])

                        # handle exception if there is no value to split
                        try:
                            a = split_data[1].split('"')[3]
                        except:
                            a = ""
                            pass

                        # append the course to list
                        temp.append(listToStr + " / " + a)

                        # append the due dates to list
                        temp.append(split_data[1].split('due on')[1].split()[0])

                        self.rows = self.rows + [temp]

            # check if data exists
            if len(self.rows) > 0:
                # Column header
                column_header = ['Names', 'Course Details', 'Due Date']
                # Getting the current date time
                current_datetime = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")

                # output name for xlsx format
                outname = "TerumoBCT Pending Training Details " + current_datetime + ".xlsx"
                outdir = "TerumoBCT Pending Training Details"

                # check if output directory exists, if not create one
                if not exists(outdir):
                    mkdir(outdir)

                # create an Excel workbook and Excel sheet
                wb = Workbook()
                ws = wb.active

                # set worksheet name
                ws.title = "TerumoBCTPendingTrainingDetails"

                # inserting the header row
                ws.append(column_header)

                # inserting the rows
                for item in self.rows:
                    ws.append(item)

                # get the dimensions to set filter
                ws.auto_filter.ref = ws.dimensions

                # border configurations
                border_type = Side(border_style='thin', color="000000")
                border = Border(top=border_type, bottom=border_type, left=border_type, right=border_type)

                # set borders to the cell
                excelsheet_range = ws[ws.dimensions]
                for cell in excelsheet_range:
                    for x in cell:
                        x.border = border

                # save the Excel file to the path "Extracted Training Due/Terumo BCT Pending YYYY-MM-DD_HH-MM-SS"
                wb.save(f"{outdir}/{outname}")

                path = abspath(f"{outdir}/{outname}")

                excel = Dispatch("Excel.Application")
                wb = excel.Workbooks.Open(path)
                ws = wb.Worksheets("TerumoBCTPendingTrainingDetails")

                # autofit the column width
                ws.Columns.AutoFit()

                # bold header
                ws.Range("A1:H1").Font.Bold = True

                # sort the excel file with names
                wb.Save()
                excel.Application.Quit()

                self.send_mail(names)

                # Data Extracted successfully message
                messagebox.showinfo('Data Extraction Message',
                                    'Data Extraction Completed.\nNotifications Sent to Employees with Unfinished Training Progress.')
                # clear all the data from the message box
                self.textbox.delete("1.0", "end")
            else:
                # Data extraction failed message
                messagebox.showinfo('Data Extraction Error', f'No Data Extracted')
        except Exception as e:
            # clear all the data from the message box
            messagebox.showinfo('Input Error', f'Please Check the input and try again.\n{e}')

    # function to open the folder path
    def openFolder(self):
        # check if output directory exists, if not create one
        outdir = "TerumoBCT Pending Training Details"
        if not path.exists(outdir):
            mkdir(outdir)
            messagebox.showinfo('File Error', f'Folder does not exist.\nHence, created a new folder named {outdir}')

        # open the folder using folder path
        outdir = r'\TerumoBCT Pending Training Details'
        folder_path = getcwd() + outdir
        startfile(folder_path)

    # function to open the latest file
    def openFile(self):
        outdir = r'\TerumoBCT Pending Training Details'
        try:
            folder_path = getcwd() + outdir
            file_type = r'\*xlsx'
            files = glob(folder_path + file_type)
            max_file = max(files, key=path.getctime)
            startfile(max_file)
        except Exception as e:
            messagebox.showinfo("File Error", "No Extracted File Found")

    # clear the content of the textbox
    def delete_text(self):
        self.textbox.delete("1.0", "end")

    # close the window
    def close(self):
        self.destroy()

    # Creating Function for Copy in the textbox
    def copy(self):
        self.textbox.event_generate("<<Copy>>")

    # Creating Function for Paste in the textbox
    def paste(self):
        self.textbox.event_generate("<<Paste>>")

    # Creating Function for cut in the textbox
    def cut(self):
        self.textbox.event_generate("<<Cut>>")

    # Creating Function for select all in the textbox
    def select_all(self):
        self.textbox.event_generate("<<SelectAll>>")

    # change UI to Light, Dark and System
    def change_appearance_mode_event(self, new_appearance_mode: str):
        set_appearance_mode(new_appearance_mode)

    # change the scaling
    def change_scaling_event(self, new_scaling: str):
        new_scaling_float = int(new_scaling.replace("%", "")) / 100
        set_widget_scaling(new_scaling_float)

    # take detsis from TerumoBCT DL email and compare with extracted data, and send mail to DL, non-DL mails accordingly
    def send_mail(self, extracted_names):
        string = '''A, Sivanidhi <sivanidhi.a@capgemini.com>; Amrawat, Rajesh <rajesh.amrawat@capgemini.com>; Amrawat, Rajesh [EXT] <rajesh.amrawat@terumobct.com>; Badgujar, Nikhil <nikhil.durgadas-badgujar@capgemini.com>; Bhamare, Prajakta <prajakta.madan-bhamare@capgemini.com>; Bharatkumar, Shah Dharmik <shah-dharmik.bharatkumar@capgemini.com>; BIPINCHANDRA PATEL, MARGI <margi.bipinchandra-patel@capgemini.com>; Biskitwala, Jigar <jigar.biskitwala@capgemini.com>; BONDRE, SNEHAL <snehal.bondre@capgemini.com>; Chaudhari, Nikhil <nikhil.c.chaudhari@capgemini.com>; Chaudhari, Prasad Balkrishna <prasad-balkrishna.chaudhari@capgemini.com>; Chavda, Kuldipsinh <kuldipsinh.chavda@capgemini.com>; Chindhade, Akash Anil <akash-anil.chindhade@capgemini.com>; Christy, Gabriela <gabriela.christy@capgemini.com>; Darji, Mihir [EXT] <mihir.darji@terumobct.com>; Darji, Mihir Ambalal <mihir.darji@capgemini.com>; Das, Tiyas <tiyas.das@capgemini.com>; Daterao, Samir <samir.daterao@capgemini.com>; Deepakbhai Dhameliya, Krutik <krutik.deepakbhai-dhameliya@capgemini.com>; Desai, Anand <anand.a.desai@capgemini.com>; Devidasrao Kadam, Dipti <dipti.devidasrao-kadam@capgemini.com>; Diliprao Kalanke, Prasad <prasad.diliprao-kalanke@capgemini.com>; Dinesh Mandal, Karan <karan.dinesh-mandal@capgemini.com>; Dixit, Piyush <piyush.dixit@capgemini.com>; DL IG TerumoBCT  Nucleus <terumobctnucleus.ig@capgemini.com>; DL IG TerumoBCT Foxtrot <terumobctfoxtrot.ig@capgemini.com>; DL IG TerumoBCT FrontRunners <terumobctfrontrunners.ig@capgemini.com>; DL IG TerumoBCT Janus <terumobctjanus.ig@capgemini.com>; DL IG TerumoBCT Optivengers <terumobctoptivengers.ig@capgemini.com>; DL IG TerumoBCT Orcas <terumobctorcas.ig@capgemini.com>; DL IG TerumoBCT Patrons <terumobctpatrons.ig@capgemini.com>; DL IG TerumoBCT TechBusters <terumobcttechbusters@capgemini.com>; DL IG TerumoBCT Technocrats <terumobcttechnocrats.ig@capgemini.com>; DL IG TerumoBCT Technophiles <terumobcttechnophiles.ig@capgemini.com>; DL IG TerumoBCT Tesseract <terumobcttesseract.ig@capgemini.com>; DL IG TerumoBCT Trivengers <terumobcttrivengers.ig@capgemini.com>; DL IN DL IG TerumoBCT Mirasol <dligterumobctmirasol@capgemini.com>; DL IN Terumo BCT ScrumMasters <terumobctscrummasters@capgemini.com>; DL IN TerumoBCT Trailblazers <terumobcttrailblazers@capgemini.com>; Fuletra, Utsavkumar <utsavkumar.fuletra@capgemini.com>; Gajane, Sanket <sanket.gajane@capgemini.com>; Gajane, Sanket [EXT] <sanket.gajane@terumobct.com>; Gajjar, Harsh <harsh-himanshu.gajjar@capgemini.com>; Gayakwad, Trupti <trupti.gayakwad@capgemini.com>; Girishbhai Patel, Jemin <jemin.girishbhai-patel@capgemini.com>; Gokani, Sandip <sandip-kantilal.gokani@capgemini.com>; Gupta, Vivek <vivek.d.gupta@capgemini.com>; Jagatsingh Yadav, Manendrasingh <manendrasingh.jagatsingh-yadav@capgemini.com>; Jagdishchandra Shah, Harsh <harsh.jagdishchandra-shah@capgemini.com>; Jalandhar Channa, Anirudh <anirudh.jalandhar-channa@capgemini.com>; Jaunjal, Vrushali Panjabrav <vrushali-panjabrav.jaunjal@capgemini.com>; Jog, Chitrangi <chitrangi.anand-jog@capgemini.com>; Joshi, Niraj <niraj.joshi@capgemini.com>; Joshi, Shaishavkumar <shaishavkumar.joshi@capgemini.com>; K M, Darshan <darshan.k-m@capgemini.com>; Khan, Sahil Salim <sahil-salim.khan@capgemini.com>; Kiritkumar Darji, Darshan <darshan.kiritkumar-darji@capgemini.com>; Koli, Amruta <amruta.koli@capgemini.com>; Kriplani, Minali <minali.kriplani@capgemini.com>; Kukreja, Bhavin Vinod <bhavin-vinod.kukreja@capgemini.com>; Kumar, Yash <yash.c.kumar@capgemini.com>; Macwan, Jinalben Sureshbhai <jinalben-sureshbhai.macwan@capgemini.com>; Macwan, Renison <renison.macwan@capgemini.com>; Mahto, Ekta <ekta.mahto@capgemini.com>; Makwana, Gopi <gopi.anil-kumar-makwana@capgemini.com>; Mehta, Pratik <pratik.b.mehta@capgemini.com>; Mevada, Dhruv Pareshbhai <manoj.upadhye@capgemini.com>; Mewada, Vikram <vikram.mewada@capgemini.com>; Mohanty, Gyanananda <gyanananda.mohanty@capgemini.com>; Mukeshbhai Ray, Vidhi <vidhi.mukeshbhai-ray@capgemini.com>; Narendra Ancharwadkar, Viranchi <viranchi.narendra-ancharwadkar@capgemini.com>; ONKAR NALAWADE, PAYAL <payal.onkar-nalawade@capgemini.com>; Panchal, Jinal <panchal-jinal.suresh@capgemini.com>; Panchal, Ronak Kanhaiyalal <ronak-kanhaiyalal.panchal@capgemini.com>; Parikh, Kruti <kruti.parikh@capgemini.com>; Parmar, Divya <divya.parmar@capgemini.com>; Patel, Darshankumar <darshankumar.patel@capgemini.com>; Patel, Dhruvkumar <dhruvkumar.a.patel@capgemini.com>; Patel, Hirenkumar <hirenkumar.patel@capgemini.com>; Patel, Jatinkumar Ramanbhai <jatinkumar-ramanbhai.patel@capgemini.com>; Patel, KalpeshKumar Ramanbhai <kalpesh.c.patel@capgemini.com>; Patel, Utsav Prashantkumar <utsav-prashantkumar.patel@capgemini.com>; Patel, VikrantKumar <vikrantkumar.kiritbhai-patel@capgemini.com>; Pathak, Vidhi <vidhi.pathak@capgemini.com>; Pathak, Vidhi <vidhi.amish-pathak@capgemini.com>; Patil, Kaivalya <kaivalya.patil@capgemini.com>; Pednekar, Sanika <sanika.vijay-pednekar@capgemini.com>; Pendse, Gauri Surendra <gauri-surendra.pendse@capgemini.com>; Pravinrao Deshmukh, Pranav <pranav.pravinrao-deshmukh@capgemini.com>; Priyam, Ankur <ankur.priyam@capgemini.com>; Raizada, Anushri <anushri.raizada@capgemini.com>; Rajendra Deshmukh, Sanket <sanket.rajendra-deshmukh@capgemini.com>; Rajwadi, Haynes <haynes.rajwadi@capgemini.com>; Rathod, Parth Niraj <parth-niraj.rathod@capgemini.com>; Sakpal, Sagar <sagar.arun-sakpal@capgemini.com>; Sanjay Patil, Vinay <vinay.sanjay-patil@capgemini.com>; Sarkar, Meghna <meghna.sarkar@capgemini.com>; Shah, Arpan <arpan.b.shah@capgemini.com>; Shah, Dhrumil <dhrumil-anilkumar.shah@capgemini.com>; Shah, Heta Nitinkumar <heta-nitinkumar.shah@capgemini.com>; Shah, Shalin <shalin.shah@capgemini.com>; Sharma, Gouri <gouri.sharma@capgemini.com>; Sharma, Simran <simran.d.sharma@capgemini.com>; Shevkani, Manish <manish.shevkani@capgemini.com>; Shimpi, Dimple <dimple-vishwanath.shimpi@capgemini.com>; Singh, Smriti <smriti.d.singh@capgemini.com>; Solanki, Mahmadshahid <mohamedshahid.solanki@capgemini.com>; Sri Neeraj, Chilamkurthi <chilamkurthi.sri-neeraj@capgemini.com>; Sudani, Vishal Kalubhai <vishal-kalubhai.sudani@capgemini.com>; Tony, Annmol <annmol.tony@capgemini.com>; Trivedi, Tejas <tejas.trivedi@capgemini.com>; Trivedi, Vatsal Jigar <vatsal-jigar.trivedi@capgemini.com>; Vinod Kalari Kandi, Liya <liya.vinod-kalari-kandi@capgemini.com>; VINOD NAIR, HARITHA <haritha.vinod-nair@capgemini.com>; Vyas, Manthan <manthan.vyas@capgemini.com>; Vyasa, Paresh <paresh-a.vyasa@capgemini.com>; Zoting, Ashutosh <ashutosh.dnyaneshwar-zoting@capgemini.com>'''
        data = {}
        names = string.split(";")
        for items in names:
            email = items.split(" ")[-1].replace("<", "").replace(">", "")
            temp = items.split(" ")[:-1]
            ename = ' '.join([str(elem) for elem in temp]).replace("[EXT]", "").strip()
            data[ename] = email

        names = list(data.keys())
        email = list(data.values())

        # filter from text file and pass names, email and rows to send_email_main function
        for items in names:
            temp1 = []
            for row in self.rows:
                if items in row[0]:
                    temp1.append(row)
            if len(temp1) > 0:
                # send email to all the people who are there in the DL
                self.send_email_main(items, data[items], temp1)
                pass

        details = []
        for row in self.rows:
            if row[0] not in names:
                details.append(row)
        # send email who are not there in DL (non-DL)
        self.send_mail_main(details)

    # DL mail function
    def send_email_main(self, name, email, val):
        f = open('email.txt', mode='r')
        email = f.read()
        email_final = email.split('\n')

        # configure outook
        ol = Dispatch('Outlook.Application')
        olmailitem = 0x0
        newmail = ol.CreateItem(olmailitem)

        # subject of the mail
        newmail.Subject = 'TerumoBCT Pending Training Details'

        # recipient of the mail
        newmail.To = email

        # carbon copy of the mail
        if len(email.replace('\n', "")) > 0:
            newmail.Cc = ';'.join(email_final)
        else:
            newmail.Cc = 'shalin.shah@capgemini.com'

        # convert list to HTML table
        t = tabulate(val, headers=["Name", "Course Details", "Due Date"], tablefmt="html")
        t = t.replace("<table>", '<table cellspacing="3" cellpadding="3" border="1" bgcolor="#000000">')
        t = t.replace("<tr>", '<tr bgcolor="#ffffff">')

        # add HTML table to the mail
        newmail.HTMLBody = f'<p>Dear {name},<br><br>I am writing to bring to your attention the list of pending training courses that require your attention. Kindly consider this as a priority and make necessary arrangements to complete the training at the earliest.<br><br>Please find the details of the pending training courses in the table below:</p>' + t + "<br><p>Thank you for your prompt attention to this matter.</p>"

        # send mail
        newmail.Send()

    # non DL mail
    def send_mail_main(self, val):
        f = open('email.txt', mode='r')
        email = f.read()
        email_final = email.split('\n')

        # configure outook
        ol = Dispatch('Outlook.Application')
        olmailitem = 0x0
        newmail = ol.CreateItem(olmailitem)

        # subject of the mail
        newmail.Subject = 'TerumoBCT Pending Training Details'

        # recipient of the mail
        if len(email.replace('\n', "")) > 0:
            newmail.To = ';'.join(email_final)
        else:
            newmail.To = 'shalin.shah@capgemini.com'

        # carbon copy of the mail
        # newmail.Cc = 'Upadhye, Manoj'

        # convert list to HTML table
        t = tabulate(val, headers=["Name", "Course Details", "Due Date"], tablefmt="html")
        t = t.replace("<table>", '<table cellspacing="3" cellpadding="3" border="1" bgcolor="#000000">')
        t = t.replace("<tr>", '<tr bgcolor="#ffffff">')

        newmail.HTMLBody = '<p>Hello,<br><br>I am writing to bring your attention to the training details that are apart from the DL group. Please find the details below in the table for your reference.</p>' + t + '<p>I would request you to take the necessary action as per your convenience and schedule.</p>'

        # send mail
        newmail.Send()

    def open_input_dialog_event(self):
        try:
            f = open('email.txt', mode='r')
            email = f.read()
            f.close()
            dialog = CTkInputDialog(text=f"Current email:\n{email}", title="Modify Email Address")
            email_text = dialog.get_input()

            if email_text is not None and len(email_text) > 0:
                email_text = email_text.replace(' ', '').replace("\n", "")
                email_final = email_text.split(',')

                email_flag = []
                valid_email = []
                for items in email_final:
                    if self.validate_email(items):
                        email_flag.append(True)
                        valid_email.append(items)
                    else:
                        email_flag.append(False)
                valid_email = [*set(valid_email)]
                if len(valid_email) > 0:
                    check_call(["attrib", "-H", 'email.txt'])
                    f = open('email.txt', mode='w+')
                    for items in valid_email:
                        f.write(items + '\n')
                    f.close()
                    check_call(["attrib", "+H", 'email.txt'])

                if False not in email_flag:
                    messagebox.showinfo('Email Update Message', 'Email Update Success: Emails Updated Successfully')
                else:
                    messagebox.showinfo('Email Error', "Error: Email Update Failed - Invalid Addresses Found. Unable "
                                                       "to Update All Emails.")
            else:
                messagebox.showinfo('Email Update Error', 'Email Update Failed: No Emails Were Updated')
        except Exception as e:
            messagebox.showinfo('Email Update Error', 'Email Update Failed: No Emails Were Updated')

    # def open_input_dialog_event(self):
    #     try:
    #         with open('email_hide.txt', mode='r') as f:
    #             email = f.read()
    #
    #         dialog = CTkInputDialog(text=f"Current email:\n{email}", title="Modify Email Address")
    #         email_text = dialog.get_input()
    #
    #         if email_text is not None and len(email_text.strip()):
    #             email_final = list(
    #                 set(item.strip() for item in email_text.split(',') if self.validate_email(item.strip())))
    #
    #             if email_final:
    #                 check_call(["attrib", "-H", 'email_hide.txt'])
    #                 with open('email_hide.txt', mode='w') as f:
    #                     f.write('\n'.join(email_final))
    #                 check_call(["attrib", "+H", 'email_hide.txt'])
    #                 messagebox.showinfo('Email Update Message', 'Emails Updated Successfully')
    #             else:
    #                 messagebox.showinfo('Email Error', "Invalid Email(s) Found. Unable "
    #                                                    "to Update All Email(s)")
    #     except Exception as e:
    #         messagebox.showinfo('Email Update Error', 'No Emails Were Updated')

    def validate_email(self, s):
        pat = r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,7}\b'
        if match(pat, s):
            return True
        return False

    def default_mail(self):
        if not path.isfile('email.txt'):
            default_email = 'shalin.shah@capgemini.com\n'
            f = open('email.txt', mode='w+')
            f.write(default_email)
            f.close()
            check_call(["attrib", "+H", 'email.txt'])

# run App()
if __name__ == "__main__":
    app = App()
    app.mainloop()
