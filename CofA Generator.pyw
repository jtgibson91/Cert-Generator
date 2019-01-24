#from __future__ import print_function
import tkinter as tk
from tkinter import ttk
import tkinter.messagebox
import tkcalendar
from mailmerge import MailMerge
import win32api
from cylinderClasses import CO2Air, Nitrogen
from win32com import client
import inspect
from decimal import Decimal

# Test log
#RESULTSLOG = r'.\CH2.LOG'
RESULTSLOG = r'C:\peak454-64bitWin8\Calibration Result Logs\CH2.LOG'
CO2AIRTEMPLATE8 = r"C:\gtj\James\CS\Python\Cert Generator\Templates\8xCO2Air10Template.docx"
CO2AIRTEMPLATE16 = r"C:\gtj\James\CS\Python\Cert Generator\Templates\16xCO2Air10Template.docx"
N2TEMPLATE = r"C:\gtj\James\CS\Python\Cert Generator\Templates\N2Template.docx"
VERICELCERTDIRECTORY = r"C:\Users\Lab\Desktop\CERTIFICATES OF ANALYSIS\GENZYME-VERICEL\2019\\"

VERICELCO2AIRPO = "PO14975"

class Application(tk.Frame):
    def __init__(self, master=None):
        super().__init__(master)
        self.create_widgets()
        self.grid()

    def create_widgets(self):
        self.nb = ttk.Notebook(root)
        self.nb.grid(row=1, column=0, columnspan=50, rowspan=49, sticky="NESW")

        # Create a notebook tab for 10% CO2-Air
        self.co2AirTab = ttk.Frame(self.nb)
        self.nb.add(self.co2AirTab, text="10% CO2 Air")
        # Create a notebook tab for Liquid Medical Nitrogen
        self.n2Tab = ttk.Frame(self.nb)
        self.nb.add(self.n2Tab, text="Nitrogen")

        # Create 3 frames to break up co2AirTab
        self.top = ttk.Frame(self.co2AirTab, height=60, width=300)
        self.top.pack(side="top")

        # Populate the grid
        rows = 0
        while rows < 30:
            self.top.rowconfigure(rows, weight=1)
            self.top.columnconfigure(rows, weight=1)
            rows += 1

        # Separator for the top section of the top frame
        self.separatorTop = ttk.Frame(self.top, height=40, width=300)
        self.separatorTop.grid(row=0, column=0, rowspan=10, columnspan=30)

        # Left frame
        self.left = ttk.Frame(self.co2AirTab, height=70, width=180)
        self.left.pack(side="left")

        # Populate the left frame grid
        rows = 0
        while rows < 30:
            self.left.rowconfigure(rows, weight=1)
            self.left.columnconfigure(rows, weight=1)
            rows += 1

        # Right frame
        self.right = ttk.Frame(self.co2AirTab, height=70, width=180)
        self.right.pack(side="right")

        # Populate the right frame grid
        rows = 0
        while rows < 30:
            self.right.rowconfigure(rows, weight=1)
            self.right.columnconfigure(rows, weight=1)
            rows += 1

        # Separator for the right side of the right frame
        self.separatorR = ttk.Frame(self.right, height=100, width=40)
        self.separatorR.grid(row=0, column=28, rowspan=30, columnspan=2)

        # Bottom frame
        self.bottom = ttk.Frame(self.co2AirTab, height=100, width=300)
        self.bottom.pack(side="bottom")
        self.bottom.pack_propagate(0)

        # Populate the bottom frame grid
        rows = 0
        while rows < 30:
            self.bottom.rowconfigure(rows, weight=1)
            self.bottom.columnconfigure(rows, weight=1)
            rows += 1

        # Separator for the bottom section of the bottom frame
        self.separatorBot = ttk.Frame(self.bottom, height=3, width=300)
        self.separatorBot.grid(row=30, column=0, rowspan=10, columnspan=30)
        self.separatorBot.grid_propagate(0)

        # Client
        self.client = ttk.Combobox(self.top, values=["Vericel"], width=7)
        self.client.set("Vericel")
        self.client.grid(row=5, column=30)

        # Number of cylinders
        self.numCylsLabel = ttk.Label(self.left, text="Number of cylinders", width=23, anchor='e')
        self.numCylsLabel.grid(row=0, column=0)

        self.numCyls = ttk.Combobox(self.right, values=[8, 16], width=2)
        # If the results log has CO2Air cylinder data in it currently, then pre-set the value to this comboxbox
        if self.gas_type_in_results_log() == "CO2Air10":
            self.numCyls.set(self.num_cyls_in_results_log())
        self.numCyls.grid(row=0, column=2, sticky='ne')

        # Separators...
        # *Note: if variable name has a 'b' at the end, it's in the right frame
        self.separator = ttk.Frame(self.left, height=10, width=100)
        self.separator.grid(row=2, column=0)
        self.separator.grid_propagate(0)

        self.separatorb = ttk.Frame(self.right, height=10, width=100)
        self.separatorb.grid(row=2, column=2)
        self.separatorb.grid_propagate(0)

        # Delivery date
        self.deliveryDateLabel = ttk.Label(self.left, text="Delivery Date", width=23, anchor='e')
        self.deliveryDateLabel.grid(row=5, column=0)
        # This is pop-up calendar from the tkcalendar module
        self.cal = tkcalendar.DateEntry(self.right, width=10)
        self.cal.grid(row=5, column=2, sticky='e')

        # Separators
        self.separator2 = ttk.Frame(self.left, height=10, width=100)
        self.separator2.grid(row=7, column=0)
        self.separator2.grid_propagate(0)

        self.separator2b = ttk.Frame(self.right, height=10, width=100)
        self.separator2b.grid(row=7, column=2)
        self.separator2b.grid_propagate(0)

        # Invoice number
        self.invoiceLabel = ttk.Label(self.left, text="Invoice #", width=23, anchor='e')
        self.invoiceLabel.grid(row=10, column=0)

        self.invoice = ttk.Entry(self.right, width=7)
        self.invoice.grid(row=10, column=2, sticky='e')

        # Separators
        self.separator3 = ttk.Frame(self.left, height=10, width=100)
        self.separator3.grid(row=12, column=0)
        self.separator3.grid_propagate(0)

        self.separator3b = ttk.Frame(self.right, height=10, width=100)
        self.separator3b.grid(row=12, column=2)
        self.separator3b.grid_propagate(0)

        # PO number
        self.POLabel = ttk.Label(self.left, text="PO #", width=23, anchor='e')
        self.POLabel.grid(row=15, column=0)

        self.PO = ttk.Combobox(self.right, values=[VERICELCO2AIRPO], width=8)
        self.PO.set(VERICELCO2AIRPO)
        self.PO.grid(row=15, column=2, sticky='e')

        # Separators
        self.separator4 = ttk.Frame(self.left, height=10, width=100)
        self.separator4.grid(row=17, column=0)
        self.separator4.grid_propagate(0)

        self.separator4b = ttk.Frame(self.right, height=10, width=100)
        self.separator4b.grid(row=17, column=2)
        self.separator4b.grid_propagate(0)

        # Operator (analyst)
        # This label was not aligned w/ the others so the width is set to 20 instead of 23. I believe it is b/c of the separator frames
        self.operatorLabel = tk.Label(self.left, text="Analyst", width=20, anchor='e')
        self.operatorLabel.grid(row=20, column=0)

        self.operator = ttk.Combobox(self.right, values=["James Gibson", "Minerva Rivas"], width=12)
        self.operator.set("James Gibson")
        self.operator.grid(row=20, column=2, sticky='e')

        # Separates self.operatorLabel from self.generateCofA
        self.separator5 = ttk.Frame(self.left, height=30, width=100)
        self.separator5.grid(row=22, column=0)
        self.separator5.grid_propagate(0)

        self.separator5b = ttk.Frame(self.right, height=10, width=100)
        self.separator5b.grid(row=22, rowspan=2, column=2)
        self.separator5b.grid_propagate(0)

        self.separator6b = ttk.Frame(self.right, height=43, width=100)
        self.separator6b.grid(row=27, rowspan=5, column=2)
        self.separator6b.grid_propagate(0)

        # Generate CofA button
        self.generateCofA = ttk.Button(self.left, text="Generate Cert", command=self.generate_co2Air_cert)
        self.generateCofA.grid(row=26, column=0, rowspan=2)

        # 'Print' option
        self.printVar = tk.IntVar()
        self.print = ttk.Checkbutton(self.bottom, variable=self.printVar, text="Print")
        self.print.grid(row=27, column=2, sticky='sw')

        # 'Open in Word' option
        self.openInWordVar = tk.IntVar()
        self.openInWord = ttk.Checkbutton(self.bottom, variable=self.openInWordVar, text="Open in Word")
        self.openInWord.grid(row=26, column=2, sticky='sw')

        ##################################
        ##                              ##
        #@        Nitrogen Tab          #@
        ##                              ##
        ##################################

        # Create 3 frames to break up co2AirTab
        self.Ntop = ttk.Frame(self.n2Tab, height=60, width=300)
        self.Ntop.pack(side="top")

        # Populate the top frame grid
        rows = 0
        while rows < 30:
            self.Ntop.rowconfigure(rows, weight=1)
            self.Ntop.columnconfigure(rows, weight=1)
            rows += 1

        # Separator for the top section of the top frame
        self.NseparatorTop = ttk.Frame(self.Ntop, height=40, width=300)
        self.NseparatorTop.grid(row=0, column=0, rowspan=10, columnspan=30)

        # Left frame
        self.Nleft = ttk.Frame(self.n2Tab, height=70, width=180)
        self.Nleft.pack(side="left")

        # Populate the left frame grid
        rows = 0
        while rows < 30:
            self.Nleft.rowconfigure(rows, weight=1)
            self.Nleft.columnconfigure(rows, weight=1)
            rows += 1

        # Right frame
        self.Nright = ttk.Frame(self.n2Tab, height=70, width=180)
        self.Nright.pack(side="right")

        # Populate the right frame grid
        rows = 0
        while rows < 30:
            self.Nright.rowconfigure(rows, weight=1)
            self.Nright.columnconfigure(rows, weight=1)
            rows += 1

        # Separator for the right side of the right frame
        self.NseparatorR = ttk.Frame(self.Nright, height=100, width=40)
        self.NseparatorR.grid(row=0, column=28, rowspan=30, columnspan=2)

        # Bottom frame
        self.Nbottom = ttk.Frame(self.n2Tab, height=100, width=300)
        self.Nbottom.pack(side="bottom")

        # Populate the bottom frame grid
        rows = 0
        while rows < 30:
            self.Nbottom.rowconfigure(rows, weight=1)
            self.Nbottom.columnconfigure(rows, weight=1)
            rows += 1

        # Separator for the bottom section of the bottom frame
        self.NseparatorBot = ttk.Frame(self.Nbottom, height=3, width=300)
        self.NseparatorBot.grid(row=30, column=0, rowspan=10, columnspan=30)
        self.NseparatorBot.grid_propagate(0)

        # Client
        self.Nclient = ttk.Combobox(self.Ntop, values=["Vericel"], width=7)
        self.Nclient.set("Vericel")
        self.Nclient.grid(row=5, column=30)

        # Number of cylinders
        self.NnumCylsLabel = ttk.Label(self.Nleft, text="Number of cylinders", width=23, anchor='e')
        self.NnumCylsLabel.grid(row=0, column=0)

        self.NnumCyls = ttk.Combobox(self.Nright, values=[1, 2, 3, 4, 5, 6, 7, 8, 9, 10], width=2)
        # If the type of cylinder in results log is N2, then pre-set this combobox
        if self.gas_type_in_results_log() == "N2":
            self.NnumCyls.set(self.num_cyls_in_results_log())
        self.NnumCyls.grid(row=0, column=2, sticky='ne')

        # Separators...
        # *Note: if variable name has a 'b' at the end, it's in the right frame
        self.Nseparator = ttk.Frame(self.Nleft, height=10, width=100)
        self.Nseparator.grid(row=2, column=0)
        self.Nseparator.grid_propagate(0)

        self.Nseparatorb = ttk.Frame(self.Nright, height=10, width=100)
        self.Nseparatorb.grid(row=2, column=2)
        self.Nseparatorb.grid_propagate(0)

        # Delivery date
        self.NdeliveryDateLabel = ttk.Label(self.Nleft, text="Delivery Date", width=23, anchor='e')
        self.NdeliveryDateLabel.grid(row=5, column=0)

        self.Ncal = tkcalendar.DateEntry(self.Nright, width=10)
        self.Ncal.grid(row=5, column=2, sticky='e')

        # Separators
        self.Nseparator2 = ttk.Frame(self.Nleft, height=10, width=100)
        self.Nseparator2.grid(row=7, column=0)
        self.Nseparator2.grid_propagate(0)

        self.Nseparator2b = ttk.Frame(self.Nright, height=10, width=100)
        self.Nseparator2b.grid(row=7, column=2)
        self.Nseparator2b.grid_propagate(0)

        # Invoice number
        self.NinvoiceLabel = ttk.Label(self.Nleft, text="Invoice #", width=23, anchor='e')
        self.NinvoiceLabel.grid(row=10, column=0)

        self.Ninvoice = ttk.Entry(self.Nright, width=7)
        self.Ninvoice.grid(row=10, column=2, sticky='e')

        # Separators
        self.Nseparator3 = ttk.Frame(self.Nleft, height=10, width=100)
        self.Nseparator3.grid(row=12, column=0)
        self.Nseparator3.grid_propagate(0)

        self.Nseparator3b = ttk.Frame(self.Nright, height=10, width=100)
        self.Nseparator3b.grid(row=12, column=2)
        self.Nseparator3b.grid_propagate(0)

        # Carbon monoxide
        self.NCOLabel = ttk.Label(self.Nleft, text="Carbon monoxide (ppm)", width=23, anchor='e')
        self.NCOLabel.grid(row=15, column=0)

        self.NCO = ttk.Combobox(self.Nright, values=["ND"], width=3)
        self.NCO.set("ND")
        self.NCO.grid(row=15, column=2, sticky='e')

        # Separators
        self.Nseparator4 = ttk.Frame(self.Nleft, height=10, width=100)
        self.Nseparator4.grid(row=17, column=0)
        self.Nseparator4.grid_propagate(0)

        self.Nseparator4b = ttk.Frame(self.Nright, height=10, width=100)
        self.Nseparator4b.grid(row=17, column=2)
        self.Nseparator4b.grid_propagate(0)

        # Operator (analyst)
        self.NoperatorLabel = tk.Label(self.Nleft, text="Analyst", width=20, anchor='e')
        self.NoperatorLabel.grid(row=20, column=0)

        self.Noperator = ttk.Combobox(self.Nright, values=["James Gibson", "Minerva Rivas"], width=12)
        self.Noperator.set("James Gibson")
        self.Noperator.grid(row=20, column=2, sticky='e')

        # Separates self.operatorLabel from self.generateCofA
        self.Nseparator5 = ttk.Frame(self.Nleft, height=30, width=100)
        self.Nseparator5.grid(row=22, column=0)
        self.Nseparator5.grid_propagate(0)

        self.Nseparator5b = ttk.Frame(self.Nright, height=10, width=100)
        self.Nseparator5b.grid(row=22, rowspan=2, column=2)
        self.Nseparator5b.grid_propagate(0)

        self.Nseparator6b = ttk.Frame(self.Nright, height=43, width=100)
        self.Nseparator6b.grid(row=27, rowspan=5, column=2)
        self.Nseparator6b.grid_propagate(0)

        # Generate CofA button
        self.NgenerateCofA = ttk.Button(self.Nleft, text="Generate Certs", command=self.generate_n2_cert)
        self.NgenerateCofA.grid(row=26, column=0, rowspan=2)

        # 'Print' option
        self.NprintVar = tk.IntVar()
        self.Nprint = ttk.Checkbutton(self.Nleft, variable=self.NprintVar, text="Print")
        self.Nprint.grid(row=27, column=1, sticky='w')


    def generate_co2Air_cert(self):
        """ Creates the CO2-Air certificate of analysis """

        # If results log does not have CO2Air in it, show user an error and return 1
        if self.gas_type_in_results_log() != "CO2Air10":
            tk.messagebox.showerror(
                    "Data Problem",
                    f"Your results log file has {self.gas_type_in_results_log()} data in it."
                )
            return 1

        if self.numCyls.get() == '8' or '16':

            # List to hold CO2Air objects
            cylinderItems = []

            # Open Peaksimple results log file. Read data into a list, resultsLog
            file = open(RESULTSLOG)
            resultsLog = file.readlines()

            # Remove trailing newlines (not sure why there are so many)
            while "\n" in resultsLog:
                resultsLog.remove("\n")

            # We now have a list of n items where n is number of cylinders analyzed

            # Extract cylinder data into individual CO2Air obj's and add them to a list
            for items in resultsLog:
                items = items.split()
                newCylItem = CO2Air(items[5], items[6], self.round_to_3(items[11]),
                                    self.round_to_3(items[17]), self.round_to_3(items[23]))
                cylinderItems.append(newCylItem)

            # MS word template document
            if self.numCyls.get() == '8':
                template = CO2AIRTEMPLATE8
            elif self.numCyls.get() == '16':
                template = CO2AIRTEMPLATE16

            # Create MS word MailMerge document
            document = MailMerge(template)

            # Get the proper date format
            self.date = self.reformat_time(self.cal.get_date())

            # Only need to enter these fields into the Word document one time. So merge them seperately first
            document.merge(
                PO=self.PO.get(),
                OPERATOR=self.operator.get(),
                INVOICE=self.invoice.get(),
                DATE=self.date,
                LOT=cylinderItems[0].Lot.strip('\"')
            )

            try:
                # Merge values into fields
                document.merge(
                    SN0=cylinderItems[0].SN.strip('\"'),
                    SN1=cylinderItems[1].SN.strip('\"'),
                    SN2=cylinderItems[2].SN.strip('\"'),
                    SN3=cylinderItems[3].SN.strip('\"'),
                    SN4=cylinderItems[4].SN.strip('\"'),
                    SN5=cylinderItems[5].SN.strip('\"'),
                    SN6=cylinderItems[6].SN.strip('\"'),
                    SN7=cylinderItems[7].SN.strip('\"'),
                    CO20=cylinderItems[0].CO2,
                    CO21=cylinderItems[1].CO2,
                    CO22=cylinderItems[2].CO2,
                    CO23=cylinderItems[3].CO2,
                    CO24=cylinderItems[4].CO2,
                    CO25=cylinderItems[5].CO2,
                    CO26=cylinderItems[6].CO2,
                    CO27=cylinderItems[7].CO2,
                    O20=cylinderItems[0].O2,
                    O21=cylinderItems[1].O2,
                    O22=cylinderItems[2].O2,
                    O23=cylinderItems[3].O2,
                    O24=cylinderItems[4].O2,
                    O25=cylinderItems[5].O2,
                    O26=cylinderItems[6].O2,
                    O27=cylinderItems[7].O2,
                )
            # Catch all errors. Most likely error is index error from not enough cylinder data in .LOG file
            except:
                tk.messagebox.showerror(
                    "Data Problem",
                    "Your results log file is incomplete/ missing."
                )

            if self.numCyls.get() == '16':
                try:
                    # Merge values into fields
                    document.merge(
                        SN8=cylinderItems[8].SN.strip('\"'),
                        SN9=cylinderItems[9].SN.strip('\"'),
                        SN10=cylinderItems[10].SN.strip('\"'),
                        SN11=cylinderItems[11].SN.strip('\"'),
                        SN12=cylinderItems[12].SN.strip('\"'),
                        SN13=cylinderItems[13].SN.strip('\"'),
                        SN14=cylinderItems[14].SN.strip('\"'),
                        SN15=cylinderItems[15].SN.strip('\"'),
                        CO28=cylinderItems[8].CO2,
                        CO29=cylinderItems[9].CO2,
                        CO210=cylinderItems[10].CO2,
                        CO211=cylinderItems[11].CO2,
                        CO212=cylinderItems[12].CO2,
                        CO213=cylinderItems[13].CO2,
                        CO214=cylinderItems[14].CO2,
                        CO215=cylinderItems[15].CO2,
                        O28=cylinderItems[8].O2,
                        O29=cylinderItems[9].O2,
                        O210=cylinderItems[10].O2,
                        O211=cylinderItems[11].O2,
                        O212=cylinderItems[12].O2,
                        O213=cylinderItems[13].O2,
                        O214=cylinderItems[14].O2,
                        O215=cylinderItems[15].O2,
                    )
                # Catch all errors. Most likely error is index error from not enough cylinder data in .LOG file
                except:
                    tk.messagebox.showerror(
                        "Data Problem",
                        "Your results log file is incomplete/ missing."
                    )

            # Generate filename
            self.filename = self.generate_filename() + ".docx"


            # Save document
            document.write(VERICELCERTDIRECTORY + self.filename)
            if self.openInWordVar.get() == 0 and self.printVar.get() == 0:
                # Report success
                tk.messagebox.showinfo(
                "Success",
                "%s created." % self.filename
                )

            # If user has checked 'Open in Word' checkbutton, then launch file in MS Word using
            # pywin32 (win32 api)
            if self.openInWordVar.get() == 1:
                win32api.ShellExecute(0, 'open', VERICELCERTDIRECTORY + self.filename, '', '', 1)


            # If user has checked 'Print' checkbutton, then print the cert
            # At work this will be replaced by the path to Vericel Certs
            if self.printVar.get() == 1:
                self.print_word_document(VERICELCERTDIRECTORY + self.filename)

    def generate_n2_cert(self):
        """ Generate CofA for Nitrogen """

        # List containing all the certs created this round
        listOfCertsCreated = []

        # If results log does not have N2 in it, show user an error and return 1
        if self.gas_type_in_results_log() != "N2":
            tk.messagebox.showerror(
                "Data Problem",
                f"Your results log file has {self.gas_type_in_results_log()} data in it."
            )
            return 1

        # Mailmerge-docx MS Word template for nitrogen
        template = N2TEMPLATE

        # List to hold Nitrogen objects
        cylinderItems = []

        # Open Peaksimple results log file. Read data into a list, resultsLog
        file = open(RESULTSLOG)
        resultsLog = file.readlines()

        # Remove trailing newlines
        while "\n" in resultsLog:
            resultsLog.remove("\n")

        # Create Nitrogen objects and add them to a list, cylinderItems
        for items in resultsLog:
            items = items.split()
            # If the sample had O2 detected by GC. items[16] = N2, items[10] = O2. Both should round to 2 decimal places
            if items[7] == "\"Oxygen\"":
                newN2Obj = Nitrogen(items[5], items[6], str(round(float(items[16]), 2)), str(round(float(items[10]), 2)) + ' %')
            # If the sample had no O2 detected by GC
            if items[7] == "\"Nitrogen\"":
                newN2Obj = Nitrogen(items[5], items[6], str(100))
            cylinderItems.append(newN2Obj)

        # Get the proper date format
        self.Ndate = self.reformat_time(self.Ncal.get_date())

        # Create one CofA for each Nitrogen cylinder object in cylinderItems
        # i = counter for creating enumerated filenames
        i = 1
        for items in cylinderItems:

            # Create MS word MailMerge document
            document = MailMerge(template)

            document.merge(
                OPERATOR=self.Noperator.get(),
                INVOICE=self.Ninvoice.get(),
                DATE=self.Ndate,
                SN=items.SN.strip('\"'),
                LOT=items.Lot.strip('\"'),
                N2=items.N2.strip('\"'),
                O2=items.O2.strip('\"'),
                CO=items.CO.strip('\"')

            )

            # Generate filename for the cert and add it to a list, listOfCertsCreated
            filename = self.generate_filename(i) + ".docx"
            listOfCertsCreated.append(filename)

            # Save document
            # At work it will be: "Vericel CofA path\" + self.filename
            document.write(VERICELCERTDIRECTORY + filename)

            i += 1

        # Print the certs  
        if self.NprintVar.get() == 1:
            for items in listOfCertsCreated:
                self.print_word_document(VERICELCERTDIRECTORY + items)
                    
    def generate_filename(self, counter=0):
        """ Generate filename for the CofA """
        
        # Get date in the right format: MM-DD-YYYY
        self.date = self.reformat_time(self.cal.get_date())

        # Use inspect module to determine caller function. Which function called this one will determine the filename.
        # curframe = this frame (called function)
        curframe = inspect.currentframe()
        # calframe = previous frame (caller function)
        calframe = inspect.getouterframes(curframe, 2)

        # Caller function = generate_co2Air_cert()
        if calframe[1][3] == "generate_co2Air_cert":
            self.filename = f"Genzyme CO2-Air {self.date}"

        # Caller function = generate_n2_cert()
        elif calframe[1][3] == "generate_n2_cert":
            self.filename = f"Genzyme LIQ N2 {self.Ndate}_{counter}"

        return self.filename

    def reformat_time(self, date):
        """ Reformats time from YYYY-MM-DD to MM-DD-YYYY """
        
        newDate = str(date)
        year = newDate[0:4]
        newDate = newDate[5:] + '-' + year
        
        return newDate

    def num_cyls_in_results_log(self):
        """ Returns the number of lines (cylinders) in the Peaksimple results log file """
        
        # Open PeakSimple results log file. Read data into a list, resultsLog
        file = open(RESULTSLOG)
        resultsLog = file.readlines()

        # Remove trailing newlines
        while "\n" in resultsLog:
            resultsLog.remove("\n")

        # The list has one item for each cylinder analyzed
        numCyls = len(resultsLog)

        file.close()

        return numCyls

    def gas_type_in_results_log(self):
        """ Returns a string describing the type of gas presently in the results log file.
        Options are: CO2Air10, CO2Air5, N2.
            Returns: String """

        # Open PeakSimple results log file. Read data into a list, resultsLog
        file = open(RESULTSLOG)
        resultsLog = file.readlines()

        # Remove trailing newlines
        while "\n" in resultsLog:
            resultsLog.remove("\n")

        # Get one cylinder entry and turn it into a list
        oneCyl = resultsLog[0].split("\t")

        # 10% CO2-Air ("CO2Air10")
        if oneCyl[6] == "\"Carbon Dioxide\"" and oneCyl[12] == "\"Oxygen\"" and oneCyl[18] == "\"Nitrogen\"" and float(oneCyl[9]) > 9 and float(oneCyl[9]) < 11:
            return "CO2Air10"

        # 5% CO2-Air ("CO2Air5")
        elif oneCyl[6] == "\"Carbon Dioxide\"" and oneCyl[12] == "\"Oxygen\"" and oneCyl[18] == "\"Nitrogen\"" and float(oneCyl[9]) > 4 and float(oneCyl[9]) < 6:
            return "CO2Air5"

        # Nitrogen (in case where O2 is detected by GC)
        elif oneCyl[6] == "\"Oxygen\"" and oneCyl[12] == "\"Nitrogen\"" and float(oneCyl[9]) < 1 and float(oneCyl[15]) > 99:
            return "N2"

        # Nitrogen (in case where no O2 is detected by GC)
        elif oneCyl[6] == "\"Nitrogen\"" and float(oneCyl[9]) == 100:
            return "N2"

    def print_word_document(self, filename):
        """ Uses COM wrappers from pywin32 to print the file to the default printer, in the background (w/o opening Word) """
    
        word = client.Dispatch("Word.Application")
        word.Documents.Open(filename)
        # The first argument is Background. Takes True or False. True prints in background (w/o opening Word)
        word.ActiveDocument.PrintOut(True)
        word.Quit()

    def round_to_3(self, number):
        """ Rounds to 3 decimal places (not 3 and SOMETIMES 2). In other words it doesn't cut off zeros.
        Input: number - a numerical string (10-99) to be rounded from 4 places after the decimal down to 3.
        Output: newNumber - a numerical string rounded to 3 places after the decimal point.
        """
        # Must convert 'number' from string to a Decimal
        amount = Decimal(number)
	
        return str(amount.quantize(Decimal("0.001")))




root = tk.Tk()
root.title("CofA Generator")
root.iconbitmap(r'C:\gtj\James\CS\Python\Cert Generator\cert.ico')
app = Application(master=root)
app.master.geometry("450x300")

rows = 0
while rows < 50:
    root.rowconfigure(rows, weight=1)
    root.columnconfigure(rows, weight=1)
    rows += 1

app.mainloop()
