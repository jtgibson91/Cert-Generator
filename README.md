# Cert-Generator

This program automates the process of creating Certificates of Analysis (CofA) using docx-Mailmerge for Python. Mailmerge is the process of creating a template file in MS Word and filling that template with whatever data you want. Data can be entered into the template dynamically using a Python script. In my case, the data is gathered from a Gas Chromatograph (TCD detector) analyzing liquid nitrogen & other gaseous mixtures. At the present time, the two samples this program can generate CofAs for are 'liquid nitrogen' and '10% CO2 balance air mixture'.

This program does the following:
  - Takes data stored in a text (.LOG) file, processes it, and inserts it into the .docx template
  - Gathers some relevant data from the user, and inserts it into the template
  - Creates the template(s), and optionally prints the template(s) in the background, and/or displays them in Word
  - Provide a GUI for the above tasks
  

## Modules Used

Python Standard Library
  - tkinter - GUI
  - inspect - Determine caller functon from the called function

Third Party Modules
  - docx-mailmerge & lxml - Mailmerge process
  - tkcalendar - Tkinter caldendar widget
  - pywin32 - Open files in MS Word and print to the default printer
  
## Who is it for?

I made this for myself and co-workers to use at work as a QC chemists at an industrial compressed gas company. It will make my life a little easier.

## Created By

James Gibson  
github.com/jtgibson91  
jtgibson91@gmail.com
