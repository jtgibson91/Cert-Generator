# Cert-Generator

This program automates the process of creating Certificates of Anlysis using docx-Mailmerge for Python. Mailmerge is a process of creating a template file in MS Word and filling that template with whatever data you want, using a Python script. 

This program does the following:
  - Takes data stored in a text (.LOG) file, processes it, and inserts it into the .docx template
  - Gathers some relevant data from the user, and inserts it into the template
  - Prints the template(s) in the background, displays them in Word, or simply creates them
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

I made this for myself to use at work as a QC chemist at an industrial compressed gas company. It will make my life a little easier.

## Created By

James Gibson  
github.com/jtgibson91  
jtgibson91@gmail.com
