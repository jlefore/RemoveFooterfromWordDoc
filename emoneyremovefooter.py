#07/20/2021 most recent. Aloha,,,,,
import os
from time import sleep
#defined parameters
#run code to test number of files. no function in order to run main. tried main from .py and would not work. 
#not working spending the time and one py file is nice. simple
filetype = ".docx" 
#dir_path = 'C:/Users/jacqu'
#on my pc had to change the brackets the other way.
dir_path = 'C:\\Users\\Jacque LeFore'
listing = os.listdir(dir_path) #files in directory path to loop thru
files_count = 0 #defines argument and starts at 0
#run_code_bollean = False #defines argument and sets to false to not open a file below. this is required as a parameter
for files in listing: #looks at every file in the directory path
    if files.endswith(filetype): #looks for the filetype defined in the path
        files_count +=1 #loops through all of the matching file types
        #print ('file #', files_count)         
        if files_count == 0:
            print("no files of the following type found ==", filetype)
        if files_count == 1:
            print("1 file of the type found ==", filetype)
            #below runs main. 
            "__main__"
        if files_count > 1:
                print("more than 1 file of the type found ==", filetype)
print("This removes footer and only used on my license cases")
#........................
def revisefooterwin32(doc):
    end = range(len(doc.Sections) + 1)
    for s in range(1,len(end)): #starting at 1 and ending at 1, the first page. Footer 1
        Footer1 = doc.Sections(s).Footers(1) #footer 1 no table
        #print("FOOTER ONE, Section #:", s)
        #print(Footer1)
        Footer1.Range.Find.Execute(FindText="by Jacque Lefore, CFP\u00AE", ReplaceWith=" ")
    for s in range(1,len(end)): #starting at 1 and ending at 1, the first page. Footer 1
        Footer2 = doc.Sections(s).Footers(2) #footer 2 none
        #print("FOOTER TWO, Section #:", s)
        #print(Footer2)

def read_open_word32(): 
    print('OPEN Word docx')
    import win32com.client as win32
    import os
    word = win32.gencache.EnsureDispatch('Word.Application')
    word.Visible = False
    for files in listing: #looks at all files like above.
        if files.endswith(filetype): #looks for the file type, 1 in this case since it passed the test.
            #reads the data in file
            file_path = os.path.join(dir_path, files)
            word.Documents.Open(file_path)
            doc = word.ActiveDocument
            return doc

def main():
    doc = read_open_word32()
    read_open_word32()
    revisefooterwin32(doc) 
    print('ok to open up the word document now')
    
if __name__ == "__main__":
    main()
    
#pywintypes.com_error: (-2147352567, 'Exception occurred.', (0, 'Microsoft Word', "Sorry, we couldn't find your file. Was it moved, renamed, or deleted?\r (C:\\//Users/Jacque%20LeFore/Mike%20and...)", 'wdmain11.chm', 24654, -2146823114), None)
## clear contents of C:\Users\<username>\AppData\Local\Temp\gen_py
#From PowerShell one can use:
#copy and paste this: Remove-Item -path $env:LOCALAPPDATA\Temp\gen_py -recurse
#python emoneyremovefooter.py