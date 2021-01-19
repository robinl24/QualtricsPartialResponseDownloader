from selenium import webdriver
import csv
import time
import win32com.client
import os
from tkinter import *
from tkinter.scrolledtext import ScrolledText
from tkinter import filedialog
from datetime import datetime

#Requires LoginInfo.py with login information

def clicker(Xpath):
    temp=driver.find_element_by_xpath(Xpath)
    temp.click()  

def duplicateFinder(listInput,item,usedList):
    start_at = -1
    while True:
        try:
            loc = listInput.index(item,start_at+1)            
        except ValueError:
            loc=-1
            break
        if loc not in usedList:
            usedList.append(loc)
            break
        elif loc in usedList:
            start_at = loc

            
    return loc,usedList

def newestFile(path):
    files = os.listdir(path)
    paths = [os.path.join(path, basename) for basename in files]
    return max(paths, key=os.path.getctime)

def initialStartUp():
    global driver
    driver = webdriver.Chrome()
    driver.get('https://login.qualtrics.com/login?lang=en')
    
    #remote Session
    driverUrl = driver.command_executor._url       #ie "http://127.0.0.1:60622/hub"
    #print(url)
    session_id = driver.session_id            #ie '4e167f26-dc1d-4f51-a207-f761eaf73c31'
    #print(session_id)



    driver.implicitly_wait(10)

    username= driver.find_element_by_id("UserName")
    username.send_keys(LoginInfo.username)
    password= driver.find_element_by_id('UserPassword')
    password.send_keys(LoginInfo.password)
    loginStartButton = driver.find_element_by_xpath('//*[@id="loginButton"]')
    loginStartButton.click()
    return driver,driverUrl,session_id
#Now on project page
#search for project
def openProject(iter,surveyID):
    projectSearch= driver.find_element_by_xpath('//*[@id="SurveyFoldersContainer"]/div[3]/div/div[2]/span[1]/input')
    projectSearch.send_keys(SVID[iter])
    #get the title of the survey
    Title[iter]=(driver.find_element_by_xpath("//*[@id=\"" + SVID[iter] + "\"]/td[3]/div/div/span").get_attribute("innerHTML").splitlines()[0])

    projectFind=driver.find_element_by_xpath('//*[@id=\"' + SVID[iter] + '\"]/td[3]/div/div/span')
    projectFind.click()

    #In the project


    #Go to Data & Analysis
    dataAnalysis=driver.find_element_by_xpath('//*[@id="center"]/div[2]/div/div/div/div[3]/div/ul/li[4]/a/span')
    dataAnalysis.click()


    #responses in progress section

    responsesProgress=driver.find_element_by_xpath('/html/body/div[1]/div[1]/div[2]/div[2]/div/div/div[1]/section/div[2]/div[2]/div[2]/div[1]/span')
    responsesProgress.click()

    #download the csv of Responses in Progress
    downloader=driver.find_element_by_xpath('/html/body/div[1]/div[1]/div[2]/div[2]/div/div/div[3]/div/div/div/div[1]/div/span/div')
    downloader.click()

    dirList = os.listdir(download_folder_path.get())
    oldFileNum=len(dirList)

    download=driver.find_element_by_xpath('/html/body/div[3]/ul/li[1]/a')
    download.click()

    attempts=0
    listLengthCurrent=len(dirList)
    
    while oldFileNum>=listLengthCurrent and attempts <=600:#waits up to 10 minutes
        dirList = os.listdir(download_folder_path.get())
        listLengthCurrent=len(dirList)
        attempts +=1
        time.sleep(1)
    #print(newestFile(download_folder_path.get()))
    RiP_csv_name=newestFile(download_folder_path.get())
    time.sleep(3)

    #go back to main data
    clicker('/html/body/div[1]/div[1]/div[2]/div[2]/div/div/div[1]/section/div[2]/div[2]/div[1]/div[1]/span')

    clicker('//*[@id="responsesViewTable"]/div[1]/div/div/div[2]/div/span[2]')

    clicker('/html/body/div[3]/ul/li[1]/a')

    
    #tracks new file
    dirList = os.listdir(download_folder_path.get())
    oldFileNum=len(dirList)

    #download current data
    clicker('/html/body/div[3]/div[1]/div/div/div[3]/button[1]/span')

    attempts=0
    listLengthCurrent=len(dirList)
    while oldFileNum>=listLengthCurrent and attempts <=600:#waits up to 10 minutes
        dirList = os.listdir(download_folder_path.get())
        listLengthCurrent=len(dirList)
        attempts +=1
        time.sleep(1)
    #print(newestFile(download_folder_path.get()))

    raw_csv_name=newestFile(download_folder_path.get())

    #close downloader
    clicker('/html/body/div[3]/div[1]/div/div/div[3]/button/span')
    
    with open (raw_csv_name, 'r') as csv_file:
        csv_reader_temp = csv.reader(csv_file)
        #create new import file
        with open (download_folder_path.get() + '\\' +  Title[iter] + '-template.csv', 'w',newline='') as new_file:
            csv_writer = csv.writer(new_file)

            #copy the Headers

            for x in range(0,3):
                line=next(csv_reader_temp)
                if x ==1:
                        DataHeaders=line
                        lengthRiPHeadersWrite=len(line)
                        #indexLastActivityWrite=line.index('lastActivity')

                csv_writer.writerow(line)
            # for line in csv_reader:
            #     duplicateResponseID.append(line[indexResponseIDWrite])

    #open the file with exported responses in progress data
    with open (RiP_csv_name, 'r') as csv_file:
        csv_reader = csv.reader(csv_file)
        #skip First Headers
        next(csv_reader)

        #get RiPHeaders
        RiPHeaders=next(csv_reader)
        indexLastActivityRead=RiPHeaders.index('LastActivity')        


        with open (download_folder_path.get() + '\\' +  Title[iter] + '-template.csv', 'a',newline='') as new_file:#append file to add to the RiPHeaders
            csv_writer = csv.writer(new_file)

            #list of used indices
            used=[]
            for line in csv_reader:
                importRow=[""] * lengthRiPHeadersWrite
                #dataLength=len(line)-1

                importRow[0]=line[7]#static StartDate
                importRow[1]=line[8]#static EndDate
                importRow[lengthRiPHeadersWrite-1]=line[indexLastActivityRead]#dynamic lastActivity
                importRow[lengthRiPHeadersWrite-2]=line[0] #dynamic tempID
                importRow[3]=line[5] #static IP Address

                #RiPHeaders, DataHeaders
                used=[0,1,lengthRiPHeadersWrite-1,lengthRiPHeadersWrite-2,3]
                [index, used] = duplicateFinder(DataHeaders,'gc',used)
                for x in range(0,len(line)):
                    #check if the header exists
                    #optimize to get two lists
                    [index, used]=duplicateFinder(DataHeaders,RiPHeaders[x],used)
                    if index != -1:
                        importRow[index]=line[x]

                csv_writer.writerow(importRow)







    #Upload new responses


    #open downloader
    clicker('//*[@id="responsesViewTable"]/div[1]/div/div/div[2]/div/span[2]')
    #dropdown
    ImportData=driver.find_element_by_xpath('/html/body/div[3]/ul/li[2]/a')
    ImportData.click()
    
    clicker('/html/body/div[3]/div[1]/div/div/div/div[2]/div/div[2]/div/div/span/span')
    time.sleep(1)
    shell = win32com.client.Dispatch("WScript.Shell")
    shell.SendKeys("^l", 2) #Sends CNTRL + L keyboard strokeAutomation_Testing-Responses in Progress.csv

    shell.SendKeys(download_folder_path.get(), 2) #sends the folder filepath
    shell.SendKeys("{ENTER}") #Sends ENTER key
    time.sleep(1)
    shell.SendKeys("%n",2) #Sends ALT + N keyboard stroke
    shell.SendKeys(Title[iter] + '-template.csv',2) #sends the file name you want to upload
    shell.SendKeys("{ENTER}")
    time.sleep(1)

    clicker('/html/body/div[3]/div[1]/div/div/div/div[3]/button[1]/span')
    #import responses final confirmation
    clicker('/html/body/div[3]/div[1]/div/div/div/div[3]/button[3]/span[2]')
    #future check for new responses
    time.sleep(10)
    #close importer
    clicker('/html/body/div[3]/div[1]/div/div/div/div[3]/button[2]')

    #go to projects tab
    clicker('/html/body/div[1]/div[1]/div[2]/div[1]/div/div/div[2]/div[2]/ul/li[1]/a')

def browse_button():
    # Allow user to select the downloads
    global download_folder_path
    filename = filedialog.askdirectory()
    download_folder_path.set(filename)
    
    #print(filename)

def submitList():
    
    SurveyIDList= textentry.get('1.0', 'end-1c').strip().split("\n")

    
    global SVID
    global Title

    SVID=SurveyIDList    
    Title=[""] * len(SurveyIDList)
    # print(SVID)
    # print(Title)
    # print(len(SurveyIDList))
    
    initialStartUp()
    progress=''
    for x in range(0,len(SurveyIDList)):
        #run script
        iter=x
        openProject(iter,SVID[iter])
        progress+='completed\n'
        output.delete(0.0, END)
        output.insert(END,progress)

    print(SurveyIDList)
    print(Title)



root = Tk()
root.title('Qualtrics - Response Collector')
root.configure(background='light blue')

Label (root, text="Welcome to the Partial Response Collector\nPlease enter in the survey ID(s) of the surveys you would like to track:", bg='light blue', fg='black', font='none 12 bold' ) .grid(row=1, column=0,sticky=W)

# textentry= Entry(root,width=30, bg='blue')
# textentry.grid(row=2, column=0, sticky=W)

textentry= ScrolledText(root,width=50, bg='white')
textentry.grid(row=2, column=0, sticky=W)

Button(root, text='SUBMIT', width=8, command=submitList).grid(row=5,column=0, sticky=W)

output = Text(root, width=40,height=24, bg='blue', fg='white')
output.grid(row=2, column=1, sticky=W)




download_folder_path = StringVar()

Label (root, text="Please select the location of your downloads folder:", bg='light blue', fg='black', font='none 12 bold' ) .grid(row=3, column=0,sticky=W)
lbl1 = Label(master=root,textvariable=download_folder_path)
lbl1.grid(row=4, column=0)
button2 = Button(text="Browse", command=browse_button)
button2.grid(row=4, column=1)

root.mainloop()