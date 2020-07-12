#GmailAPI
#Importing the libraries we want

import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import email
import base64
import shutil

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/gmail.readonly']

def get_service():
    
    """ Checks the credentials and return the service object which is be used for other tasks """   
    
    print("CHECKING CREDENTIALS...")
    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build('gmail', 'v1', credentials=creds)
    print("CREDENTIALS CHECKED,OK!")
    
    return service


def search_messages(service,user_id,search_string):
    
    """ Searches the message by using the search string and returns the message ids """
    
    try:
        print("SEARCHING MESSAGES...")
        search_id=service.users().messages().list(userId=user_id,q=search_string).execute()
        number_of_messages=search_id['resultSizeEstimate']
        if (number_of_messages!=0):
            message_ids=search_id['messages']
            ids=[]
            for message_id in message_ids:
                ids.append(message_id['id'])
            print("MESSAGE IDS FOUND! PRINTING IDS:{}".format(ids))    
            return ids
        else:
            print('There is no search results for the search string.Returning a empty string')
            return ''
        
    except Exception as e:
        print ('An error occurred:',e) 
    
        
def get_messages(service,user_id,message_id):
    
    """Returns the body of the message"""
    
    try:
        print("GETTING THE MESSAGE...")
        msg_raw = service.users().messages().get(userId=user_id,id=message_id,format='raw').execute()
        msg_bytes = base64.urlsafe_b64decode(msg_raw['raw'].encode('ASCII')) 
        msg = email.message_from_bytes(msg_bytes)  
        content_type = msg.get_content_maintype()
        if (content_type == 'multipart'):
            print("THE MESSAGE BODY CONTAINS:")
            #part 1 is plain text ,part 2 is html text
            if (content_type == 'multipart'):
                part1,part2= msg.get_payload()
                if (part1.get_content_maintype() == 'multipart'):
                    c1,c2=part1.get_payload()
                    print(c1.get_payload())
                    return(c1.get_payload())
                else:
                    print(part1.get_payload())
                    return(part1.get_payload())
        else:
            print(msg.get_payload())
            return msg.get_payload()
       
    except Exception as e:
            print ('An error occurred:',e) 
    
    
def get_attachments(service,user_id,message_id):
    
    """Downloads the attachment and save it to the desired folder"""
    
    try:
        msg = service.users().messages().get(userId=user_id,id=message_id).execute()
        for part in msg['payload']['parts']:
            if part['filename']:
                if 'data' in part['body']:
                    data = part['body']['data']
                else:
                    attachment_id = part['body']['attachmentId']
                    attachment = service.users().messages().attachments().get(userId=user_id, messageId=message_id,id=attachment_id).execute()
                    data = attachment['data']
                    file_data = base64.urlsafe_b64decode(data.encode('UTF-8'))
                    file_name = part['filename']

                    with open(file_name, 'wb') as f:
                        f.write(file_data)
                        
                    destination_path = os.getcwd()+'\\'+'Downloaded_Attachments' +'\\' + file_name
                    destination_dir=os.getcwd()+'\\'+'Downloaded_Attachments'
                    while True:
                        if os.path.exists(destination_dir):
                            shutil.move(file_name,destination_path)
                            break
                        else:
                            os.mkdir('Downloaded_Attachments')
                    print("THE ATTACHMENT IS STORED IN THE LOCATION "+destination_path)
                    return destination_path
                
    except Exception as e:
        print ('An error occurred:',e) 

#SELENIUM 
#importing the libraries needed

from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

#Type in your location to the chrome driver
chrome_path=''

#This dictionary contains the name of the element in the webpage along with its index for corresponding data 
html_fields={'Name':('entry.749278213',0),"Father's name":('entry.321721136',1),"Residential address":('entry.1785907816',3),"Father's phone number":('entry.568321194',4),'English mark ':('entry.789531703',6),'Tamil mark':('entry.46651444',7),'Mathematics mark':('entry.948760241',8),'Science mark':('entry.1803145887',9),'Social Science mark':('entry.152129990',10)}

def fill_form(file_path):
    
    """Automates the data entry into the webform by providing the filepath that contains the data to be entered """
    
    #opens chrome
    driver=webdriver.Chrome(chrome_path)
    #goto the provided webform
    driver.get('https://forms.gle/E84ANMvfkyoaK7Bx8')
    print("PAGE LOADED")
    with open (file_path,'r') as fr:
        #reading the file that contains the required data
        csv_fields=fr.readline().strip().split(',') 
        contents=fr.readlines()
        print("FILE READ")
        i=0
        for row in contents:
            i+=1
            #filling the form
            data=row.strip().split(',')
            total_iterations=len(list(html_fields.keys()))
            print("FILLING FORM:",i)
            for index in range(total_iterations):
                name=list(html_fields.values())[index][0]
                search=driver.find_element_by_name(name)
                data_index=list(html_fields.values())[index][1]
                search.send_keys(data[data_index])
            #clicking on the submit button
            submit=driver.find_element_by_xpath('//*[@id="mG61Hd"]/div/div/div[3]/div[1]/div/div/span/span')
            submit.click()
            #waiting till the next page show up
            try:
                wait = WebDriverWait(driver, 10)
                #clicking the submit another response button
                another_response = wait.until(EC.element_to_be_clickable((By.XPATH,'/html/body/div[1]/div[2]/div[1]/div/div[4]/a')))            
                another_response.click()
                print("FORM {} COMPLETED".format(i))
            except:
                #if the next page isn't get loaded ,print the message
                print('Error')
    print("COMPLETED FILLING.DONE!")
        
def automate():
    
    """The complete workflow from downloading the attachment to filling the webform"""
    
    #change the directory where the credentials are stored 
    os.chdir('')
    #user_id and search_string are provided
    user_id='me'
    search_string='filename:csv'
    service=get_service()
    ids=search_messages(service,user_id,search_string)
    #passing the recent message_id
    message=get_messages(service,user_id,ids[0])
    loc=get_attachments(service,user_id,ids[0])
    fill_form(loc)


#Let's do this
automate()