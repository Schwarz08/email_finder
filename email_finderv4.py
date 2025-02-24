import win32com.client
import pandas as pd
import os
import re
import time as timer
import datetime

'''
Created by Jan Kyle Lewis T. Nolasco
'''

'''
msg.ReceivedTime
msg.Subject
msg.Body
msg.To
msg.Size
msg.Attachments
'''

def parse_add_kw(kw_string):
    try:
        #split string by ","
        kw_list=kw_string.split(",")
        #convert all string to lower
        lower_kw_list=[x.lower() for x in kw_list]

    except:
        lower_kw_list=[]

    return lower_kw_list

def find_email_subj_kw(email_list, folder_list, subj_kw, subj_kw_add_list, body_kw_list, msg_identifier1,
                           msg_identifier2):
    #set up win32 outlook API
    outlook = win32com.client.gencache.EnsureDispatch('Outlook.Application')
    out_namespace = outlook.GetNamespace("MAPI")

    msg_count = 0
    #loop through email list
    for email in email_list:

        #loop through folder list
        for folder in folder_list:
            #filter email and folder
            curr_folder = out_namespace.Folders[email].Folders[folder].Items

            #filter subject keyword
            filter = f"@SQL=(urn:schemas:httpmail:subject LIKE '%{subj_kw}%')"
            messages= curr_folder.Restrict(filter)

            #check if there are any filtered messages
            total_messages=len(list(messages))
            if total_messages>0:

                #loop through all messages to check for subject and body keywords
                for message in list(messages):

                    #check if message has body and convert to lower case
                    try:
                        body=message.Body.lower()
                    except:
                        continue

                    #check if subj kw list is empty
                    if len(subj_kw_add_list)>0:
                        #loop through additional subj kw list to check if all kw are present
                        for subj_kw_add in subj_kw_add_list:
                            if subj_kw_add in message.Subject.lower():
                                subj_flag=True
                            else:
                                subj_flag=False
                                break
                    else:
                        subj_flag=True

                    #check if body kw list is empty
                    if len(body_kw_list)>0:
                        #loop through additional body kw list to check if all kw are present
                        for body_kw in body_kw_list:
                            if body_kw in body:
                                body_flag=True
                            else:
                                body_flag=False
                                break
                    else:
                        body_flag=True

                    #get received time
                    try:
                        receive_time=message.ReceivedTime.strftime("%Y%m%d %H:%M:%S")
                    except:
                        receive_time="no date"

                    #check subj and body_flag, if both are true save email to file_path
                    if subj_flag and body_flag:
                        # print if subj keywords are found
                        print(f"Subject Keyword: [{subj_kw}], {subj_kw_add_list} found!")

                        # print if body keywords are found
                        if len(body_kw_list) > 0:
                            print(f"Body Keyword: {body_kw_list} found!")

                        new_folder = msg_identifier1 + "_" + msg_identifier2
                        file_path = os.path.join(os.getcwd(), "filtered_emails")
                        des_path = os.path.join(file_path, new_folder)
                        if not os.path.exists(des_path):
                            os.mkdir(des_path)
                        msg_count=msg_count+1
                        email_name = receive_time+"_"+email+"_"+msg_identifier1 + "_" + msg_identifier2 + "_"+str(msg_count)
                        print("Saving: ", email_name)
                        email_name = re.sub('[^A-Za-z0-9]+', ' ', email_name)
                        email_path = f"{des_path}\\{email_name}.msg"
                        message.SaveAs(email_path)

    if msg_count == 0:
        return False
    else:
        return True

def main():
    '''
    Replace email_list with your own
    '''
    email_list=[]

    '''
    Modify folder_list to include/exclude folders to search emails
    '''
    folder_list=["Inbox", "Sent Items"]

    dtype_dic={
        'MSG Identifier 1': str,
        'MSG Identifier 2': str,
        'Subject Keyword': str,
        'Subject Keyword (Add)': str,
        'Body Keywords': str
    }

    search=pd.read_csv("kw_list.csv", dtype=dtype_dic)
    search['found?']='not found'
    file_path=os.path.join(os.getcwd(), "filtered_emails")

    if not os.path.exists(file_path):
        os.mkdir(file_path)
    for index in search.index:
        msg_identifier1=search.loc[index, 'MSG Identifier 1']
        msg_identifier2=search.loc[index, 'MSG Identifier 2']
        subj_kw=search.loc[index, 'Subject Keyword']
        subj_kw_add_list=parse_add_kw(search.loc[index, 'Subject Keyword (Add)'])
        body_kw_list = parse_add_kw(search.loc[index, 'Body Keywords'])

        if not pd.isnull(search.loc[index, 'Subject Keyword']):
            found_email = find_email_subj_kw(email_list, folder_list, subj_kw, subj_kw_add_list, body_kw_list,
                                                 msg_identifier1, msg_identifier2)
            if found_email:
                search.loc[index, 'found?']="found"

        else:

            print("No subject keyword: ", msg_identifier1, msg_identifier2)
    search.to_csv(f"kw_list_updated.csv", index=False)

    return

if __name__ == '__main__':
    start=timer.time()
    main()
    end=timer.time()
    total_time=(end-start)/60
    print(f"Elapsed Time: {total_time} mins", )

