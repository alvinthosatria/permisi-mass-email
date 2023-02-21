import pandas as pd
import yagmail
import time
'''
PLEASE REMEMBER TO pip install yagmail
'''
 
try:
    #change the excel file according to the one that is already sorted
    '''
    FILES
    - send_success_workshop_weave.xlsx
    - send_success_workshop_damar.xlsx
    - send_success_workshop_batik.xlsx
    - send_success_workshop_arjuna.xlsx
    - send_success_app_booth1.xlsx
    - send_success_app_booth2.xlsx
    - send_fail_message_workshop1.xlsx
    - send_fail_workshop2.xlsx
    - send_fail_message_booth1.xlsx
    - send_fail_message_booth2.xlsx
    '''
    df = pd.read_excel('change to one of the template file above')

    data = []
    for i in range(df.shape[0]):
        data.append(df.iloc[i].tolist())

    messages = []
    email_list = []
    counter = 0
    #print the mass email messages for each individual
    '''
    EMAIL TEMPLATE FOR FAIL BOOTH
    Dear [name],

    Thank you very much for your registration for the captioned event. Due to the overwhelming response, [this session] is at full capacity. Please find the registration form below to check if another session is available.

    https://permisi.hk/permisi-indofest-2023/ 

    We would like to express our utmost gratitude for your interest in joining this event. If there is no other session available, please still feel free to come to the designated venue for an on-the-spot visit if a slot becomes available!

    Kind regards,
    PERMISI

    '''
    '''
    EMAIL TEMPLATE FOR FAIL WORKSHOP
    Dear [name],

    Thank you very much for your registration for the captioned event. Due to the overwhelming response, [this workshop session] is at full capacity. However, please still feel free to come to the designated venue for an on-the-spot visit if a slot becomes available!

    We would like to express our utmost gratitude for your interest in joining this event. 

    Kind regards,
    PERMISI

    '''
    for person in data:
        messages.append(f'Dear {person[1]},\n\nThank you very much for your registration for the captioned event. Due to the overwhelming response, {person[3]} is at full capacity. Please find the registration form below to check if another session is available.\n\nhttps://permisi.hk/permisi-indofest-2023/ \n\nWe would like to express our utmost gratitude for your interest in joining this event. If there is no other session available, please still feel free to come to the designated venue for an on-the-spot visit if a slot becomes available!\n\nKind regards,\nPERMISI\n')
        #messages.append(f'Dear {person[1]},\nThank you very much for your registration for the captioned event. Due to the overwhelming response, {person[3]} is at full capacity. Please find the registration form below to check if another session is available.\n\nBest,\n\nIndoFest23 Team\n')
        #add more template
        email_list.append(person[2])
        '''
        person[0] : date
        person[1] : name
        person[2] : email
        person[3] : type of activity (booth or workshop) look at the excel file names
        '''
        
    # initiating connection with SMTP server
    yag = yagmail.SMTP("email@gmail.com","app password") #change to permisi email and make an app password from gmail
    
    # Adding Content and sending it
    for content, email in zip(messages, email_list):
        yag.send(email, #recipient email:  (BECAREFUL BEFORE RUNNING PROGRAM)
                "Subject", #Subject
                content) #Content from each mail is different
        counter += 1
        if not counter % 10: #prevent API rate limit, send 10 emails at a time
            time.sleep(60)
        print("Email Sent")
except:
    print("Error")
#If error please try to debug as there maybe some bugs in the code
