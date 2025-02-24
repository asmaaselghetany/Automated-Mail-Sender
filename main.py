import pandas as pd
import datetime
import smtplib
import time
import random
from win10toast import ToastNotifier
  
#for desktop notification
toast = ToastNotifier()

#your gmail credentials here
My_gmail_id = 'Your_Gmail_Id'
My_gmail_pwd = 'Your_Gmail_psswd'
  
# define a function for sending email
def send_Email(to, sub, msg):
    
    # conncection to gmail
    connection = smtplib.SMTP('smtp.gmail.com', 587) 
      
    # starting the session
    connection.starttls()     
      
    # login using credentials
    connection.login(My_gmail_id, My_gmail_pwd)   
      
    # sending email
    connection.sendmail(My_gmail_id, to, 
                   f"Subject : {sub}\n\n{msg}") 
      
    # quit the session
    connection.quit()  
      
    print("Email sent to " + str(to) + " with subject " 
          + str(sub) + " and message :" + str(msg))
      
    toast.show_toast("Email Sent!" , 
                     f"{name} was sent e-mail",
                     threaded = True, 
                     icon_path = None,
                     duration = 6)
  
    while toast.notification_active():
        time.sleep(0.1)
  
if __name__=="__main__":
    
      # read the excel sheet having all the details
    dataframe = pd.read_excel('birthday.xlsx')   
      
    # today date in format : DD-MM
    today = datetime.datetime.now().strftime("%d-%m") 
 
    # writeindex list
    writeInd = []                                                   
  
    for index,item in dataframe.iterrows():
        
        num = random.randint(1,3)
        with open(f"letters/letter_{num}.txt") as letter:
            lines = letter.readlines()
            message = "".join(lines) 
                 
        # stripping the birthday in excel 
        name = item['Name']
        # sheet as : DD-MM
        b_day = item['Birthday'].strftime("%d-%m")        
          
        # condition checking
        if (today == b_day):    
              
            # calling the sendEmail function
            send_Email(item['Email'], "Happy Birthday",lines)    
      