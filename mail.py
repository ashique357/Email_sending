import smtplib
import xlrd  
import sys
import pandas as pd

file_path = ("mail_list.xlsx")
email_list=[]

login_credentials={
    "email":"ashiquejubayer357@gmail.com",
    "password":"yvziusbsbsvgdpmn"
}

errors={
    "server_error":"Server Connection failed",
    "authentication_error":"Authentication failed",
    "email_not_found_error":"Email not found",
}
sucess="Sucessfully Sent the email"

body="This is the body of the email"

try:
    df=pd.read_excel(file_path, index_col=0)
    for rows in df.iterrows():
        email =rows[0].lower()
        email_list.append(email)

    try:
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        
        try:
            server.login(login_credentials["email"],login_credentials["password"])
            
            for to_email in email_list:
                try:
                    server.sendmail(login_credentials["email"], to_email, body)
                    print(sucess)
                
                except:
                    print (errors["email_not_found_error"])    

            server.close()

        except:
            print (errors["authentication_error"])            
    except:
        print (errors["server_error"])
    
except Exception as e:
    print("Oops!", e.__class__, "occurred.")
