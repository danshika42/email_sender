import pandas as pd
import datetime
import smtplib
import schedule
import time
import json



def sendemail(to,sub,msg):
    s=smtplib.SMTP('smtp.gmail.com',587)
    s.starttls()
    s.login(data["GMAIL_ID"],data["GMAIL_PSWD"])
    s.sendmail(data["GMAIL_ID"],to,f"Subject:{sub}\n\n{msg}")
    s.quit()



def birthday_automator():
    df=pd.read_excel("data.xlsx")
    today=datetime.datetime.now().strftime("%d-%m")
    yearnow=datetime.datetime.now().strftime("%Y")
    writeInd=[]
    for index,item in df.iterrows():
        bday=item["Birthday"].strftime("%d-%m")
        if (today==bday) and (yearnow not in str(item["Year"])):
            sendemail(item["Email"],"Happy Subject",item["Dialouge"])
            writeInd.append(index)

    for i in writeInd:
        df.loc[i,"Year"]=str(yearnow)+','+str(df.loc[i,"Year"])

    df.to_excel("data.xlsx")



    

if __name__=="__main__":
    
    with open('data.json', 'r') as f:
        data = json.load(f)
    schedule.every().day.at("17:14").do(birthday_automator)

    while True:
        schedule.run_pending()
        time.sleep(30) 
  
   