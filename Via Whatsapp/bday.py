import pandas as pd
import datetime
import os
import pywhatkit as kit

os.chdir(r"C:\Users\Aditya Agrawal\Desktop\Python\My codes\Birthday Wisher\Via Whatsapp")


if __name__=="__main__":
    df=pd.read_excel("data.xlsx")
    #print(df)
    today=datetime.datetime.now().strftime("%d-%m")
    yearNow=datetime.datetime.now().strftime("%Y")
    #print(today)
    writeInd = []
    for index, item in df.iterrows():
        #print(index,item['Birthday'])
        bday=item['Birthday'].strftime("%d-%m")
        #print(bday)
        #print(format((int(bday.split("-")[0])-1),'02d'))
        curr_day=str(format((int(bday.split("-")[0])-1),'02d'))+"-"+bday.split("-")[1]
        #print(curr_day)
        
        if(today==curr_day) and yearNow not in str(item['Year']):
            print("sending msg to "+item['Name'])
            string_number="+91"+str(item['Number'])
            kit.sendwhatmsg(string_number,item['Dialogue'],00,00)
            writeInd.append(index)

    #print(writeInd)
    for i in writeInd:
        yr=df.loc[i,'Year']
        df.loc[i,'Year']= str(yr) + ", " + str(yearNow)
        #print(df.loc[i,'Year'])

    #print(df)
    df.to_excel('data.xlsx',index=False)
