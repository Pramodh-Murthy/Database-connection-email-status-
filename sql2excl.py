import pandas as pd
import pyodbc
import smtplib
from email.message import EmailMessage
import logging
#----------------------------------------------------------------------------------------------------------------------------------------
#successfull mail if there is no error in program
def successmsg():
    msg = EmailMessage()
    msg['Subject'] = 'Training'
    msg['From'] = 'Pramodh team'
    msg['To'] = 'Email@gmail.com'

    with open('success.txt') as myfile:
        data = myfile.read()
        msg.set_content(data)

    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
        server.login("Email@gmail.com", "pavkgxojakwbpgga")
        server.send_message(msg)
        print("emailsent !!!")
#------------------------------------------------------------------------------------------------------------------------------------
#failur mail if there is any error in program with log file as attachment
def failmsg():
    msg = EmailMessage()
    msg['Subject'] = 'Training'
    msg['From'] = 'Pramodh team'
    msg['To'] = 'EmailTo@gmail.com'

    with open('Failed.txt') as myfile:
        data = myfile.read()
        msg.set_content(data)

    with open("example.log", "rb") as f:
        file_data = f.read()
        file_name = f.name
        msg.add_attachment(file_data, maintype="application", subtype="csv", filename=file_name)

    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
        server.login("email@gmail.com", "pavkgxojakwbpgga")
        server.send_message(msg)
    print("emailsent !!!")

#---------------------------------------------------------------------------------------------------------------------------------------
logging.basicConfig(filename='example.log', level=logging.DEBUG,format='%(asctime)s -%(name)s-%(levelname)s %(message)s', datefmt='%d-%b-%y %H:%M:%S')
#reading the csv file 
df = pd.read_csv(r'C:\Users\Dell\Desktop\Task\database.csv')

#creating database connection ,table and inserting data 
try:
   connection = pyodbc.connect('Driver={SQL Server};''Server=DESKTOP-SC8V9DU\SQLEXPRESS;''Database=training;''Trusted_Connection=yes;')
   logging.info('DatBase Connected')
   cursor = connection.cursor()
   cursor.execute('''     CREATE TABLE DATA (ID int primary key,NAME varchar(50),AGE int,GENDER VARCHAR(10))     ''')
   logging.info('table created')
   

   for row in df.itertuples():
       cursor.execute('''   INSERT INTO DATA (ID, NAME, AGE,GENDER)VALUES (?,?,?,?)    ''', row.ID,row.NAME,row.AGE,row.GENDER)
       connection.commit()
   logging.info('data inserted')

#writing data into excel file
   with pd.ExcelWriter("Output.xlsx", engine='xlsxwriter',engine_kwargs={'options': {'strings_to_numbers': True}}) as writer:
            df = pd.read_sql("Select top 5 * from data", connection)
            df.to_excel(writer, sheet_name="Sheet1", header=True, index=False)
            logging.info("File saved successfully!")
            print(df)
            successmsg()
           
except Exception as e:
    cursor.close()
    connection.close()
    logging.error("Exception occurred", exc_info=True) 
    failmsg()