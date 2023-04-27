import html
import pandas as pd
import os
import smtplib
import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
from bs4 import BeautifulSoup
import matplotlib.pyplot as plt
import io
import base64
import matplotlib as mpl




# Load the data from the original Excel file
data = pd.read_excel('agents.xlsx')
data.rename(columns={'מייל של סוכן':'כתובת מייל סוכן'},inplace=True)
data.rename(columns={'שדה כללי 1':'סטטוס'},inplace=True)
data.rename(columns={'סוכן בכרטיס לקוח':'שם סוכן'},inplace=True)

#set date left to right
data['תאריך יצירה'] = data['תאריך יצירה'].dt.strftime('%d-%m-%Y')
data['תאריך שינוי'] = data['תאריך שינוי'].dt.strftime('%d-%m-%Y')


no_email_agents = []
agents = data['שם סוכן'].unique()

# Loop foreach agent and create a new Excel file for them
for agent in agents:
    wb = openpyxl.Workbook()
    ws = wb.active
    agent_data = data[data['שם סוכן'] == agent]

    if agent_data['כתובת מייל סוכן'].isnull().values.all():
        no_email_agents.append(agent)
        continue

    for r in dataframe_to_rows(agent_data.drop(columns=['כתובת מייל סוכן']), index=False, header=True):
        ws.append(r)

    # Format the file as a table
    tab = openpyxl.worksheet.table.Table(displayName="Table", ref=ws.dimensions)
    style = openpyxl.worksheet.table.TableStyleInfo(name="TableStyleMedium9", showFirstColumn=False, showLastColumn=False, showRowStripes=True, showColumnStripes=False)
    tab.tableStyleInfo = style
    ws.add_table(tab)

    wb.save(f'{agent}.xlsx')
    workbook = openpyxl.load_workbook(f'{agent}.xlsx')
    worksheet = workbook.active
    worksheet.sheet_view.rightToLeft = True


    worksheet.sheet_format.defaultColWidth = 12
    worksheet.sheet_format.defaultRowHeight = 15
    
    #Delete and save again to avoid issues   
    os.remove(f'{agent}.xlsx')
    workbook.save(f'{agent}.xlsx')
    
    #The design of the table that will be sent in the body of the email
    df = pd.read_excel(f'{agent}.xlsx')
    df.index += 1
    styled_table = (df.style
    .set_properties(**{'color': '#000000'})
    .set_table_styles([
        {'selector': 'th', 'props': [('background-color', '#E6DDC4')]},
        {'selector': 'table', 'props': [('border-collapse', 'collapse')]},
        {'selector': 'th, td', 'props': [('border', '2px solid #678983')]}
    ])
    .set_table_attributes('style="width: 100%; max-width: 700%;"'))
    styled_table.set_properties(subset=['תאריך יצירה'], **{'width': '10%'})
    styled_table.set_properties(subset=['תאריך שינוי'], **{'width': '10%'})
    html_table = styled_table.to_html().replace('nan', '')
    

    #Pie chart design
    plt.clf()
    reversed_statuses = agent_data["סטטוס"].apply(lambda x: x[::-1])#Reverses the statuses from right to left
    colors=['#FD8A8A','#F1F7B5','#A8D1D1','#9EA1D4','#FFD4B2','#ECA869']
    graph = reversed_statuses.value_counts().plot(kind='pie',autopct='%1.1f%%',colors=colors) 
    graph.set_ylabel('')
    
    # Save the pie chart as an image file
    img_buffer = io.BytesIO()
    plt.savefig(img_buffer, format='png')
    img_buffer.seek(0)
    img_base64 = base64.b64encode(img_buffer.read()).decode('utf-8')


    # Create a HTML email body and attaches the table and the graph
    body = f"""
<html>
    <head>
        <meta charset="utf-8">
        <title>{agent} - סטטוס לקוחות</title>
        <style>
            table {{
                border-collapse: collapse;
                width: 100%;
                max-width: 700%;
            }}
            th {{
                background-color: #E6DDC4;
                border: 2px solid #678983;
                padding: 6px;
                text-align: center;
            }}
            td {{
                border: 2px solid #678983;
                padding: 6px;
                text-align: center;
                color:black;
            }}
            .header {{
                color: #000000;
                font-weight: bold;
                margin-bottom: 10px;
            }}
            
        </style>
    </head>
    <body dir="rtl">
        <p class="header">היי {agent},</p>
        <p>מצ"ב סטטוס לקוחות שלך <b>מתחילת שנה</b>.</p>
        <p>מצורף קובץ אקסל לנוחיותך.</p>
        <table>{html_table}</table>
        <center>
            <img src="data:image/png;base64,{img_base64}" alt="" title="logo" style="display:block;">
        </center>
        <hr>
        <p>בברכה,</p>     
        <p>name</p>
    </body>
    </html>"""



    # Defines the email

    fromaddr = 'email-addres'#Change the email as you wish
    toaddr = agent_data['כתובת מייל סוכן'].iloc[0]  # foreach agent
    msg = MIMEMultipart()
    msg['From'] = fromaddr
    msg['To'] = toaddr
    msg['Subject'] = f'{agent} - סטטוס לקוחות'
    msg.attach(MIMEText(body, 'html'))

    # Attach the Excel file to the email
    attachment = open(f'{agent}.xlsx', 'rb')
    part = MIMEBase('application', 'octet-stream')
    part.set_payload((attachment).read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', 'attachment', filename=f'{agent}-סטטוס לקוחות.xlsx')
    msg.attach(part)

    # Send the email
    #server = smtplib.SMTP('smtp.gmail.com', 587) #Gmail server...optional port - 465
    
    server = smtplib.SMTP('smtp.office365.com', 587)
    
    server.starttls()
    server.login(fromaddr, 'passowrd')
    text = msg.as_string()
    server.sendmail(fromaddr, toaddr, text)
    server.quit()

    # Delete the Excel file...
    #os.remove(f'{agent}.xlsx')
    
    
    #----------------------------end for--------------------------
    
    
    
    #Prints the agents that did not have an email address in the original file
if no_email_agents:
    print("The following agents still do not have an email address:")
    print(no_email_agents)