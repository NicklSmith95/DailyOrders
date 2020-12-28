import smtplib, ssl
import pandas as pd
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pretty_html_table import build_table
import datetime

df = pd.read_excel('Y:\January 2021 Hardware-Software Orders.xlsx', sheet_name='Hardware Orders', skiprows=1, dtype=object)
formatted_df = df
formatted_df['Date'] = df['Date'].dt.strftime('%Y-%m-%d')
df['Date'] = formatted_df['Date']
cdate = datetime.date.today()
fdate = cdate.strftime('%Y-%m-%d')
new = df[['Date', '    Attn to:', 'Cost Center', 'Hardware ordered', 'Ticket #', 'Shipping Method', 'Address']]
current = new['Date'] == fdate
currentorders = new.loc[current]
now = datetime.datetime.now()
ddate = now.strftime("%m/%d")


def send_email():
    user = 'nsmith@tbccorp.com'
    password = #######
    recipients = ['itsupport@directtechnologygroup.com','itpurchaseorders@tbccorp.com', 'ITAM@TBCCORP.com']
    message = MIMEMultipart("alternative")
    message["Subject"] = f"Orders {ddate}"
    message["From"] = user
    message["To"] = ", ".join(recipients)
    html_table = build_table(currentorders, 'blue_light', font_size='small')

    # Create the plain-text and HTML version of your message
    html = f"""\
    <html>
    <body>
        <p>Good Afternoon,<br>
        <br>
        Below are the orders for today. Please reply to all with a PO when ready<br>
        {html_table}<br>
        <br>
        Thanks,
        </p>
    </body>
    </html>
    """

    # Turn these into plain/html MIMEText objects
    part2 = MIMEText(html, "html")

    # Add HTML/plain-text parts to MIMEMultipart message
    message.attach(part2)

    try:
        server = smtplib.SMTP('mailrelay.tbccorp.com', 587)
        server.connect('mailrelay.tbccorp.com', 587)
        server.starttls()
        server.ehlo()
        server.login(user, password)
        server.sendmail(user, recipients, message.as_string())
        server.close()

        print('Email sent!')
    except Exception as e:
        print(e)

send_email()
