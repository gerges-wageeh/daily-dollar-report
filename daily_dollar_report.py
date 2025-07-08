import requests 
import pandas as pd
from datetime import datetime
import os 
import smtplib
from email.message import EmailMessage
from dotenv import load_dotenv
load_dotenv()

#الحصول علي سعر دولار مقابل الجنيه من API
url = "https://api.exchangerate-api.com/v4/latest/USD"
r = requests.get(url)

try:
    r.raise_for_status()
    data = r.json()
except Exception as e:
    print("An error occurred while fetching the US dollar data.",e)
    exit()

#استخراج سعر دولار مقابل الجنيه
egp_rate = data['rates']['EGP']
today = datetime.now().strftime('%Y-%m-%d')

#تجهيز البيانات في شكل جدول
df =pd.DataFrame([{
    "Date":today,
    "Dollar price in pounds":egp_rate,
}])

# التأكد إذا كان ملف Excel موجود، نضيف عليه
filename = 'dollar_report.xlsx'

if os.path.exists(filename):
    old_df = pd.read_excel(filename)
    new_df = pd.concat([old_df,df],ignore_index=True)
else:
    new_df = df 

#حفظ ملف
new_df.to_excel(filename,index=False)

#ارسال رساله علي الايمال
# إعدادات الإيميل
sender = os.getenv("SENDER")
password = os.getenv("PASSWORD")
receiver = sender

# بناء الرسالة
msg = EmailMessage()
msg['Subject'] = "تقرير الدولار اليومي"
msg['From'] = sender
msg['To'] = receiver
msg.set_content("تم تحديث تقرير الدولار اليوم بنجاح")

# إرسال الرسالة
with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
    try:
       smtp.login(sender, password)
       smtp.send_message(msg)
    except Exception as e:
       print("فشل إرسال الإيميل:", e)
       exit()

print(" تم تنفيذ المشروع بنجاح.") 



