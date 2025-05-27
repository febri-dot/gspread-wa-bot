import tkinter as tk
import gspread 
from oauth2client.service_account import ServiceAccountCredentials
import pywhatkit as kit
from datetime import datetime, timedelta
import time
import threading
import calendar
import pyautogui
import gspread.utils

# Start Row at Google Sheet
start_bills = "A6"
end_bills   = "D"

import sys
import os

# Setup Google Sheet API
if getattr(sys, 'frozen', False):
   base_path = sys._MEIPASS
else:
   base_path = os.path.abspath(".")

cred_path = os.path.join(base_path, "credentials.json")
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name(cred_path, scope)
client = gspread.authorize(creds)

# Open Google Sheet 
sheet = client.open("Rekap Keuangan KTM")
bills = sheet.worksheet("Rekap Tagihan").get(f"{start_bills}:{end_bills}") 
members = sheet.worksheet("Member Info").get_all_records()


def convert_idr_to_int(value) :
   if not value:
      return None
   
   cleaned = value.replace("Rp", "").replace(".", "").replace(" ", "")
   if "," in cleaned:
      cleaned = cleaned.split(",")[0]
   
   return int(cleaned)

def convert_int_to_idr(value) :
   try:
      number = int(value)
      formatted = f"Rp{number:,.0f}".replace(",", ".")
   except (ValueError, TypeError):
      formatted = "-"
   return formatted

def search_member_info(name) :
   for member in members:
      if member["Nama"] == name :
         return member
   
   return None

def get_last_day() :
   last_day = calendar.monthrange(int(datetime.now().strftime("%Y")), int(datetime.now().strftime("%m")))[1]
   last_day = f"{last_day}-{int(datetime.now().strftime("%m"))}-{int(datetime.now().strftime("%Y"))}"

   return last_day

def get_greeting_text() :
   now = int(datetime.now().strftime("%H"))
   if now >= 0 and now <= 4 or now >= 18 :
      greeting = "Selamat malam,\n\n"
   elif now > 4 and now <= 10 :
      greeting = "Selamat pagi,\n\n"
   elif now > 10 and now <= 14 :
      greeting = "Selamat siang,\n\n"
   else :
      greeting = "Selamat sore,\n\n"
   
   return greeting

def send_whatsapp_message(sender) :
   global output 

   greeting = get_greeting_text()
   period   = datetime.now().strftime("%m-%Y")
   opening = f"""Kami dari bendahara KTM ingin mengingatkan terkait pembayaran kas bulanan dan arisan untuk periode {period}. Dengan rincian sebagai berikut:\n"""
   closing = """    💳 Metode pembayaran:
      1.Transfer ke:
         Febriyanti Putri – BCA
         No. Rekening: [Nomor Rekening]
      2.Dibawa langsung saat rapat rutin
      3.Diantar ke rumah bendahara
      4.Dititipkan ke teman yang hadir rapat

Mohon konfirmasi setelah pembayaran ya, agar bisa kami data.
Terima kasih bagi teman-teman yang sudah menyelesaikan kewajibannya 🙏
Jika ada pertanyaan atau kesulitan, jangan ragu hubungi kami.

Terimakasih

— Bendahara KTM"""

   how_many = 0

   # Get Dynamic Column 
   headers = sheet.worksheet("Member Info").row_values(1)
   if period in headers:
      column = headers.index(period) + 1
   else:
      error = "Periode belum tersedia di header Google Sheet."
      output.after(0, update_output, error)
      raise Exception(error)

   # Send WhatsApp Message
   for bill in bills :
      status = ""
      name = bill[0]
      bill.pop(0)
      if name :
         bill_list = [convert_idr_to_int(b) for b in bill]
         
         member_info = search_member_info(name)
         if bill_list[2] != 0 :
            if member_info is not None and sender.lower() == member_info["PJ Chat"].lower() and member_info[period] not in ["Done", "No Dues"]:
               billing = f"""*Tagihan atas nama {"sdr." if member_info["Jenis Kelamin"] == "L" else "sdri."} {member_info["Nama Panggilan"]} :*
      🔹 Kas bulanan: *{convert_int_to_idr(bill_list[0])}*
      🔹 Arisan: *{convert_int_to_idr(bill_list[1])}*
      💰 Total: *{convert_int_to_idr(bill_list[2])}*
      📅 Batas pembayaran: *{get_last_day()}*\n\n"""
               message = greeting + opening + billing + closing

               # Schedule message 2 minutes from now
               # now = datetime.now() + timedelta(minutes=2)
               # hour = now.hour
               # minute = now.minute

               try:
                  # kit.sendwhatmsg(f"+{member_info["Nomor WA"]}", message, hour, minute)
                  kit.sendwhatmsg_instantly(f"+{member_info['Nomor WA']}", message, wait_time=20, tab_close=False)
                  
                  time.sleep(15) 
                  pyautogui.press("enter")
                  print("Message scheduled successfully!")
                  output.after(0, update_output, f"Pesan untuk {member_info['Nama Panggilan']} sudah terkirim\n")

                  how_many += 1
                  status = "Done"
                  # time.sleep(50)
               except Exception as e:
                  print(f"An error occurred: {e}")
                  status = "Failed"
         else :
            status = "No Dues"
         
         if status and sender.lower() == member_info["PJ Chat"].lower():
            row = int(member_info["No"]) + 1
            
            cell = gspread.utils.rowcol_to_a1(row, column)
            print(cell, status)
            sheet.worksheet("Member Info").update_acell(cell, status)
   
   
   success_output = f"Pesan terkirim oleh {sender} : {how_many}\n" if how_many > 1 else f"Tidak ada pesan yang perlu dikirim oleh {sender}.\n"
   output.after(0, update_output, success_output)

def update_output(message) :
   output.insert(tk.END, message)
   output.see(tk.END)


def who_are_you() :
   global ROOT, sender

   name = sender.get().strip()
   if not name:
      output.insert(tk.END, "Silakan masukkan nama pengirim terlebih dahulu.\n")
      output.see(tk.END)
      return

   output.insert(tk.END, f"Memulai pengiriman sebagai {name}...\n")
   output.see(tk.END)

   # Start new thread 
   thread = threading.Thread(target=send_whatsapp_message, args=(name,))
   thread.start()



def main() :
   global ROOT, sender, output

   ROOT = tk.Tk()
   ROOT.title("KTM 39 WA Billing Sender Bot")

   tk.Label(ROOT, text="Sender :").pack()
   sender = tk.Entry(ROOT)
   sender.pack()

   tk.Button(ROOT, text="Kirim Pesan", command=who_are_you).pack()

   output = tk.Text(ROOT, height=15)
   output.pack()

   ROOT.mainloop()

if __name__ == "__main__" :
   main()