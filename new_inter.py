import pyautogui
import time
import pandas as pd
import webbrowser as web
from urllib.parse import quote

pyautogui.FAILSAFE = False  # Disable the fail-safe mechanism

def send_message(excel_file, message_file, x_cord=467, y_cord=886):
    try:
        df = pd.read_excel(excel_file, dtype={"Mobile no.": str})
        
        # Check if the columns exist
        if 'Name' not in df.columns or 'Mobile no.' not in df.columns:
            raise ValueError("Required columns 'Name' and 'Mobile no.' not found in the Excel file.")
        
        with open(message_file) as f:
            message_template = f.read()

        web.open("https://web.whatsapp.com")
        time.sleep(15)

        message = message_template.format("Recipient Name")  # Replace with the desired recipient name

        for i in range(len(df)):
            contact_number = df.loc[i, 'Mobile no.']

            # Open a new chat with the contact
            web.open(f"https://web.whatsapp.com/send?phone={contact_number}&text={quote(message)}")
            time.sleep(10)

            # Click on the input field to focus
            pyautogui.click(x_cord, y_cord)
            time.sleep(2)

            # Send the message
            pyautogui.write(message, interval=0.1)  # Adjust the interval as needed
            pyautogui.press("enter")
            time.sleep(2)

            # Close the chat
            pyautogui.hotkey("ctrl", "w")
            time.sleep(1)

            print(f"Message sent to {contact_number}")

        print("\nMessage sent to all contacts.")
    except Exception as e:
        print(f"An error occurred: {e}")

excel_path = "C:/Users/omkar/OneDrive/Desktop/new inter/new_inter.xlsx"
message_path = "C:/Users/omkar/OneDrive/Desktop/new inter/message.txt"

send_message(excel_path, message_path)
