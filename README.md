# Automated-WhatsApp-Message-Sender-Using-Python

This Python script demonstrates a robust automation process for sending personalized WhatsApp messages to multiple recipients using the PyAutoGUI library. It integrates with Pandas for data manipulation and the webbrowser module for web interaction. The script enables users to streamline communication efforts by automating the message sending process, which is especially beneficial for large-scale messaging campaigns.

The code begins by importing necessary libraries such as pyautogui, time, pandas, webbrowser, and urllib.parse. It configures the PyAutoGUI library and defines a function named send_message, which serves as the main mechanism for sending messages to WhatsApp contacts. The function first reads data from an Excel file provided by the user. It checks whether the required columns, 'Name' and 'Mobile no.', exist in the Excel file. If these columns are not found, it raises a ValueError to ensure data integrity.

To enable dynamic messaging, the script reads the contents of a message template file, ensuring that the placeholder for the recipient's name is appropriately replaced. The code then initiates the WhatsApp web interface using the webbrowser module and waits for the page to load before proceeding. It retrieves the message template and customizes it with the recipient's name.

For each contact in the Excel file, the script opens a new chat window with the corresponding WhatsApp number and pastes the customized message into the chat input field. It utilizes PyAutoGUI's automation capabilities to simulate mouse clicks and keyboard inputs for sending the message. Subsequently, the script closes the chat window and moves on to the next contact. Throughout the process, the script incorporates time delays to ensure proper synchronization between actions.

In case of any exceptions during the execution, the script catches and prints the corresponding error message to facilitate debugging and error handling.

This script demonstrates how to integrate Python with automation libraries and data manipulation tools, enabling seamless communication automation. By referring to the PyAutoGUI documentation, users can explore its capabilities in automating various GUI interactions. Additionally, the Pandas documentation provides insights into efficient data manipulation, while the webbrowser module documentation illustrates its functionalities in web interaction. Understanding these resources can further enhance the script's capabilities and facilitate the development of more sophisticated automation solutions.

Reference Links:

1. PyAutoGUI Documentation: https://pyautogui.readthedocs.io/en/latest/
2. Pandas Documentation: https://pandas.pydata.org/docs/
3. urllib.parse Documentation: https://docs.python.org/3/library/urllib.parse.html
4. WebBrowser Module Documentation: https://docs.python.org/3/library/webbrowser.html
5. Openpyxl Documentation (for Excel manipulation): https://openpyxl.readthedocs.io/en/stable/
