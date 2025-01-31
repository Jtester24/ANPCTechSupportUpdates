import logging
from kivy.app import App
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.button import Button
from kivy.uix.dropdown import DropDown
from kivy.uix.label import Label
from kivy.uix.textinput import TextInput
from kivy.uix.checkbox import CheckBox
from kivy.uix.gridlayout import GridLayout
from kivy.graphics import Color, Rectangle
from datetime import datetime
import pandas as pd
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import os
import zipfile
from kivy.core.window import Window
from tkinter import Tk
from tkinter.filedialog import askdirectory
import subprocess
import json
import base64
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build

# Set up logging
logging.basicConfig(filename='app.log', level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')

# If modifying these SCOPES, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/gmail.send']

class ANPCSupportTicket(App):
    def build(self):
        # Set the window size to 80% of the screen width and full screen height
        Window.size = (int(Window.width * 0.8), Window.height)
        
        # Center the window horizontally on the screen
        Window.left = (Window.width - Window.size[0]) // 2
        Window.top = 0  # Align the window to the top of the screen

        root = BoxLayout(orientation='vertical')
        with root.canvas.before:
            Color(0.1, 0.1, 0.5, 1)  # Dark blue background
            self.rect = Rectangle(size=root.size, pos=root.pos)
            root.bind(size=self._update_rect, pos=self._update_rect)

        layout = BoxLayout(orientation='vertical', padding=10, spacing=10, size_hint=(1, 1))

        # Title
        title = Label(text='ANPC Tech Support Ticket', font_size='26sp', size_hint=(1, None), height=50, color=(1, 1, 1, 1))
        layout.add_widget(title)

        # Dropdown for site selection
        self.site_dropdown = DropDown()
        sites = ["SITE 1 TLS2038", "SITE 2 TLS2030", "SITE 3 TLS2028", "SITE 4 TLS2035"]
        self.site_button = Button(text='Select Site', font_size='18sp', size_hint=(1, None), height=50, background_color=(0.2, 0.6, 0.8, 1))
        self.site_button.bind(on_release=self.site_dropdown.open)

        for site in sites:
            btn = Button(text=site, font_size='18sp', size_hint_y=None, height=40, background_color=(0.8, 0.8, 0.8, 1))
            btn.bind(on_release=lambda btn: self.set_site(btn.text))
            self.site_dropdown.add_widget(btn)

        layout.add_widget(self.site_button)

        # Contact info
        contact_layout = GridLayout(cols=2, padding=10, spacing=10, size_hint=(1, None), height=150)
        contact_layout.add_widget(Label(text='Name:', font_size='18sp', color=(1, 1, 1, 1)))
        self.name_input = TextInput(hint_text='Name', font_size='18sp', size_hint_y=None, height=40)
        contact_layout.add_widget(self.name_input)

        contact_layout.add_widget(Label(text='Email:', font_size='18sp', color=(1, 1, 1, 1)))
        self.email_input = TextInput(hint_text='Email', font_size='18sp', size_hint_y=None, height=40)
        contact_layout.add_widget(self.email_input)

        contact_layout.add_widget(Label(text='Whatsapp:', font_size='18sp', color=(1, 1, 1, 1)))
        self.whatsapp_input = TextInput(hint_text='Whatsapp', font_size='18sp', size_hint_y=None, height=40)
        contact_layout.add_widget(self.whatsapp_input)

        layout.add_widget(contact_layout)

        # Alert/Alarm Code and Checkboxes
        info_layout = GridLayout(cols=2, padding=10, spacing=10, size_hint=(1, None), height=300)
        info_layout.add_widget(Label(text='Alert/Alarm Code:', font_size='18sp', color=(1, 1, 1, 1)))
        self.alert_code_input = TextInput(hint_text='Alert/Alarm Code', font_size='18sp', size_hint_y=None, height=40)
        info_layout.add_widget(self.alert_code_input)

        info_layout.add_widget(Label(text='Was Guidance Interrupted?', font_size='18sp', color=(1, 1, 1, 1)))
        self.guidance_interrupted_checkbox = CheckBox(size_hint=(None, None), size=(20, 20))
        info_layout.add_widget(self.guidance_interrupted_checkbox)

        info_layout.add_widget(Label(text='Can TLS Provide CAT 1 Guidance?', font_size='18sp', color=(1, 1, 1, 1)))
        self.cat1_guidance_checkbox = CheckBox(size_hint=(None, None), size=(20, 20))
        info_layout.add_widget(self.cat1_guidance_checkbox)

        info_layout.add_widget(Label(text='Is an RMA Required?', font_size='18sp', color=(1, 1, 1, 1)))
        self.maintenance_required_checkbox = CheckBox(size_hint=(None, None), size=(20, 20))
        info_layout.add_widget(self.maintenance_required_checkbox)

        info_layout.add_widget(Label(text='Is there a Power Issue?', font_size='18sp', color=(1, 1, 1, 1)))
        self.power_issue_checkbox = CheckBox(size_hint=(None, None), size=(20, 20))
        info_layout.add_widget(self.power_issue_checkbox)

        info_layout.add_widget(Label(text='Is there a Network Issue?', font_size='18sp', color=(1, 1, 1, 1)))
        self.network_issue_checkbox = CheckBox(size_hint=(None, None), size=(20, 20))
        info_layout.add_widget(self.network_issue_checkbox)

        info_layout.add_widget(Label(text='Submit an Archive', font_size='18sp', color=(1, 1, 1, 1)))
        self.archive_button = Button(text='Select Folder', font_size='18sp', size_hint=(1, None), height=50, background_color=(0.8, 0.8, 0.8, 1))
        self.archive_button.bind(on_release=self.open_folderchooser)
        info_layout.add_widget(self.archive_button)

        layout.add_widget(info_layout)

        # Comments
        layout.add_widget(Label(text='Comments:', font_size='18sp', color=(1, 1, 1, 1)))
        self.comments_input = TextInput(hint_text='Comments', font_size='18sp', multiline=True, size_hint=(1, 1))
        layout.add_widget(self.comments_input)

        # Network connection button
        self.network_button = Button(text='CONNECT', font_size='18sp', size_hint=(1, None), height=50, background_color=(1, 0, 0, 1))  # Default to red
        self.network_button.bind(on_release=self.connect_to_network)
        layout.add_widget(self.network_button)

        # Submit button
        self.submit_button = Button(text='Submit', font_size='18sp', size_hint=(1, None), height=50, background_color=(0.2, 0.6, 0.8, 1))
        self.submit_button.bind(on_release=self.submit_ticket)
        layout.add_widget(self.submit_button)

        root.add_widget(layout)
        return root

    def _update_rect(self, instance, value):
        self.rect.pos = instance.pos
        self.rect.size = instance.size

    def set_site(self, site_name):
        self.site_button.text = site_name
        self.site_dropdown.dismiss()  # Dismiss the dropdown after selection

    def open_folderchooser(self, instance):
        # Open a folder dialog using tkinter without hiding the Kivy window
        Tk().withdraw()  # Prevents the root window from appearing
        self.archive_folder = askdirectory()
        
        # Update the button text with the selected folder name
        self.archive_button.text = os.path.basename(self.archive_folder) if self.archive_folder else 'Select Folder'

    def connect_to_network(self, instance):
        try:
            # Get the list of preferred networks
            command = 'netsh wlan show profiles'
            result = subprocess.run(command, shell=True, capture_output=True, text=True)
            if result.returncode != 0:
                logging.error(f"Failed to get preferred networks: {result.stderr}")
                self.network_button.background_color = (1, 0, 0, 1)  # Red
                self.network_button.text = "CONNECT - Failed to get preferred networks"
                return

            # Parse the output to get the list of profiles
            profiles = []
            for line in result.stdout.split('\n'):
                if "All User Profile" in line:
                    profile = line.split(":")[1].strip()
                    profiles.append(profile)

            if not profiles:
                logging.error("No preferred networks found")
                self.network_button.background_color = (1, 0, 0, 1)  # Red
                self.network_button.text = "CONNECT - No preferred networks found"
                return

            # Connect to the first preferred network
            ssid = profiles[0]
            command = f'netsh wlan connect name="{ssid}"'
            result = subprocess.run(command, shell=True, capture_output=True, text=True)
            if result.returncode != 0:
                logging.error(f"Failed to connect to {ssid}: {result.stderr}")
                self.network_button.background_color = (1, 0, 0, 1)  # Red
                self.network_button.text = f"CONNECT - Failed to connect to {ssid}"
                return

            logging.info(f"Connected to {ssid} successfully")
            self.network_button.background_color = (0, 1, 0, 1)  # Green
            self.network_button.text = "CONNECT - Connected"

            # Authenticate and get Gmail credentials
            creds = None
            self.token_path = os.path.join(os.path.dirname(__file__), 'token.json')
            if os.path.exists(self.token_path):
                creds = Credentials.from_authorized_user_file(self.token_path, SCOPES)
            if not creds or not creds.valid:
                if creds and creds.expired and creds.refresh_token:
                    creds.refresh(Request())
                else:
                    # Use relative path for client_secret.json
                    client_secret_path = os.path.join(os.path.dirname(__file__), 'client_secret.json')
                    flow = InstalledAppFlow.from_client_secrets_file(client_secret_path, SCOPES)
                    creds = flow.run_local_server(port=0)
                with open(self.token_path, 'w') as token:
                    token.write(creds.to_json())
            self.creds = creds  # Store the credentials for later use
            logging.info("Authenticated with Gmail successfully")
            self.network_button.background_color = (0, 1, 0, 1)  # Green
            self.network_button.text = "CONNECT - Authenticated"
        except Exception as e:
            logging.error(f"Error connecting to network: {e}")
            self.network_button.background_color = (1, 0, 0, 1)  # Red
            self.network_button.text = f"CONNECT - Error: {str(e)}"

    def submit_ticket(self, instance):
        try:
            site = self.site_button.text
            timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            name = self.name_input.text
            email = self.email_input.text
            whatsapp = self.whatsapp_input.text
            alert_code = self.alert_code_input.text
            guidance_interrupted = self.guidance_interrupted_checkbox.active
            cat1_guidance = self.cat1_guidance_checkbox.active
            archive_folder = self.archive_folder if hasattr(self, 'archive_folder') else ''
            comments = self.comments_input.text

            data = {
                'Site': site,
                'Timestamp': timestamp,
                'Name': name,
                'Email': email,
                'Whatsapp': whatsapp,
                'Alert Code': alert_code,
                'Guidance Interrupted': guidance_interrupted,
                'CAT 1 Guidance': cat1_guidance,
                'Archive Folder': archive_folder,
                'Comments': comments
            }

            # Create a DataFrame and save it to an Excel file
            df = pd.DataFrame([data])
            excel_file_path = 'support_ticket.xlsx'
            df.to_excel(excel_file_path, index=False)

            # Zip the contents of the selected folder
            if archive_folder:
                zip_file_path = 'archive.zip'
                with zipfile.ZipFile(zip_file_path, 'w') as zipf:
                    for root, dirs, files in os.walk(archive_folder):
                        for file in files:
                            zipf.write(os.path.join(root, file), os.path.relpath(os.path.join(root, file), archive_folder))

            # Email sending functionality
            sender_email = "anpctechs@gmail.com"  # Replace with your email
            recipients = ["jtester@anpc.com"]  # Add your recipients here
            subject = "New Support Ticket"
            body = f"""Site: {site}
Timestamp: {timestamp}
Name: {name}
Email: {email}
Whatsapp: {whatsapp}
Guidance Interrupted: {guidance_interrupted}
Comments: {comments}"""

            msg = MIMEMultipart()
            msg['From'] = sender_email
            msg['To'] = ", ".join(recipients)
            msg['Subject'] = subject
            msg.attach(MIMEText(body, 'plain'))

            # Attach the Excel file
            with open(excel_file_path, 'rb') as attachment_file:
                attachment = MIMEBase('application', 'octet-stream')
                attachment.set_payload(attachment_file.read())
            encoders.encode_base64(attachment)
            attachment.add_header('Content-Disposition', f'attachment; filename={os.path.basename(excel_file_path)}')
            msg.attach(attachment)

            # Attach the zipped folder
            if archive_folder:
                with open(zip_file_path, 'rb') as zip_attachment_file:
                    zip_attachment = MIMEBase('application', 'octet-stream')
                    zip_attachment.set_payload(zip_attachment_file.read())
                encoders.encode_base64(zip_attachment)
                zip_attachment.add_header('Content-Disposition', f'attachment; filename={os.path.basename(zip_file_path)}')
                msg.attach(zip_attachment)

            # Authenticate and send email using Gmail API
            if not hasattr(self, 'creds') or not self.creds:
                logging.error("No Gmail credentials available. Please connect to the network first.")
                return

            try:
                service = build('gmail', 'v1', credentials=self.creds)
                raw = base64.urlsafe_b64encode(msg.as_bytes()).decode()
                message = {'raw': raw}
                service.users().messages().send(userId='me', body=message).execute()
                logging.info("Email sent successfully")
            except Exception as e:
                if 'invalid_grant' in str(e):
                    logging.error("Token expired or revoked, re-authenticating...")
                    os.remove(self.token_path)
                    self.connect_to_network(instance)
                    self.submit_ticket(instance)
                else:
                    raise e

            # Close the application
            App.get_running_app().stop()
        except Exception as e:
            logging.error(f"Error in submit_ticket: {e}")

if __name__ == '__main__':
    try:
        ANPCSupportTicket().run()
    except Exception as e:
        logging.error(f"Error running app: {e}")