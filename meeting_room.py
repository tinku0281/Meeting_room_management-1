import streamlit as st
import datetime
from datetime import timedelta
import random 
import pandas as pd
import re
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pytz import timezone 
import pytz
import gspread
from oauth2client.service_account import ServiceAccountCredentials
def set_app_style():
    st.markdown(
        """
        <style>
        /* Base dark theme */
        .stApp {
            background-color: #121212;
            background-image: linear-gradient(to bottom, #0a0a0a 0%, #1e1e1e 100%);
            color: #e0e0e0 !important;
        }
        
        /* Typography - White/Blue hierarchy */
        h1 {
            color: #64b5f6 !important;
            font-family: 'Segoe UI', sans-serif;
            font-weight: 700;
            border-bottom: 2px solid #333;
            padding-bottom: 8px;
            margin-bottom: 1.5rem;
            text-shadow: 0 1px 2px rgba(0,0,0,0.3);
        }
        h2 {
            color: #90caf9 !important;
            font-family: 'Segoe UI', sans-serif;
            font-weight: 600;
            margin        
        /* Content containers */
        .main-container {
            background-color: rgba(255, 255, 255, 0.95);
            border-radius: 12px;
            box-shadow: 0 4px 20px rgba(0, 0, 0, 0.05);
            padding: 2.5rem;
            margin-bottom: 100px;
            border: 1px solid #f0f0f0;
        }
        
        /* Sidebar - light gray with dark text */
        [data-testid="stSidebar"] {
            background-color: #f8f9fa !important;
            border-right: 1px solid #eaeaea;
        }
        [data-testid="stSidebar"] * {
            color: #333333 !important;
        }
        
        /* Interactive elements */
        .stButton>button {
            background-color: #1a3e8c !important;
            color: white !important;
            border-radius: 8px !important;
            border: none;
            padding: 8px 16px;
            transition: all 0.3s;
        }
        .stButton>button:hover {
            background-color: #2a56c6 !important;
            transform: translateY(-1px);
            box-shadow: 0 2px 8px rgba(0,0,0,0.1);
        }
        
        /* Input fields */
        .stTextInput>div>div>input,
        .stSelectbox>div>select,
        .stDateInput>div>input,
        .stTimeInput>div>input,
        .stNumberInput>div>input {
            border: 1px solid #e0e0e0 !important;
            border-radius: 8px !important;
            padding: 10px !important;
        }
        
        /* Footer */
        .footer {
            position: fixed;
            bottom: 0;
            width: 100%;
            background: rgba(255, 255, 255, 0.9);
            padding: 12px 0;
            border-top: 1px solid #eaeaea;
            z-index: 100;
            backdrop-filter: blur(5px);
        }
        .footer p {
            color: #666 !important;
            font-size: 0.75rem;
            margin: 0;
            font-family: 'Segoe UI', sans-serif;
        }
        .footer a {
            color: #2a56c6 !important;
            text-decoration: none;
        }
        
        /* Tables/DataFrames */
        .stDataFrame, .stTable {
            border-radius: 8px !important;
            box-shadow: 0 2px 10px rgba(0,0,0,0.05) !important;
        }
        </style>
        """,
        unsafe_allow_html=True
    )

set_app_style()

# --- Configuration ---
st.set_page_config(
    page_title="Meeting Room Booking",
    page_icon=":calendar:",
    initial_sidebar_state="expanded",
)

# --- Constants ---
IST = pytz.timezone('Asia/Kolkata')
CURRENT_TIME_IST = datetime.datetime.now(IST)
CTIF = CURRENT_TIME_IST.strftime("%y-%m-%d %H:%M:%S")

# Room capacities
ROOM_CAPACITY = {
    "HIMALAYA - Basement": 20,
    "NEELGIRI - Ground Floor": 7,
    "ARAVALI  - Ground Floor": 7,
    "KAILASH - 1 Floor": 7,
    "ANNAPURNA - 1 Floor": 4,
    "EVEREST  - 2 Floor": 12,
    "KANANACJUNGA - 2 Floor": 7,
    "SHIVALIK - 3 Floor": 4,
    "TRISHUL - 3 Floor": 4,
    "DHAULAGIRI - 3 Floor": 7,
}

# --- Google Sheets Setup ---
def init_google_sheets():
    scope = ["https://spreadsheets.google.com/feeds", 
             "https://www.googleapis.com/auth/drive"]
    
    creds_dict = {
        "type": st.secrets["gsheets"]["type"],
        "project_id": st.secrets["gsheets"]["project_id"],
        "private_key_id": st.secrets["gsheets"]["private_key_id"],
        "private_key": st.secrets["gsheets"]["private_key"],
        "client_email": st.secrets["gsheets"]["client_email"],
        "client_id": st.secrets["gsheets"]["client_id"],
        "auth_uri": st.secrets["gsheets"]["auth_uri"],
        "token_uri": st.secrets["gsheets"]["token_uri"],
        "auth_provider_x509_cert_url": st.secrets["gsheets"]["auth_provider_x509_cert_url"],
        "client_x509_cert_url": st.secrets["gsheets"]["client_x509_cert_url"]
    }
    
    creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
    client = gspread.authorize(creds)
    
    try:
        spreadsheet = client.open("Meeting_Room_Bookings")
    except gspread.SpreadsheetNotFound:
        spreadsheet = client.create("Meeting_Room_Bookings")
        spreadsheet.share(st.secrets["gsheets"]["client_email"], perm_type='user', role='writer')
    
    try:
        worksheet = spreadsheet.worksheet("Bookings")
    except gspread.WorksheetNotFound:
        worksheet = spreadsheet.add_worksheet(title="Bookings", rows=1000, cols=20)
        headers = [
            "booking_id", "date", "start_time", "end_time", "room", 
            "name", "email", "description", "cc_emails", "created_at"
        ]
        worksheet.append_row(headers)
    
    return worksheet

# Initialize Google Sheets
worksheet = init_google_sheets()

# --- Data Management Functions ---
def get_all_bookings():
    records = worksheet.get_all_records()
    booking_data = {"room_bookings": {}, "room_availability": {}}
    
    for record in records:
        booking_id = int(record["booking_id"])
        booking_data["room_bookings"][booking_id] = {
            "booking_id": booking_id,
            "date": record["date"],
            "start_time": record["start_time"],
            "end_time": record["end_time"],
            "room": record["room"],
            "name": record["name"],
            "email": record["email"],
            "description": record["description"],
            "cc_emails": record.get("cc_emails", ""),
        }

        if record["date"] not in booking_data["room_availability"]:
            booking_data["room_availability"][record["date"]] = {}
        if record["room"] not in booking_data["room_availability"][record["date"]]:
            booking_data["room_availability"][record["date"]][record["room"]] = []
        booking_data["room_availability"][record["date"]][record["room"]].append(
            (record["start_time"], record["end_time"])
        )
    
    return booking_data

def add_booking_to_sheet(booking_data):
    row = [
        booking_data["booking_id"],
        booking_data["date"],
        booking_data["start_time"],
        booking_data["end_time"],
        booking_data["room"],
        booking_data["name"],
        booking_data["email"],
        booking_data["description"],
        booking_data.get("cc_emails", ""),
        CTIF,
    ]
    worksheet.append_row(row)

def remove_booking_from_sheet(booking_id):
    cell = worksheet.find(str(booking_id))
    if cell:
        worksheet.delete_row(cell.row)

# Load existing booking data
booking_data = get_all_bookings()

# --- Utility Functions ---
def is_valid_time(time_str):
    try:
        datetime.datetime.strptime(time_str, '%H:%M')
        return True
    except ValueError:
        return False

def is_room_available(date, start_time, end_time, room):
    if date not in booking_data["room_availability"]:
        return True

    if room not in booking_data["room_availability"][date]:
        return True

    for booking in booking_data["room_availability"][date][room]:
        b_start_time, b_end_time = booking
        if not (end_time <= b_start_time or start_time >= b_end_time):
            return False

    return True

def generate_random_booking_id():
    return random.randint(1000, 9999)

def is_upcoming(booking, current_datetime):
    date_str = booking["date"]
    time_str = booking["start_time"]
    booking_date = datetime.datetime.strptime(date_str, '%Y-%m-%d').date()
    booking_time = datetime.datetime.strptime(time_str, '%H:%M:%S').time()
    booking_datetime = datetime.datetime.combine(booking_date, booking_time)
    current_datetime = datetime.datetime.strptime(current_datetime, '%y-%m-%d %H:%M:%S')
    return booking_datetime > current_datetime

# --- Email Functions ---
def send_email(to_email, cc_emails, subject, html_content):
    sender_email = st.secrets['email']['sender_email']
    sender_password = st.secrets['email']['sender_password']
    
    msg = MIMEMultipart()
    msg['From'] = 'Meeting Room Booking System'
    msg['To'] = to_email
    msg['Subject'] = subject
    
    if cc_emails:
        msg['Cc'] = ", ".join(cc_emails) if isinstance(cc_emails, list) else cc_emails
    
    msg.attach(MIMEText(html_content, 'html'))
    
    try:
        with smtplib.SMTP("smtp.gmail.com", 587) as server:
            server.starttls()
            server.login(sender_email, sender_password)
            
            recipients = [to_email]
            if cc_emails:
                if isinstance(cc_emails, str):
                    recipients.extend([email.strip() for email in cc_emails.split(',')])
                elif isinstance(cc_emails, list):
                    recipients.extend(cc_emails)
            
            server.sendmail(sender_email, recipients, msg.as_string())
        return True
    except Exception as e:
        st.error(f"Error sending email: {str(e)}")
        return False

def send_confirmation_email(booking_info):
    html_content = f"""
    <html>
    <body>
        <p>Hello {booking_info['name']}!</p>
        <p>We're thrilled to confirm your booking. Here are the details of your reservation:</p>
        <table style="width: 100%; border-collapse: collapse;">
            <tr style="border-bottom: 1px solid #ddd;">
                <td style="padding: 8px;"><strong>Booking ID:</strong></td>
                <td style="padding: 8px;">{booking_info['booking_id']}</td>
            </tr>
            <tr style="border-bottom: 1px solid #ddd;">
                <td style="padding: 8px;"><strong>Meeting Title:</strong></td>
                <td style="padding: 8px;">{booking_info['description']}</td>
            </tr>
            <tr style="border-bottom: 1px solid #ddd;">
                <td style="padding: 8px;"><strong>Date:</strong></td>
                <td style="padding: 8px;">{booking_info['date']}</td>
            </tr>
            <tr style="border-bottom: 1px solid #ddd;">
                <td style="padding: 8px;"><strong>Location:</strong></td>
                <td style="padding: 8px;">{booking_info['room']}</td>
            </tr>
            <tr style="border-bottom: 1px solid #ddd;">
                <td style="padding: 8px;"><strong>Start Time:</strong></td>
                <td style="padding: 8px;">{booking_info['start_time']}</td>
            </tr>
            <tr style="border-bottom: 1px solid #ddd;">
                <td style="padding: 8px;"><strong>End Time:</strong></td>
                <td style="padding: 8px;">{booking_info['end_time']}</td>
            </tr>
        </table>
        <p>Get ready for a productive meeting!</p>
        <p>Best regards,<br>Meeting Room Booking Team <br> SUGAM GROUP</p>
    </body>
    </html>
    """
    
    cc_emails = booking_info.get('cc_emails', '')
    if cc_emails and isinstance(cc_emails, str):
        cc_emails = [email.strip() for email in cc_emails.split(',') if email.strip()]
    
    subject = f"✅ Meeting Room Booking Confirmation: (ID-{booking_info['booking_id']})"
    return send_email(booking_info['email'], cc_emails, subject, html_content)

def send_cancellation_email(booking_info):
    html_content = f"""
    <html>
    <body>
        <p>Hello {booking_info['name']}!</p>
        <p>Your booking has been canceled. Here are the details:</p>
        <table style="width: 100%; border-collapse: collapse;">
            <tr style="border-bottom: 1px solid #ddd;">
                <td style="padding: 8px;"><strong>Booking ID:</strong></td>
                <td style="padding: 8px;">{booking_info['booking_id']}</td>
            </tr>
            <tr style="border-bottom: 1px solid #ddd;">
                <td style="padding: 8px;"><strong>Meeting Title:</strong></td>
                <td style="padding: 8px;">{booking_info['description']}</td>
            </tr>
            <tr style="border-bottom: 1px solid #ddd;">
                <td style="padding: 8px;"><strong>Date:</strong></td>
                <td style="padding: 8px;">{booking_info['date']}</td>
            </tr>
            <tr style="border-bottom: 1px solid #ddd;">
                <td style="padding: 8px;"><strong>Location:</strong></td>
                <td style="padding: 8px;">{booking_info['room']}</td>
            </tr>
            <tr style="border-bottom: 1px solid #ddd;">
                <td style="padding: 8px;"><strong>Start Time:</strong></td>
                <td style="padding: 8px;">{booking_info['start_time']}</td>
            </tr>
            <tr style="border-bottom: 1px solid #ddd;">
                <td style="padding: 8px;"><strong>End Time:</strong></td>
                <td style="padding: 8px;">{booking_info['end_time']}</td>
            </tr>
        </table>
        <p>Contact us if you have any questions.</p>
        <p>Best regards,<br>Meeting Room Booking Team</p>
    </body>
    </html>
    """
    
    cc_emails = booking_info.get('cc_emails', '')
    if cc_emails and isinstance(cc_emails, str):
        cc_emails = [email.strip() for email in cc_emails.split(',') if email.strip()]
    
    subject = f"🚫 Cancellation Confirmation: (ID-{booking_info['booking_id']})"
    return send_email(booking_info['email'], cc_emails, subject, html_content)

# --- Booking Functions ---
def book_room():
    st.header("Choose Meeting Room")
    date = st.date_input("Select Date:", min_value=CURRENT_TIME_IST.date(), value=None)
    current_date = CURRENT_TIME_IST.date()
    
    if date:
        office_start_time = datetime.time(8, 0)
        office_end_time = datetime.time(20, 0)
        start_times = [office_start_time]
        
        while start_times[-1] < office_end_time:
            next_time = (datetime.datetime.combine(date, start_times[-1]) + timedelta(minutes=15)).time()
            start_times.append(next_time)
        
        start_time = st.selectbox("Start Time:", start_times, index=None)
        current_time = CURRENT_TIME_IST.time()
        
        if start_time:
            if date == current_date and start_time < current_time:
                st.warning("Start time must be in the future.")
            else:
                end_of_day = min(office_end_time, datetime.time(23, 59))
                available_end_times = [
                    datetime.datetime.combine(date, start_time) + timedelta(minutes=i) 
                    for i in range(15, (end_of_day.hour - start_time.hour) * 60 + 1, 15)
                ]
                formatted_end_times = [et.strftime('%H:%M:%S') for et in available_end_times]
                end_time = st.selectbox("End Time:", formatted_end_times, index=None)
                
                if end_time:
                    available_room_options = []
                    for room, capacity in ROOM_CAPACITY.items():
                        if is_room_available(str(date), str(start_time), str(end_time), room):
                            available_room_options.append(f"{room} (Capacity: {capacity})")
                    
                    if not available_room_options:
                        st.warning("No rooms available during this time.")
                    else:
                        st.info("Available Rooms")
                        room_choice = st.selectbox("Select Room:", available_room_options, index=None)
                        
                        if room_choice:
                            st.subheader('Booking Details')
                            selected_room = room_choice.split(" (Capacity: ")[0]
                            description = st.text_input("Meeting Title:")
                            name = st.text_input("Your Name:")
                            email = st.text_input("Your Email:")
                            cc_emails = st.text_input("CC Emails (optional, comma separated):", 
                                                     help="Additional email addresses to receive notifications")
                            
                            if not re.match(r"[^@]+@[^@]+\.[^@]+", email):
                                st.warning("Please enter a valid email address.")
                                return
                            
                            if cc_emails:
                                cc_list = [e.strip() for e in cc_emails.split(',')]
                                for cc_email in cc_list:
                                    if not re.match(r"[^@]+@[^@]+\.[^@]+", cc_email):
                                        st.warning(f"Invalid CC email: {cc_email}")
                                        return
                            
                            if not name or not description:
                                st.warning("All fields are required.")
                            else:
                                if st.button("Confirm Booking"):
                                    booking_id = generate_random_booking_id()
                                    booking_info = {
                                        "booking_id": booking_id,
                                        "date": str(date),
                                        "start_time": str(start_time),
                                        "end_time": str(end_time),
                                        "room": selected_room,
                                        "name": name,
                                        "email": email,
                                        "description": description,
                                        "cc_emails": cc_emails,
                                    }
                                    
                                    booking_data["room_bookings"][booking_id] = booking_info
                                    
                                    if str(date) not in booking_data["room_availability"]:
                                        booking_data["room_availability"][str(date)] = {}
                                    if selected_room not in booking_data["room_availability"][str(date)]:
                                        booking_data["room_availability"][str(date)][selected_room] = []
                                    booking_data["room_availability"][str(date)][selected_room].append(
                                        (str(start_time), str(end_time))
                                    )
                                    
                                    add_booking_to_sheet(booking_info)
                                    
                                    if send_confirmation_email(booking_info):
                                        st.success(f"Booking confirmed! ID: {booking_id}")
                                        st.success("Confirmation email sent.")
                                    else:
                                        st.success(f"Booking confirmed! ID: {booking_id}")
                                        st.warning("Email could not be sent.")

def cancel_room():
    st.header("Cancel Booking")
    booked_rooms = list(booking_data["room_bookings"].values())

    if not booked_rooms:
        st.warning("No existing reservations to cancel.")
        return

    current_datetime = CTIF
    upcoming_reservations = [booking for booking in booked_rooms if is_upcoming(booking, current_datetime)]

    if not upcoming_reservations:
        st.warning("No upcoming bookings to cancel.")
        return

    st.subheader("Select booking to cancel:")
    selected_reservation = st.selectbox(
        "Upcoming Bookings", 
        [f"ID: {booking_id} - {booking_data['room_bookings'][booking_id]['description']} ({booking_data['room_bookings'][booking_id]['date']})" 
         for booking_id in booking_data["room_bookings"].keys() 
         if is_upcoming(booking_data["room_bookings"][booking_id], current_datetime)], 
        index=None
    )

    if selected_reservation:
        booking_id = int(selected_reservation.split("ID: ")[1].split(" - ")[0])
        reservation = booking_data["room_bookings"][booking_id]
        
        st.write(f"**Meeting:** {reservation['description']}")
        st.write(f"**Date:** {reservation['date']}")
        st.write(f"**Time:** {reservation['start_time']} to {reservation['end_time']}")
        st.write(f"**Room:** {reservation['room']}")
        
        user_email = st.text_input("Enter your registered email to confirm cancellation:")
        
        if user_email and st.button("Cancel Booking"):
            if user_email.lower() == reservation["email"].lower():
                # Remove from data structures
                room = reservation["room"]
                date = reservation["date"]
                start_time = reservation["start_time"]
                end_time = reservation["end_time"]

                if date in booking_data["room_availability"] and room in booking_data["room_availability"][date]:
                    booking_data["room_availability"][date][room] = [
                        booking for booking in booking_data["room_availability"][date][room]
                        if (start_time, end_time) != (booking[0], booking[1])
                    ]

                del booking_data["room_bookings"][booking_id]
                
                # Update Google Sheet
                remove_booking_from_sheet(booking_id)
                
                # Send cancellation email
                if send_cancellation_email(reservation):
                    st.success("Booking cancelled successfully.")
                    st.success("Cancellation email sent.")
                else:
                    st.success("Booking cancelled successfully.")
                    st.warning("Cancellation email could not be sent.")
            else:
                st.error("Email does not match booking record.")

def view_reservations():
    st.header("View Bookings")
    booked_rooms = list(booking_data["room_bookings"].values())

    if not booked_rooms:
        st.warning("No existing reservations.")
    else:
        current_datetime = datetime.datetime.strptime(CTIF, '%y-%m-%d %H:%M:%S')
        past_bookings = []
        upcoming_bookings = []

        for booking in booked_rooms:
            if is_upcoming(booking, CTIF):
                upcoming_bookings.append(booking)
            else:
                past_bookings.append(booking)

        past_bookings = sorted(past_bookings, key=lambda x: (x["date"], x["start_time"]))
        upcoming_bookings = sorted(upcoming_bookings, key=lambda x: (x["date"], x["start_time"]))

        tab1, tab2 = st.tabs(["Upcoming Bookings", "Booking History"])
        
        with tab1:
            st.subheader("Upcoming Bookings")
            if not upcoming_bookings:
                st.warning("No upcoming bookings.")
            else:
                upcoming_df = pd.DataFrame(upcoming_bookings)
                upcoming_df = upcoming_df[["booking_id", "date", "start_time", "end_time", "room", "name", "description"]]
                upcoming_df.columns = ["ID", "Date", "Start", "End", "Room", "Booked By", "Meeting"]
                st.dataframe(upcoming_df, hide_index=True)

        with tab2:
            st.subheader("Past Bookings")
            if not past_bookings:
                st.warning("No past bookings.")
            else:
                past_df = pd.DataFrame(past_bookings)
                past_df = past_df[["booking_id", "date", "start_time", "end_time", "room", "name", "description"]]
                past_df.columns = ["ID", "Date", "Start", "End", "Room", "Booked By", "Meeting"]
                st.dataframe(past_df, hide_index=True)

# --- Main App ---
st.title(" SUGAM GROUP ")
st.title("_Meeting_ _Room_ _Booking_ _System_ :calendar:")

date = CURRENT_TIME_IST.date()
time1 = CURRENT_TIME_IST.time()
current_time1 = f"{time1.hour:02d}:{time1.minute:02d}"

st.sidebar.button('Timezone 📍 Asia/Kolkata')


st.sidebar.button(f"Today's Date  \n 🗓️ {date} ")
st.sidebar.button(f'''Current Time ⏰ {current_time1} ''')

menu_choice = st.sidebar.selectbox("Menu", ["Book a Room", "Cancel Booking", "View Bookings"])

if menu_choice == "Book a Room":
    book_room()
elif menu_choice == "Cancel Booking":
    cancel_room()
elif menu_choice == "View Bookings":
    view_reservations()

    
    # ===== FOOTER =====
    st.markdown(
        """
        <div class="footer">
            <img src="https://www.sugamgroup.com/wp-content/uploads/logo-1.png">
            <p>© 2025 Sugam Group </p>
        </div>
        """,
        unsafe_allow_html=True
    )


