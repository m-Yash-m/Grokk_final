import pandas as pd
import requests
import json
import re
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import streamlit as st
from dotenv import load_dotenv
import gspread
from google.oauth2 import service_account
import time
import os
import toml

# Load environment variables
load_dotenv(override=True)

# Google Sheets API setup
SCOPE = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/spreadsheets",
         "https://www.googleapis.com/auth/drive.file", "https://www.googleapis.com/auth/drive"]

API_KEY = os.getenv('API_KEY')
SENDER_EMAIL = os.getenv('SENDER_EMAIL')
SENDER_PASSWORD = os.getenv('SENDER_PASSWORD')

# Load the secrets from the TOML file
secrets = toml.load(".streamlit/secrets.toml")

# Extract the service account key JSON string
service_account_json = secrets["google_cloud"]["service_account_key"]

# Parse the JSON string into a dictionary
service_account_info = json.loads(service_account_json)

# Helper function to get Google Sheets service
def get_gsheet_service():
    # Create credentials from the parsed JSON content
    creds = service_account.Credentials.from_service_account_info(service_account_info, scopes=SCOPE)
    
    # Authorize the client
    client = gspread.authorize(creds)
    
    return client

# Function to evaluate a candidate and generate email with Grok API
def evaluate_and_generate_email(candidate_data: str, cohort_details: dict) -> str:
    prompt = f"""
    You are an AI designed to create warm and professional invitation emails for an entrepreneurial cohort program. 
    This program provides experiential learning, expert mentorship, and a launchpad into the startup ecosystem.

    Based on the candidate data provided below and the cohort program details, generate a formal yet friendly invitation email. 
    Use emojis sparingly to highlight the event details (e.g., 📍 for location, 📅 for date, ⏰ for time). 
    Note don't use asterisks or stars
    Conclude the email with "Thanks and Regards, 18 Startup".

    Do not include any evaluations or scores in the email. The email should be concise and only focus on inviting the candidate.

    Candidate data:
    {candidate_data}

    Event details:
    Venue: {cohort_details['venue']}
    Date: {cohort_details['date']}
    Time: {cohort_details['time']}
    Description: {cohort_details['description']}
    """

    suitability_payload = {
        "messages": [
            {"role": "system", "content": "You are Grok, an AI designed to generate formal yet friendly invitations."},
            {"role": "user", "content": prompt}
        ],
        "model": "grok-beta",
        "temperature": 0.7
    }

    suitability_response = requests.post("https://api.x.ai/v1/chat/completions", headers={
        "Content-Type": "application/json",
        "Authorization": f"Bearer {API_KEY}"
    }, data=json.dumps(suitability_payload))

    if suitability_response.status_code == 200:
        result = suitability_response.json().get("choices", [{}])[0].get("message", {}).get("content", "")
        return result.strip()
    else:
        return "Error in generating email content."
    
def parse_grok_response(response_text: str) -> dict:
    """Extract rating and reason from Grok's response."""
    rating_match = re.search(r'(\d{1,2})/10', response_text)
    rating = int(rating_match.group(1)) if rating_match else 0
    reason = response_text.split("Reason:", 1)[1].strip() if "Reason:" in response_text else response_text
    return {"rating": rating, "reason": reason} 

def evaluate_candidate_with_grok(content: str, name: str) -> dict:
    """Get rating and evaluation using Grok API."""
    payload = {
        "messages": [
            {
                "role": "system",
                "content": "You are Grok, an AI designed to evaluate candidate suitability."
            },
            {
                "role": "user",
                "content": f"Evaluate the following LinkedIn content for candidate '{name}' on a scale of 1-10 for suitability. Provide the rating and reasoning.\n\n{content}"
            }
        ],
        "model": "grok-beta",
        "temperature": 0.7
    }
    response = requests.post("https://api.x.ai/v1/chat/completions", headers={
        "Content-Type": "application/json",
        "Authorization": f"Bearer {API_KEY}"
    }, data=json.dumps(payload))
    
    if response.status_code == 200:
        result = response.json().get("choices", [{}])[0].get("message", {}).get("content", "")
        evaluation = parse_grok_response(result)
        st.write(f"Grok evaluation for {name}: {evaluation}")
        return evaluation
    st.write(f"Failed to evaluate candidate {name} - Status Code: {response.status_code}")
    return {"rating": 0, "reason": "Error in retrieving data"}
# Function to send emails to candidates
def send_email(email, email_subject, email_content):
    msg = MIMEMultipart()
    msg['From'] = SENDER_EMAIL
    msg['To'] = email
    msg['Subject'] = email_subject
    msg.attach(MIMEText(email_content, 'plain'))

    try:
        with smtplib.SMTP('smtp.gmail.com', 587) as server:
            server.starttls()
            server.login(SENDER_EMAIL, SENDER_PASSWORD)
            server.sendmail(SENDER_EMAIL, email, msg.as_string())
        return f"Email sent successfully to {email}"
    except Exception as e:
        return f"Failed to send email to {email}: {e}"

# Function to add missing columns in Google Sheets
def add_missing_columns(worksheet):
    header = worksheet.row_values(1)  # Get the header row
    missing_columns = ['Status', 'Evaluation Reason', 'Rating']
    
    for col in missing_columns:
        if col not in header:
            worksheet.add_cols(1)  # Add one column if it's missing
            worksheet.update_cell(1, len(header) + 1, col)  # Add the column name in the first row
            header.append(col)  # Update header for further checks

# Main function for chatbot-style interaction
def main():
    st.title("18 Startup Interactive Invitation Automation")

    # Step 1: Google Sheet Configuration
    st.markdown("### Google Sheet Configuration")
    google_sheet_id = st.text_input("What is the Google Sheet ID where the candidate data is stored?")
    if google_sheet_id:
        st.markdown("### Cohort Program Details")

        # Step 2: Cohort Program Details
        cohort_description = st.text_area("Describe the cohort program briefly:")
        if cohort_description:
            venue = st.text_input("Where will the event be held (Venue)?")
            if venue:
                event_date = st.date_input("What is the event date?")
                if event_date:
                    event_time = st.time_input("What time will the event start?")
                    if event_time:
                        st.markdown("### Candidate Evaluation")

                        # Step 3: Candidate Evaluation
                        min_rating = st.slider("What is the minimum rating to invite candidates?", 1, 10, 5)
                        email_subject = st.text_input("What should the email subject be?")

                        if email_subject and st.button("Send Invitations"):
                            st.write("Fetching data from Google Sheets...")
                            try:
                                # Fetch Google Sheets data
                                client = get_gsheet_service()
                                sheet = client.open_by_key(google_sheet_id)
                                worksheet = sheet.get_worksheet(0)
                                data = worksheet.get_all_records()
                                df = pd.DataFrame(data)
                                df = df.map(str)

                                # Add missing columns if necessary
                                add_missing_columns(worksheet)

                                cohort_details = {
                                    "venue": venue,
                                    "date": event_date.strftime('%Y-%m-%d'),
                                    "time": event_time.strftime('%H:%M:%S'),
                                    "description": cohort_description,
                                    "min_rating": min_rating
                                }

                                st.write("Evaluating candidates...")

                                header = worksheet.row_values(1)  # Fetch the header again here, after columns might have been added

                                for idx, candidate in df.iterrows():
                                    # Convert row to Python-native dictionary
                                    candidate_data = json.dumps(candidate.to_dict())
                                    response = evaluate_and_generate_email(candidate_data, cohort_details)
                                    evaluation = evaluate_candidate_with_grok(candidate_data, candidate['Applicant Name'])

                                    rating = evaluation['rating']
                                    df.at[idx, 'Rating'] = rating
                                    df.at[idx, 'Evaluation Reason'] = evaluation['reason']

                                    # Ensure valid email format
                                    email_regex = r"^[a-zA-Z0-9._%+-]+@[a-zAZ0-9.-]+\.[a-zA-Z]{2,}$"
                                    
                                    email = candidate.get('Emails', '').strip()
                                    if re.match(email_regex, email):
                                        if rating >= cohort_details["min_rating"]:
                                            df.at[idx, 'Status'] = "Selected and Email Sent"
                                            email_content = response
                                            st.write(send_email(email, email_subject, email_content))

                                            # Update Google Sheets
                                            row_number = idx + 2  # Row in Google Sheets (1-based index)
                                            worksheet.update_cell(row_number, header.index('Status') + 1, df.at[idx, 'Status'])
                                            worksheet.update_cell(row_number, header.index('Evaluation Reason') + 1, df.at[idx, 'Evaluation Reason'])
                                            worksheet.update_cell(row_number, header.index('Rating') + 1, df.at[idx, 'Rating'])
                                        else:
                                            df.at[idx, 'Status'] = "Not Selected"
                                    else:
                                        st.warning(f"Invalid or missing email for {candidate.get('Applicant Name', 'Unknown')}.")

                                st.success("Process completed!")

                            except Exception as e:
                                st.error(f"Error occurred: {e}")

if __name__ == "__main__":
    main()

