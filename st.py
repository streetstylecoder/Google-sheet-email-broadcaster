import streamlit as st
import pandas as pd
from st_aggrid import AgGrid, GridOptionsBuilder
import re
from io import StringIO
import requests
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from urllib.parse import urlparse
import os
import time

# Set page config
st.set_page_config(page_title="Email Broadcaster", layout="wide")

# Custom CSS for the app
st.markdown("""
<style>
    .dragbox {
        padding: 5px;
        margin: 5px;
        background-color: var(--secondary-background-color);
        border: 1px solid var(--primary-color);
        border-radius: 5px;
        cursor: pointer;
        display: inline-block;
    }
    .dragbox:hover {
        background-color: var(--primary-color);
        color: white;
    }
    .editor-container {
        display: flex;
        gap: 20px;
    }
    .column-list {
        flex: 1;
        padding: 10px;
        background-color: var(--secondary-background-color);
        border-radius: 5px;
        max-width: 200px;
    }
    .editor-area {
        flex: 3;
    }
    .stTextArea textarea {
        font-size: 14px;
        font-family: monospace;
    }
    .preview-box {
        background-color: var(--secondary-background-color);
        padding: 15px;
        border-radius: 5px;
        margin: 10px 0;
        border: 1px solid var(--primary-color);
    }
    .status-success {
        color: #00c853;
        font-weight: bold;
    }
    .status-error {
        color: #ff1744;
        font-weight: bold;
    }
    .status-pending {
        color: #ffd600;
        font-weight: bold;
    }
    .column-container {
        display: flex;
        flex-wrap: wrap;
        gap: 8px;
        margin-bottom: 15px;
        padding: 10px;
        border-radius: 5px;
        background-color: var(--secondary-background-color);
    }
    .stButton button {
        padding: 2px 10px;
        font-size: 14px;
        height: auto;
        white-space: nowrap;
    }
</style>
""", unsafe_allow_html=True)

def download_drive_file(drive_link):
    """Download file from Google Drive"""
    try:
        if 'drive.google.com' in drive_link:
            if '/file/d/' in drive_link:
                file_id = drive_link.split('/file/d/')[1].split('/')[0]
            else:
                return (None, "Unsupported Drive link format")

            session = requests.Session()
            try:
                response = session.get(
                    f"https://drive.google.com/uc?id={file_id}&export=download",
                    stream=True,
                    timeout=10
                )
                response.raise_for_status()

                token = None
                for key, value in response.cookies.items():
                    if key.startswith('download_warning'):
                        token = value
                        break

                if token:
                    response = session.get(
                        f"https://drive.google.com/uc?id={file_id}&export=download&confirm={token}",
                        stream=True,
                        timeout=10
                    )
                    response.raise_for_status()

                content_disposition = response.headers.get('content-disposition', '')
                filename = 'attachment'
                if 'filename=' in content_disposition:
                    filename = re.findall('filename="(.+)"', content_disposition)[0]
                else:
                    orig_filename = drive_link.split('/')[-2] if '/file/d/' in drive_link else 'attachment'
                    filename = orig_filename + '.pdf'

                return (response.content, filename)

            except requests.RequestException as e:
                return (None, f"Download failed: {str(e)}")

    except Exception as e:
        return (None, f"Error processing drive link: {str(e)}")

    return (None, "Unknown error occurred")

def send_email(sender_email, app_password, recipient_email, cc_recipients, subject, body, attachment_url=None):
    """Send email using Gmail SMTP"""
    try:
        msg = MIMEMultipart('alternative')
        msg['From'] = sender_email
        msg['To'] = recipient_email
        if cc_recipients:
            msg['Cc'] = cc_recipients
        msg['Subject'] = subject

        # Plain text version
        text_part = MIMEText(
            body.replace('<br>', '\n')
                .replace('<a href="', '\n')
                .replace('">', ': ')
                .replace('</a>', '')
                .replace('<b>', '')
                .replace('</b>', '')
                .replace('<i>', '')
                .replace('</i>', ''),
            'plain'
        )
        
        # HTML version
        html_part = MIMEText(body, 'html')
        
        msg.attach(text_part)
        msg.attach(html_part)

        if attachment_url:
            try:
                content, filename = download_drive_file(attachment_url)
                if content:
                    file_ext = filename.split('.')[-1].lower()
                    mime_types = {
                        'pdf': 'application/pdf',
                        'doc': 'application/msword',
                        'docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                        'jpg': 'image/jpeg',
                        'jpeg': 'image/jpeg',
                        'png': 'image/png',
                        'txt': 'text/plain'
                    }
                    mime_type = mime_types.get(file_ext, 'application/octet-stream')
                    
                    attachment = MIMEApplication(content, _subtype=file_ext)
                    attachment.add_header(
                        'Content-Disposition', 
                        'attachment', 
                        filename=filename
                    )
                    msg.attach(attachment)
            except Exception as e:
                return False, f"Attachment error: {str(e)}"

        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(sender_email, app_password)
        
        all_recipients = [recipient_email]
        if cc_recipients:
            all_recipients.extend([email.strip() for email in cc_recipients.split(',')])
        
        server.sendmail(sender_email, all_recipients, msg.as_string())
        server.quit()
        return True, "Email sent successfully"
    except Exception as e:
        return False, f"Send error: {str(e)}"

def get_sheet_id(url):
    """Extract the sheet ID from the Google Sheets URL"""
    try:
        if '/d/' in url:
            sheet_id = url.split('/d/')[1].split('/')[0]
        elif 'key=' in url:
            sheet_id = url.split('key=')[1].split('&')[0]
        else:
            raise ValueError("Invalid Google Sheets URL format")
        return sheet_id
    except Exception as e:
        st.error(f"Error extracting sheet ID: {str(e)}")
        return None

def load_google_sheet(sheet_url):
    """Load data from public Google Sheet"""
    try:
        sheet_id = get_sheet_id(sheet_url)
        if sheet_id:
            csv_url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv"
            df = pd.read_csv(csv_url)
            return df
    except Exception as e:
        st.error(f"Error loading Google Sheet: {str(e)}")
        return None

def load_excel(file):
    """Load data from uploaded Excel file"""
    try:
        df = pd.read_excel(file)
        return df
    except Exception as e:
        st.error(f"Error loading Excel file: {str(e)}")
        return None

def display_aggrid(df):
    """Display dataframe using AG Grid"""
    gb = GridOptionsBuilder.from_dataframe(df)
    gb.configure_pagination(paginationAutoPageSize=True)
    gb.configure_side_bar()
    gb.configure_default_column(
        groupable=True,
        value=True,
        enableRowGroup=True,
        aggFunc="sum",
        editable=True,
        floatingFilter=True
    )
    gb.configure_selection('multiple', use_checkbox=True)
    grid_options = gb.build()
    
    return AgGrid(
        df,
        gridOptions=grid_options,
        enable_enterprise_modules=True,
        height=400,
        width='100%',
        theme='streamlit',
        allow_unsafe_jscode=True,
        reload_data=True
    )

def get_placeholder_columns(text):
    """Extract placeholders from text enclosed in curly braces"""
    return re.findall(r'\{([^}]+)\}', text)

def replace_placeholders(text, row_data, columns):
    """Replace placeholders with actual values from the row"""
    replaced_text = text
    for col in columns:
        placeholder = f"{{{col}}}"
        if col in row_data:
            replaced_text = replaced_text.replace(placeholder, str(row_data[col]))
    return replaced_text

def create_text_editor(df, key_prefix, label=""):
    """Create a text editor with clickable column names"""
    if f"text_{key_prefix}" not in st.session_state:
        st.session_state[f"text_{key_prefix}"] = ""

    button_cols = st.columns(len(df.columns))
    
    text = st.text_area(
        label,
        value=st.session_state[f"text_{key_prefix}"],
        height=200,
        key=f"editor_{key_prefix}",
        help="Click fields above to insert them into your text. You can use HTML tags."
    )
    
    st.session_state[f"text_{key_prefix}"] = text
    
    for idx, col in enumerate(df.columns):
        with button_cols[idx]:
            if st.button(col, key=f"btn_{key_prefix}_{col}"):
                current_text = st.session_state[f"text_{key_prefix}"]
                st.session_state[f"text_{key_prefix}"] = f"{current_text} {{{col}}}"
                st.rerun()
    
    return text

def create_email_broadcaster(df):
    """Main email broadcaster interface"""
    st.markdown("### 📧 Email Broadcaster")
    
    with st.expander("Email Configuration"):
        sender_email = st.text_input("Gmail Address:", placeholder="your.email@gmail.com")
        app_password = st.text_input("App Password:", type="password", 
                                   help="Use an App Password from Google Account settings")
        
        st.markdown("""
        **How to get App Password:**
        1. Enable 2-Factor Authentication in your Google Account
        2. Go to Google Account → Security → App Passwords
        3. Generate a new App Password for this application
        """)
    
    email_column = st.selectbox(
        "Select the column containing email addresses:",
        options=df.columns
    )
    
    unique_emails = df[email_column].unique().tolist()
    st.write(f"Total unique emails found: {len(unique_emails)}")
    
    email_selection = st.radio(
        "Select emails to send to:",
        ["All emails", "Select specific emails"]
    )
    
    if email_selection == "All emails":
        selected_emails = unique_emails.copy()
    else:
        selected_emails = st.multiselect(
            "Choose specific emails:",
            options=unique_emails,
            default=[]
        )
    
    cc_recipients = st.text_input(
        "CC Recipients (comma-separated emails):"
    )
    
    st.markdown("### Email Content")
    
    subject_template = create_text_editor(
        df, 
        "subject", 
        "Subject Line:"
    )
    
    body_template = create_text_editor(
        df, 
        "body", 
        "Email Body:"
    )
    
    has_attachments = st.checkbox("Include attachments from Google Drive?")
    attachment_column = None
    if has_attachments:
        attachment_column = st.selectbox(
            "Select the column containing Google Drive links:",
            options=df.columns
        )
        st.info("Make sure the Drive links are publicly accessible")
    
    if "email_status" not in st.session_state:
        st.session_state.email_status = {}
    
    col1, col2 = st.columns(2)
    preview_button = col1.button("Generate Email Previews")
    send_button = col2.button("Send Emails")

    if preview_button or send_button:
        if len(selected_emails) == 0:
            st.warning("Please select at least one email recipient.")
            return
            
        if send_button and (not sender_email or not app_password):
            st.error("Please provide Gmail credentials to send emails.")
            return

        st.markdown("### 📩 Email Status and Previews")
        
        subject_placeholders = get_placeholder_columns(subject_template)
        body_placeholders = get_placeholder_columns(body_template)
        all_placeholders = list(set(subject_placeholders + body_placeholders))
        
        invalid_placeholders = [p for p in all_placeholders if p not in df.columns]
        if invalid_placeholders:
            st.error(f"Invalid placeholders found: {', '.join(invalid_placeholders)}")
            return

        status_containers = {}
        for email in selected_emails:
            status_containers[email] = st.empty()
            st.session_state.email_status[email] = "Pending"
        
        if send_button:
            progress_bar = st.progress(0)
            status_text = st.empty()
        
        for idx, email in enumerate(selected_emails):
            if send_button:
                st.session_state.email_status[email] = "Processing"
                status_containers[email].markdown(f"📧 **{email}**: *Processing...*")
                status_text.text(f"Processing email {idx + 1} of {len(selected_emails)}")
            
            row_data = df[df[email_column] == email].iloc[0]
            final_subject = replace_placeholders(subject_template, row_data, all_placeholders)
            final_body = replace_placeholders(body_template, row_data, all_placeholders)
            
            attachment_link = None
            attachment_info = ""
            if has_attachments and attachment_column:
                drive_link = row_data[attachment_column]
                content, result = download_drive_file(drive_link)
                
                if content is not None:
                    attachment_link = drive_link
                    attachment_info = f"**Attachment:** \n- Filename: {result}\n- Original Link: {drive_link}"
                else:
                    attachment_info = f"**Attachment Error:** \n- Error: {result}\n- Original Link: {drive_link}"
                    st.warning(f"Attachment issue for {email}: {result}")
                    
            if send_button:
                success, message = send_email(
                    sender_email,
                    app_password,
                    email,
                    cc_recipients,
                    final_subject,
                    final_body,
                    attachment_link
                )
                
                if success:
                    status = f"✅ **{email}**: Sent successfully"
                    st.session_state.email_status[email] = "Success"
                else:
                    status = f"❌ **{email}**: Failed - {message}"
                    st.session_state.email_status[email] = "Failed"
                
                status_containers[email].markdown(status)
                progress_bar.progress((idx + 1) / len(selected_emails))
                time.sleep(0.5)
            
            with st.expander(f"Preview for: {email} - {st.session_state.email_status[email]}"):
                st.markdown("**To:** " + email)
                if cc_recipients:
                    st.markdown("**CC:** " + cc_recipients)
                st.markdown("**Subject:** " + final_subject)
                st.markdown("**Body:**")
                st.markdown(final_body, unsafe_allow_html=True)
                
                if has_attachments and attachment_column:
                    st.markdown(attachment_info)
        
        if send_button:
            status_text.text("Finished processing all emails!")
            
            success_count = sum(1 for status in st.session_state.email_status.values() if status == "Success")
            fail_count = sum(1 for status in st.session_state.email_status.values() if status == "Failed")
            
            st.markdown(f"""
            ### Summary
            - ✅ Successfully sent: {success_count}
            - ❌ Failed: {fail_count}
            - 📧 Total processed: {len(selected_emails)}
            """)
            
            if success_count == len(selected_emails):
                st.balloons()

def main():
    st.title("📊 Sheet Data Email Broadcaster")
    
    tab1, tab2 = st.tabs(["📝 Google Sheets", "📎 Excel Upload"])
    
    with tab1:
        sheets_url = st.text_input(
            "Enter Google Sheets URL:",
            placeholder="Paste your Google Sheets URL here..."
        )
        
        if sheets_url:
            with st.spinner('Loading data from Google Sheets...'):
                df = load_google_sheet(sheets_url)
                if df is not None:
                    st.success("✅ Google Sheet loaded successfully!")
                    st.markdown("### Data Preview")
                    grid_response = display_aggrid(df)
                    create_email_broadcaster(df)
    
    with tab2:
        uploaded_file = st.file_uploader(
            "Choose an Excel file",
            type=['xlsx', 'xls']
        )
        
        if uploaded_file is not None:
            with st.spinner('Loading data from Excel file...'):
                df = load_excel(uploaded_file)
                if df is not None:
                    st.success("✅ Excel file loaded successfully!")
                    st.markdown("### Data Preview")
                    grid_response = display_aggrid(df)
                    create_email_broadcaster(df)

    st.markdown("""
    ---
    Made with ❤️ using Streamlit
    """)

if __name__ == "__main__":
    main()
