import paramiko
import os
import time
from datetime import datetime
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.utils import formatdate
from email import encoders
import getpass
import sys

def run_remote_powershell_and_email(
    server_ip,
    server_port,
    username,
    password,
    powershell_script,
    network_share_path,
    email_recipients,
    smtp_server,
    smtp_port,
    email_sender,
    email_password
):
    """
    Connects to remote Windows server, executes PowerShell script,
    gets the Excel file, and emails it to colleagues.
    """
    # Generate filename with date
    today = datetime.now().strftime('%Y%m%d')
    excel_filename = f"AD_Machines_Report_{today}.xlsx"
    local_file_path = os.path.join(os.getcwd(), excel_filename)
    
    try:
        # Connect to the Windows server using SSH (requires OpenSSH on Windows server)
        print(f"Connecting to {server_ip}:{server_port}...")
        ssh = paramiko.SSHClient()
        ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
        ssh.connect(hostname=server_ip, port=server_port, username=username, password=password)
        
        # Execute the PowerShell script remotely
        print("Executing PowerShell script on remote server...")
        # Assuming script outputs file to a specific location
        remote_script_command = f"powershell.exe -ExecutionPolicy Bypass -File {powershell_script}"
        stdin, stdout, stderr = ssh.exec_command(remote_script_command)
        
        # Wait for script execution to complete
        exit_status = stdout.channel.recv_exit_status()
        if exit_status != 0:
            error = stderr.read().decode()
            print(f"Error executing PowerShell script: {error}")
            ssh.close()
            return False
        
        print("PowerShell script executed successfully")
        
        # Copy file to network share (assuming PowerShell script outputs to a known location)
        # This command copies from PowerShell script output to network share
        remote_output_path = f"C:\\Temp\\{excel_filename}"  # Adjust based on your PowerShell script
        copy_command = f'powershell.exe Copy-Item -Path "{remote_output_path}" -Destination "{network_share_path}\\{excel_filename}" -Force'
        stdin, stdout, stderr = ssh.exec_command(copy_command)
        
        exit_status = stdout.channel.recv_exit_status()
        if exit_status != 0:
            error = stderr.read().decode()
            print(f"Error copying file to network share: {error}")
            ssh.close()
            return False
        
        print(f"File copied to network share: {network_share_path}\\{excel_filename}")
        
        # Now download the file to local machine using SFTP
        sftp = ssh.open_sftp()
        network_file_path = f"{network_share_path}\\{excel_filename}".replace('\\', '/')
        
        print(f"Downloading file to local machine: {local_file_path}")
        sftp.get(network_file_path, local_file_path)
        sftp.close()
        
        # Close SSH connection
        ssh.close()
        
        # Send email with attachment
        send_email_with_attachment(
            send_from=email_sender,
            send_to=email_recipients,
            subject=f"AD Machines Report - {today}",
            message="Please find attached the AD Machines report for this month.",
            file_path=local_file_path,
            server=smtp_server,
            port=smtp_port,
            username=email_sender,
            password=email_password
        )
        
        print("Process completed successfully!")
        return True
        
    except Exception as e:
        print(f"Error: {str(e)}")
        return False

def send_email_with_attachment(send_from, send_to, subject, message, file_path, 
                              server, port, username=None, password=None):
    """Send an email with an attachment"""
    
    # Create message container
    msg = MIMEMultipart()
    msg['From'] = send_from
    msg['To'] = ", ".join(send_to) if isinstance(send_to, list) else send_to
    msg['Date'] = formatdate(localtime=True)
    msg['Subject'] = subject

    # Add message body
    msg.attach(MIMEText(message))

    # Add attachment
    with open(file_path, "rb") as file:
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(file.read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', f'attachment; filename="{os.path.basename(file_path)}"')
    msg.attach(part)

    # Send email
    try:
        smtp = smtplib.SMTP(server, port)
        smtp.ehlo()
        smtp.starttls()
        smtp.ehlo()
        
        if username and password:
            smtp.login(username, password)
            
        smtp.sendmail(send_from, send_to, msg.as_string())
        smtp.close()
        print("Email sent successfully!")
    except Exception as e:
        print(f"Failed to send email: {str(e)}")

if __name__ == "__main__":
    # You can set these as environment variables or hardcode them (less secure)
    # Or prompt for input each time
    
    # Server connection details
    server_ip = input("Enter AD server IP address: ")
    server_port = int(input("Enter server port (default is 22): ") or "22")
    username = input("Enter server username: ")
    password = getpass.getpass("Enter server password: ")
    
    # PowerShell script location on remote server
    powershell_script = input("Enter path to PowerShell script on server: ")
    
    # Network share path
    network_share_path = input("Enter network share path (e.g., '\\\\server\\share'): ")
    
    # Email settings
    email_recipients = input("Enter email recipients (comma-separated): ").split(',')
    smtp_server = input("Enter SMTP server (e.g., smtp.office365.com): ")
    smtp_port = int(input("Enter SMTP port (default is 587): ") or "587")
    email_sender = input("Enter sender email: ")
    email_password = getpass.getpass("Enter email password: ")
    
    # Run the main function
    run_remote_powershell_and_email(
        server_ip=server_ip,
        server_port=server_port,
        username=username,
        password=password,
        powershell_script=powershell_script,
        network_share_path=network_share_path,
        email_recipients=email_recipients,
        smtp_server=smtp_server,
        smtp_port=smtp_port,
        email_sender=email_sender,
        email_password=email_password
    )