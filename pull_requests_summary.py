import smtplib
import requests
import datetime
import xlsxwriter
from email import encoders
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart

def pull_requests():
    try:
        last_week = datetime.datetime.now() - datetime.timedelta(days=7)
        # Convert datetime to ISO 8601 format
        last_week_iso = last_week.isoformat(timespec='seconds')
        # Get pull requests
        params = {
            "state": "all",
            "sort": "created",
            "direction": "desc",
            "per_page": 100,
            "since": last_week_iso
        }
        headers = {
            "Accept": "application/vnd.github.v3+json"
        }
        response = requests.get(
            f"{base_url}/repos/{repository}/pulls", params=params, headers=headers)
        if response.status_code == 200:
            pull_requests = response.json()
            # Initialize counters
            opened_count = 0
            closed_count = 0
            draft_count = 0
            # Print summary to console
            print("Summary of Pull Requests in the Last Week")
            print("========================================")
            # Create a new Excel workbook
            workbook = xlsxwriter.Workbook('pull_requests.xlsx')
            bold = workbook.add_format({'bold': True})
            sheet1 = workbook.add_worksheet('Open')
            # Create the first sheet
            sheet1.write('A1', 'Pull Request No.', bold)
            sheet1.write('B1', 'Username', bold)
            sheet1.write('C1', 'PR Summary', bold)
            sheet1.write('D1', 'Timestamp', bold)
            # Create the second sheet
            sheet2 = workbook.add_worksheet('Closed')
            sheet2.write('A1', 'Pull Request No.', bold)
            sheet2.write('B1', 'Username', bold)
            sheet2.write('C1', 'PR Summary', bold)
            sheet2.write('D1', 'Timestamp', bold)
            # Create the third sheet
            sheet3 = workbook.add_worksheet('Draft')
            sheet3.write('A1', 'Pull Request No.', bold)
            sheet3.write('B1', 'Username', bold)
            sheet3.write('C1', 'PR Summary', bold)
            sheet3.write('D1', 'Timestamp', bold)
            for pr in pull_requests:
                # print(pr)
                pr_state = pr["state"]
                pr_number = f"{pr['number']}"
                login=f"{pr['user']['login']}"
                timestamp = f"{pr['updated_at']}"
                if(f"{pr['body']}" == 'None'):
                    description = f"{pr['body']}"
                else:
                    description = f"{pr['body'].splitlines()[0]}"

                if pr_state == "open" and pr['draft'] == False:
                    opened_count += 1
                    print("Open Requests :- ")
                    print("PR Number :- ", pr_number)
                    print("Description :- ",description)
                    print("Username :- ", login)
                    print("Timestamp :- ", timestamp)
                    sheet1.write(opened_count, 0, pr_number)
                    sheet1.write(opened_count, 1, login)
                    sheet1.write(opened_count, 2, description)
                    sheet1.write(opened_count, 3, timestamp)
                if pr_state == "closed":
                    closed_count += 1
                    print("Closed Requests :- ")
                    print("PR Number :- ", pr_number)
                    print("Description :- ",description)
                    print("Username :- ", login)
                    print("Timestamp :- ", timestamp)
                    sheet2.write(closed_count, 0, pr_number)
                    sheet2.write(closed_count, 1, login)
                    sheet2.write(closed_count, 2, description)
                    sheet2.write(closed_count, 3, timestamp)
                if pr_state == "open" and pr['draft'] == True:
                    draft_count += 1
                    print("Draft Requests :- ")
                    print("PR Number :- ", pr_number)
                    print("Description :- ",description)
                    print("Username :- ", login)
                    print("Timestamp :- ", timestamp)
                    sheet3.write(draft_count, 0, pr_number)
                    sheet3.write(draft_count, 1, login)
                    sheet3.write(draft_count, 2, description)
                    sheet3.write(draft_count, 3, timestamp)
            workbook.close()
        else:
            print(
                f"Failed to retrieve pull requests. Status code: {response.status_code}")
    except Exception as e:
        print("Error Occured :- ", e)

def send_email():
    try:
        fromaddr = ''
        toaddr = ''
        msecret=''
        message=f"Hi {toaddr}, \nPlease find the attached excel file for the PR summary of open, closed, and draft PRs for github repository https://github.com/{repository}.git"
        msg = MIMEMultipart()
        msg['From'] = fromaddr
        msg['To'] = toaddr
        msg['Subject'] = 'Pull Requests Summary'
        msg.attach(MIMEText(message, 'html'))
        filename = 'pull_requests.xlsx'
        attachment = open('pull_requests.xlsx', "rb")
        p = MIMEBase('application', 'octet-stream')
        p.set_payload((attachment).read())
        encoders.encode_base64(p)
        p.add_header('Content-Disposition', "attachment; filename= %s" % filename)
        msg.attach(p)
        s = smtplib.SMTP('smtp.gmail.com', 587)
        s.starttls()
        s.login(fromaddr, msecret)
        text = msg.as_string()
        s.sendmail(fromaddr, toaddr, text)
        s.quit()
    except Exception as e:
        print("Error Occured :- ", e)

if __name__ == "__main__":
    base_url = "https://api.github.com"
    repository = "idealo/mongodb-performance-test"
    pull_requests()
    send_email()