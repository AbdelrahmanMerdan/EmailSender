# Python library to read and write in Excel
from openpyxl import Workbook, load_workbook
from email.message import EmailMessage
import win32com.client

# SEND EMAIL
wb = load_workbook('emaillist.xlsx')
ws = wb.active
# number of users range
for row in range(2, 131):
    outlook = win32com.client.Dispatch('outlook.application')

    body = """Hello {},<br><br>
    This is the final remainder to please fill out the STARS Primer spreadsheet for this specific courses:<br><br>
    -	{}<br><br>
    The goal of this spreadsheet is to collect information for Seneca’s Sustainability Tracking Assessment & Rating System (STARS) report of 2023 at the course level and create a sustainability course inventory.<br><br>
    As your STARS champion, I am here to support you with answering any of your questions in completing the spreadsheet.<br><br>
    Through this spreadsheet, we ask you to review your courses for any sustainability content and links to the UNSDGs. You are asked to provide responses based on the courses you are teaching during the academic year from 2022 Sept to 2023 Summer. Courses that are offered in multiple sections and/or in multiple semesters will only need to be submitted once. Additionally, courses that are offered in different programs and/or schools will only need to be submitted once. In other words, if a course is taught by multiple faculty members, only the course lead or the main course contact is required to complete the course review.<br><br>
    On average, it may take about 5- 15 minutes to complete.<br><br>
    Access to this spreadsheet can be found <a href="https://seneca.sharepoint.com/:x:/s/STARS/EeKMGd5z0KFFgSFbhY04cnoBzzf5hZCGG_cRtmb-LMZF8w?e=4BLZ6K">HERE</a>. <br><br>
    Submissions for unique courses offered in the 2023 Winter and Summer terms are due <b>August 8th.</b> <br><br>
    Submissions for unique courses offered in the 2022 Fall term are due by <b>September 15th.</b> <br><br>
    A short <a href="https://seneca.sharepoint.com/sites/STARS/_layouts/15/stream.aspx?id=%2Fsites%2FSTARS%2FShared%20Documents%2FGeneral%2FResources%2FSTARS%20Form%20Video%20How%2DTo%2Emp4&ga=1">how-to-video</a> is available to guide you on the steps to complete the Excel spreadsheet. Please note that the form is no longer needed. Additional resources on course examples and FAQ, etc. are provided in the Resource folder <a href="https://seneca.sharepoint.com/sites/STARS/Shared%20Documents/Forms/AllItems.aspx?id=%2Fsites%2FSTARS%2FShared%20Documents%2FGeneral%2FContinuing%20Education&p=true&ga=1">here</a>. <br> <br>
    Your response will help Seneca fulfill our STARS report requirement and contribute to the sustainability course inventory which will provide a baseline for Seneca to assess the college’s offerings for identifying our strengths and opportunities for growth. Also, this sustainability course list will inform current and prospective students when they search for sustainability course offerings at Seneca.<br><br>
    We appreciate your time and support in this initiative. If you have any questions in the process, please do not hesitate to contact me.<br><br>
    
    A quick explanation of how to check if your course includes sustainability. If the course is sustainability focused, then it includes one of these: <br><br>
• Foundational courses with a primary and explicit focus on sustainability (e.g., Introduction to Sustainability).<br>  
• Courses with a primary and explicit focus on the application of sustainability within a field (e.g., Sustainable Agriculture).  <br>
• Courses with a primary and explicit focus on understanding or solving a major sustainability challenge (e.g., Climate Change Science).  <br><br>
However, if the course isn't sustainability focused, it can still be sustainability inclusive if it addresses sustainability in a prominent way. This can be done with one of the following:  <br><br>
•Incorporate a unit or module on sustainability or a sustainability challenge <br>
• Include one or more sustainability-focused activities, OR  <br>
• Integrate sustainability issues and concepts throughout the course. <br><br>
We appreciate your time and support in this initiative. If you have any questions in the process, please email your STARS Champion. <br><br>
Best,<br><br><br>
    
    STARS Champion: <br>

    """.format(str(ws['B' + str(row)].value), str(ws['A' + str(row)].value))

    for account in outlook.Session.Accounts:
        # Email here
        if account.DisplayName == "":

            mail = outlook.CreateItem(0)
            mail.To = str(ws['C' + str(row)].value)

            print(str(ws['C' + str(row)].value))
            mail.Subject = "STARS Primer for Continuing Education Courses " \
                           " Final Reminder"
            mail.HTMLBody = body
            mail.Importance = 2
            mail.Send()

