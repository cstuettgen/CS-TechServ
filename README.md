# CS-TechServ
Add Mailer.py to same directory as your code.

Copy 'def notify():' method to the code you are running.
    modify the subject and body-vars to match your subject 
    and order of variables in the body of of your HTML email body.

Create folders 'Config', 'Attachments', 'Email_Message', 'Images'.

Add HTML email body to 'Email_Message' and title it 'email.html'.

Images to be embedded will be labeled in alphabetical order starting with index0.
Add images to your HTML email body, 'email.html'
    Ex.| <img width=469 height=263  src="cid:image0">
       | <img width=469 height=263  src="cid:image1">

Create credentials.txt in 'Config' folder and label creds that will be sending the email 
with Email: and Password: on separate lines.
    Ex.| Email:  john.doe@wherever.net
       | Password:  ***********

Create 'recipients.csv' and place it in the 'Config' folder.
    Add this header to file: 'first_name,last_name,to_email,carbon_copy'.
    Begin adding recipients on line 2 of csv.

Add any files to be attached to the 'Attachments' folder.

add notify() to your processes.
