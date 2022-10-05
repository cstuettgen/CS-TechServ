INTRODUCTION: 

	Mailer is a fast way to merge data into an E-mail and send E-mail en masse.

	-Add this folder structure to your project:


	[Your Project Directory]	
	   L Mailer.py
		 |
		 |
		 +[Attachments]
		 |	L attachment1.pdf
		 |	L attachment2.xlsx
		 |
		 |
		 +[Config]
 		 |	L credentials.txt
		 | 	L recipients.csv	 
		 |
		 |
		 +[Email_message]
 		 |	L Email.html
		 |
		 |
		 +[Images]
		  	L Image1.jpeg
		    	L Image2.png


EDITING YOUR EMAIL MESSAGE:

	Create your message using HTML and name it 'email.html'

	Edit your html email file to include your header column names from recipients.csv in curly brackets 
    to add the data listed under that header for each recipient in your email body

	
EMBEDDING IMAGES:

	Embed pictures in your HTML email.html using the source "src=cid:imagex" where x 
	starts at 0 for the first picture in alphabetical order in the 'Images' directory. Size as needed using HTML


ADDING ATTACHMENTS:
	
	Add your files to attach to the 'Attachments' directory
	
EDITING RECIPIENTS:

	Fill in each row of 'recipients.csv' as outlined by the headers. Additional headers can be created to pass custom data to 
	the email body by adding the the header name to your HTML in curly brackets. The 'subject' column will also allow 
	for adding the headers in curly brackets to allow for custom subject lines. 


SETTING CREDENTIALS:

	Credentials are set in 'credentials.txt' which is located in the 'Config' directory.
	
	-Contents for 'credentials.txt should be formatted as follows:
		
		Email:		john.doe@wherever.net
		Password:	***************
	

SENDING YOUR EMAILS:

	add 'Import Mailer', and use 'Mailer.notify()' to send emails.

	'Mailer.notify()' will default to using data from the root of the 'Attachments', 'Images', and 'Email_Message' 


SENDING INDIVIDUALIZED EMAILS TO ONE OR MORE RECIPIENTS
	
	create identically named folders in 'Attachments', 'Images', and 'Email_Message'.

	In the 'recipients.csv' file, modify the recipient with the name of this folder under the 'directory' header

	Add your files to be used to the respective new folders you created in each directory

	instead of email.html, your HTML message should be labeled the same as the folders you created.

	-Should look as follows:
	 
	[Your Project Directory]	
	   L Mailer.py
		 |
		 |
		 +[Attachments]
		 |  L +[Some_Company_name]
		 |  	L attachment1.pdf
		 |  	L attachment2.xlsx
		 |
		 |
		 +[Config]
 		 |	L credentials.txt
		 | 	L recipients.csv	 
		 |
		 |
		 +[Email_message]
 		 |  L +[Some_Company_Name]
 		 | 		L Some_Company_Name.html
		 |
		 |
		 +[Images]
 		 |  L +[Some_Company_Name]
		 |		L Image1.jpeg
		 |		L Image2.png


SENDING A HYBRID OF STANDARDIZED EMAILS AND INDIVIDUALIZED EMAILS:

	Mailer will only try to send an individualized email only if there is a directory name listed for the recipient 
    in the 'recipients.csv' in the 'directory' column.

	If the 'directory' column is blank for that recipient an email will be sent 
    from the root of the 'Attachments', 'Images', and 'Email_Message' as outlined with our first example.

	  A hybrid file structure will look like the following:


	[Your Project Directory]	
	   L Mailer.py
		 |
		 |
		 +[Attachments]
 		 |   |	L attachment1.pdf
		 |   |	L attachment2.xlsx
		 |   |	
		 |   +[Some_Company_name]
		 |		L attachment1.pdf
		 |		L attachment2.xlsx
		 |
		 |
		 +[Config]
 		 |	L credentials.txt
		 | 	L recipients.csv	 
		 |
		 |
		 +[Email_message]
 		 |   |	L email.html
		 |   |
		 |   |	
 		 |   +[Some_Company_Name]
 		 |		L Some_Company_Name.html
		 |
		 |
		 +[Images]
 		 |   |	L Image1.jpeg
		 |   |	L Image2.png
		 |   |	
 		 |   +[Some_Company_Name]
		 |		L Image1.jpeg
		 |		L Image2.png


NOTE:

	The script will exit with a message if a directory is listed in 'recipients.csv' and one or 
    more matching directories is not found in 'Attachments', 'Images', and 'Email_Message'.
