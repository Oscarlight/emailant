# Email Ant

by Mingda Oscar Li (ml888@cornell.edu)

## Introduction

	Email Ant is used to send conference decision (reject/poster/oral) emails. 

## Step-by-step

1. Dependencies:

	python 2.7
	python packages: email, smtplib, xlrd
	
2. Allow thrid party access to your Gmail account:

	Go to: https://www.google.com/settings/security/lesssecureapps
	Allow less secure apps

3. Set up accounts:

	In sendEmail.py:
		line 47 and 48, insert your gmail account and password
		
4. Customized your email:

	In sendEmail.py:
		line 52, insert your subject
		
	In createEmail.py:
		line 56-86, change the email body if needed.
		
5. Prepared the decision result sheet:

	Create the decision result sheet in the same format of example.xls
	(You may want to first test on a testing decision sheet to make sure all the formats are correct)
	
6. Run the code 

	run python sendEmail.py
	
7. (Optional) Turn off thrid party access to Gmail account.
	
