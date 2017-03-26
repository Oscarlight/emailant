# Email Ant

by Mingda Oscar Li (ml888@cornell.edu)

## Introduction

Email Ant is used to send conference decision (reject/poster/oral) emails. 

## Step-by-step

1. Dependencies:

	python 2.7
	
	python packages: email, smtplib, xlrd. "email" and "smtplib" normally come with python. You need to install xlrd to read excel sheets:
	1. check whether you have pip:
	
		Assume you are using a mac, open a terminal, type: 
		
		```
		which pip
		```
	
		if it return nothing, you need to install pip first; otherwise jump to step iii.
	
	2. Install pip
	
		In the terminal, type:
	
		```
		sudo easy_install pip
		```
	
		it will ask for your password. It will be the same password you used to unlock your computer when logging in.
	
	3. Install xlrd
	
		In the terminal, type:
	
		```
		sudo pip install xrld
		```
	
		In the last line of return message, it should say "Successfully installed xlrd..."
	
2. Allow thrid party access to your Gmail account:

	Go to: https://www.google.com/settings/security/lesssecureapps to allow less secure apps
	
3. Download the code:
	
	Download by click the green Clone or download button on the top right corner. Download it as a zip file ("emailant-master.zip")
	
	Unzip it in your computer.

4. Set up accounts:

	* In sendEmail.py: go to line 47 and 48, insert your gmail account and password (e.g. say your password is 123, then change "password" to "123").
		
5. Customized your email:

	* In sendEmail.py: go to line 52, insert your subject
		
	* In createEmail.py: go to line 56-86, change the email body if needed. (\n is the line breaker)
		
6. Prepared the decision result sheet:

	Create the decision result sheet in the same format of example.xls
	(You may want to first test on a testing decision sheet to make sure all the formats are correct)
	
7. Run the code

	In step 3, most likely, you download to Downloads folder (again, I am assuming you are using a Mac). In the terminal, type:
	
	```
	cd Downloads/emailant-master
	```
	then type:
	```
	python sendEmail.py
	```
8. (Optional) Turn off thrid party access to Gmail account.
	
