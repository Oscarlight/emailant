import xlrd

class createEmail(object):
	def __init__(self, dirName):
		wb = xlrd.open_workbook(dirName)
		self.ws = wb.sheet_by_index(0)
		self.length = self.ws.nrows

	def getCol(self, length, ws, colNum):
		cl = []
		for i in range(1, length):
			cl.append(ws.cell(i, colNum).value.encode('utf-8'))
		return cl		

	def getEmailList(self):
		return self.getCol(self.length, self.ws, 1)

	# def getPgList(self):
	# 	return self.getCol(self.length, self.ws, 0)

	def getIDList(self):
		""" ID type: float """
		cl = []
		for i in range(1, self.length):
			cl.append(str(self.ws.cell(i, 3).value).encode('utf-8'))
		return cl

	def getTitleList(self):
		return self.getCol(self.length, self.ws, 5)

	def getActionList(self):
		return self.getCol(self.length, self.ws, 8)

	# def getSalList(self):
	# 	return self.getCol(self.length, self.ws, 4)

	# def getFirstNameList(self):
	# 	return self.getCol(self.length, self.ws, 5)

	def getLastNameList(self):
		return self.getCol(self.length, self.ws, 7)

	# def getTopicList(self):
	# 	return self.getCol(self.length, self.ws, 7)

	def getVerdictList(self):
		return self.getCol(self.length, self.ws, 4)	

	def getBodyList(self):
		bodyList = []
		emailList = []

		for i in range(0, self.length - 1):
			head = "Dear " + self.getLastNameList()[i] + ", \n"

			if self.getVerdictList()[i] == "Reject" and self.getActionList()[i] == "Send":
				meg = ""
				meg1 = "On behalf of the Device Research Conference Program Committee, I regret to inform you that your paper titled \" " + self.getTitleList()[i] + " \" (id:\" " + self.getIDList()[i][0:9] + " \") was not accepted at the 74th annual Device Research Conference. \n"
				meg2 = "Because of the large number of submissions, many quality papers such as yours could not be incorporated into the technical program this year. We still hope that you and your colleagues will consider attending the meeting at University of Delaware from June 19-22, 2016, held just before the Electronic Materials Conference (EMC). \n"
				meg3 = "Further details about the Device Research Conference including the advance program, registration, visa and accommodations will be found at the conference web site: http://www.deviceresearchconference.org/. \n"
				meg4 = "Sincerely, \nDebdeep Jena \nProfessor Depts. of ECE and MSE \nCornell University \nEmail: djena@cornell.edu \nhttp://djena.engineering.cornell.edu/ \n"
				meg = head + "\n" + meg1 + "\n" + meg2 +"\n"+ meg3 +"\n"+ meg4
				emailList.append(self.getEmailList()[i])
				bodyList.append(meg)

			if self.getVerdictList()[i] == "Oral" and self.getActionList()[i] == "Send":
				meg = ""
				meg1 = "On behalf of the Device Research Conference Program Committee, I am pleased to inform you that your paper titled \" " + self.getTitleList()[i] + " \" (id:\" " + self.getIDList()[i][0:9] + " \") has been accepted for an oral presentation at the 74rd annual Device Research Conference at University of Delaware, held between June 19-22, 2016. \n"
				meg2 = "Each oral presentation is allotted 20 minutes, which includes time for questions. Student presenters are eligible for the Best Student Presentation award. Your talk should utilize PowerPoint or PDF slides. \n"
				meg3 = "Further details about the Device Research Conference including the advance program, registration, visa and accommodations will be found at the conference web site: http://www.deviceresearchconference.org/. \n"
				meg4 = "Sincerely, \nDebdeep Jena \nProfessor Depts. of ECE and MSE \nCornell University \nEmail: djena@cornell.edu \nhttp://djena.engineering.cornell.edu/ \n"
				meg5 = "We look forward to seeing you in University of Delaware this summer. Please contact me for further assistance or information. \n"
				meg = head + "\n" + meg1 + "\n" + meg2 +"\n"+ meg3 +"\n"+ meg5 + "\n" + meg4
				emailList.append(self.getEmailList()[i])
				bodyList.append(meg)

			if self.getVerdictList()[i] == "Poster" and self.getActionList()[i] == "Send":
				meg = ""
				meg1 = "On behalf of the Device Research Conference Program Committee, I am pleased to inform you that your paper titled \" " + self.getTitleList()[i] + " \" (id:\" " + self.getIDList()[i][0:9] + " \") has been accepted for a poster presentation at the 74rd annual Device Research Conference at University of Delaware from June 19-22, 2016. \n"
				meg2 = "The poster presentations have become a highlight of the meeting since their introduction in 2001. Student poster presenters are eligible for the Best Student Poster award. A poster board capable of holding a 4 ft (tall) x 6 ft (wide) poster will be provided for your use at the conference.\n"
				meg3 = "Further details about the Device Research Conference including the advance program, registration, visa and accommodations will be found at the conference web site: http://www.deviceresearchconference.org/. \n"
				meg4 = "Sincerely, \nDebdeep Jena \nProfessor Depts. of ECE and MSE \nCornell University \nEmail: djena@cornell.edu \nhttp://djena.engineering.cornell.edu/ \n"
				meg5 = "We look forward to seeing you in University of Delaware this summer. Please contact me for further assistance or information. \n"
				meg = head + "\n" + meg1 + "\n" + meg2 +"\n"+ meg3 +"\n"+ meg5 + "\n" + meg4
				emailList.append(self.getEmailList()[i])
				bodyList.append(meg)

		return [emailList, bodyList]		


if __name__ == "__main__":
	# (* for testing *)
	c = createEmail("example_decision_sheet.xlsx")
	# print c.getEmailList()
	# print c.getActionList()
	# print c.getLastNameList()
	# print c.getIDList()[0][0:9]
	print c.getBodyList()


