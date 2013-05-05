from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import NoSuchElementException
import unittest, time, re
import smtplib
from email.mime.text import MIMEText

#####
# Author: Mike Adler - May 3, 2013
#
# GOESInterviewChecker - Used to automate checking of GOES interview times.
#
# This program will log in to your GOES account with the set values below,
# and will read both your current booked interview, and the earliest
# date at your preferred enrollment center. 
# If an earlier date is found, an email message will be sent alerting you
# that an earlier date has been found.
#
# NOTE: If you have no current booking, or wish to specify a date to use
#       for comparison, please set the  compareToDate  value below. Otherwise
#       your current booking date will be found and used for comparison.
#####

class GOESInterviewChecker(unittest.TestCase):
	
	# GOES Account Info
	# Set these values
	# Note: Preferred enrollment center value is the text value in the enrollment center dropdown menu
	GOES_USERNAME = ""
	GOES_PASSWORD = ""
	GOES_BASE_URL = "https://goes-app.cbp.dhs.gov/" 
	GOES_PREFERRED_ENROLLMENT_CENTER = "Buffalo-Ft. Erie Enrollment Center - 10 CENTRAL AVENUE, FORT ERIE, ON L2A1G6, CA"
	
	# Email Info
	# Set these values
	SENDER = "" #sender email address - and login email for smtp server
	EMAIL_PASSWORD = "" #password for smtp server
	TO =  "" #email address to send the notifications
	SUBJECT = "GOES - Earlier Interview Found!" #subject of the email notification
	SMTP_SERVER = "" #smtp server address to use for email sending
	SMTP_PORT = 0 #smtp server port number
	
	# Optionally set this value
	# Note: If set to None, comparison will use current booking date
	#       otherwise it will use this date to compare
	# **Must be in the format: "June 1, 2013 08:00" or set to  None
	compareToDate = None 
	
	# Used to store the current booking date found on GOES page
	currentBookingDate = ""
	
	#########
	# Custom functions to process GOES Website Information
	#########
	def test_g_o_e_s_interview_checker(self):
		driver = self.driver
		driver.get(self.GOES_BASE_URL + "/main/goes")
		driver.find_element_by_id("user").clear()
		driver.find_element_by_id("user").send_keys(self.GOES_USERNAME)
		driver.find_element_by_id("password").clear()
		driver.find_element_by_id("password").send_keys(self.GOES_PASSWORD)
		driver.find_element_by_id("SignIn").click()
		driver.find_element_by_link_text("Enter").click()
		driver.find_element_by_css_selector("img[alt=\"Manage Interview Appointment\"]").click()
		
		# set current booking date raw text value
		currentBookingRaw = driver.find_element_by_xpath("//div[@class='maincontainer']")
		currentBookingRawText = currentBookingRaw.text
		
		driver.find_element_by_name("reschedule").click()
		
		dropdown = driver.find_element_by_id("selectedEnrollmentCenter")
		
		# select preferred enrollment center
		for option in dropdown.find_elements_by_tag_name('option'):
			if option.text == self.GOES_PREFERRED_ENROLLMENT_CENTER:
				option.click() 
				
		driver.find_element_by_css_selector("img[alt=\"Next\"]").click()
		
		# find enrollment center earliest date value
		dateTds = driver.find_elements_by_xpath("//div[@class='maincontainer']/table/tbody/tr/td[contains(.,'Date:')]")
		# get plain text values and store them
		stringDateTds = []
		for td in dateTds:
			stringDateTds.append(td.text)
		
		# log-off of your GOES account
		driver.find_element_by_link_text("Log off").click()
		
		# after logging off now perform all the processing and logic
		# in case an error occurs with the processing, the website
		# will be sure to have already logged off
		
		# parse and set current booking date
		self.currentBookingDate = self.parseCurrentBookingDate(currentBookingRawText)
		
		#parse and check if the enrollment centers have earlier dates
		earlierDates = []
		for stringDate in stringDateTds:
			#print td.text
			if (self.isEarlierDate(stringDate)):	
				earlierDates.append(stringDate)

		# build email message
		emailMessage = '\n'.join(earlierDates)
		
		# if there's earlier dates, send the email
		if (len(earlierDates)):
			print "%d Earlier Date(s) Found!" % len(earlierDates)
			print "Sending Email Message: %s" % emailMessage
			self.sendEmail(emailMessage)
		else:
			print "NO Earlier Dates Found."
		
	def parseCurrentBookingDate(self, currentDateStr):
		dateStr = ""
		
		dateStr = currentDateStr.split("Date:")[1]
		dateStr = dateStr.split("Enrollment Center")[0]
		dateStr = dateStr.replace("Interview Time:", "")
		dateStr = dateStr.strip()
		
		currentDate = time.strptime(dateStr, "%B %d, %Y %H:%M")
		
		return currentDate
		
	def parseAvailDates(self, availDateStr):
		dateString = availDateStr.split("Date:")[1]
		dateString = dateString.split("End Time:")[0]
		dateString = dateString.replace("Start Time:", "")
		dateString = dateString.strip()
			
		return dateString
		
	def getDateForString(self, dateString):
		parsedDateString = self.parseAvailDates(dateString)
		dateStamp = time.strptime(parsedDateString, "%Y-%m-%d, %H%M,")
		
		return dateStamp
		
	def isEarlierDate(self, dateTd):
		containsEarlierDate = False
		
		# check the date against a set date
		availableDate = self.getDateForString(dateTd)
		dateToCompare = ""

		if (self.compareToDate is not None):
			dateToCompare = time.strptime(self.compareToDate, "%B %d, %Y %H:%M")
		else:
			dateToCompare = self.currentBookingDate
		
		if (availableDate < dateToCompare):
			containsEarlierDate = True
		
		return containsEarlierDate
		
	def sendEmail(self, message):
		# Create a text/plain message
		msg = MIMEText(message)

		msg['Subject'] = self.SUBJECT
		msg['From'] = self.SENDER
		msg['To'] = self.TO

		# Send the message via provided server
		session = smtplib.SMTP(self.SMTP_SERVER, self.SMTP_PORT)
		
		# Login to the smtp server
		session.ehlo()
		session.starttls()
		session.ehlo
		session.login(self.SENDER, self.EMAIL_PASSWORD)
		
		session.sendmail(self.SENDER, [self.TO], msg.as_string())
		session.quit()
		
	#######
	# Selenium Functions - DO NOT MODIFY
	######
	def setUp(self):
		self.driver = webdriver.Firefox()
		self.driver.implicitly_wait(30)
		self.verificationErrors = []
		self.accept_next_alert = True
		
	def is_element_present(self, how, what):
		try: self.driver.find_element(by=how, value=what)
		except NoSuchElementException, e: return False
		return True

	def is_alert_present(self):
		try: self.driver.switch_to_alert()
		except NoAlertPresentException, e: return False
		return True

	def close_alert_and_get_its_text(self):
		try:
			alert = self.driver.switch_to_alert()
			alert_text = alert.text
			if self.accept_next_alert:
				alert.accept()
			else:
				alert.dismiss()
			return alert_text
		finally: self.accept_next_alert = True

	def tearDown(self):
		self.driver.quit()
		self.assertEqual([], self.verificationErrors)

######
# Main - initiates the program
######
if __name__ == "__main__":
	unittest.main()
	
