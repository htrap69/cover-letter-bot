from docxtpl import DocxTemplate
import datetime as dt
import os
from docx2pdf import convert
from geopy.geocoders import Nominatim

class datax:
    passion1 = 'passion1'
    passion2 = 'passion2'
    passion3 = 'passion3'
    field = 'field'
    core1 = 'core1'
    core2 = 'core2'
    project1 = 'project1'
    other = 'other'
    big_com = 'big_com'
    soft_skill1 = "soft_skill1"
    soft_skill2 = "soft_skill2"
    work_exp = "work_exp"
    work_exp1 = "work_exp1"

    def __init__(self, inp):
        self.inp = inp


geolocator = Nominatim(user_agent="geoapiExercises")
f_input = input("Enter the focus of the letter: ")
focus = datax(f_input)
print(focus.inp)

salutation = input("Is name of Hiring manager mentioned? \n")
if salutation == "Yes":
    sal_name = input("Enter the name of manager with pronoun \n")
    salutation = "Dear" + sal_name+','
else:
    salutation = "Dear Hiring Manager,"
print(os.path.abspath('./cover.docx'))


company_name = input("Enter company name: \n")
position_name = input("Enter position name: \n")
address_inp = input("Enter Company Address\n")
location = geolocator.geocode(address_inp)
data = location.raw
loc_data = data['display_name'].split()
company_address2 = loc_data[-4]+' ' + \
    loc_data[-3]+' '+loc_data[-2]+' '+loc_data[-1]
print(company_address2)
add_inp = input("Enter City\n")
company_address1 = add_inp


today_date = dt.datetime.today().strftime('%d %B %Y')
isBig = input("Big Company?\n")
whatField = input("What is the field of the company?\n")
if isBig == 'Yes' or isBig == 'Y':
    focus.big_com = "and a long-time admirer of " + company_name
else:
    focus.big_com = ','
if f_input == 'Java':
    focus.passion1 = 'Java Language and its applications.'
    focus.passion2 = "software development, and commitment towards continuous learning"
    focus.passion3 = "Java, SQL and related concepts"
    focus.field = whatField
    focus.core1 = "object-oriented programming, data structures, algorithms and software design"
    focus.core2 = "supplemented by external courses in Data Stuctures and Algorithms and in School courses on Object Orineted Programming"
    focus.project1 = "Air traffic simulation using Queue, Stack, LinkedList, file handling with FileOutputStream and DataOutputStream to write activity log."
    focus.other = " C, C++ and Python"
    focus.soft_skill1 = "attention to detail"
    focus.soft_skill2 = "understanding client needs"
    focus.work_exp = " During my time at Central Transport, i worked to enter data from 150 bills daily into company software while also keeping the number of errors below 2%"
    focus.work_exp1 = "I also worked at Golden Spoon restaurant, where I applied the skills of empathy, listening attentively and multitasking while taking and preparing orders."

if f_input == 'Python':
    focus.passion1 = ' the Java Language and its applications.'
    focus.passion2 = "automation of tasks,  and commitment towards continuous learning"
    focus.passion3 = "Python language and Artificial Intelligence concepts"
    focus.project1 = "AI assistant named Friday.It uses the pyttsx3 and speech_recognition libraries for text-to-speech and speech-to-text functionality, respectively.It greets the user based on the current time, takes voice commands, and performs various tasks such as opening websites, providing information from Wikipedia, and playing music."
    focus.core1 = "object-oriented programming, data structures, algorithms and software design"
    focus.core2 = "supplemented by external courses on the fundamentals of Artificial Intelligence from Udemy and a course on Python programming at my university"
    if whatField.lower() == "fintech":
        focus.project1 = "I have also worked on several projects, including a virtual desktop assistant and an automated trading bot using Python frameworks like NumPy, Pandas, and Seaborn in combination with the OANDA API."
    focus.other = " C, C++ and Java "
    focus.soft_skill1 = "understanding client needs"
    focus.soft_skill2 = "analytical skills"
    focus.work_exp = " During my time at RG Digital Printing, I worked to make 2000 flyers for a driving school client using Adobe Illustrator while understanding the client's needs and leading a team of 2"
    focus.work_exp1 = "I also worked at at Central Transport, entering data from 150 bills daily into company software while also keeping the number of errors below 2%."

print(today_date)
context = {
    'today_date': today_date,
    'salutation': salutation,
    'company_name': company_name,
    'position_name': position_name,
    'passion1':  focus.passion1,
    'passion2':  focus.passion2,
    'passion3':  focus.passion3,
    'field':  focus.field,
    'core1':  focus.core1,
    'core2':  focus.core2,
    'project1':  focus.project1,
    'other':  focus.other,
    'big_com':  focus.big_com,
    'soft_skill1':  focus.soft_skill1,
    'soft_skill2':  focus.soft_skill2,
    'work_exp': focus.work_exp,
    'work_exp1': focus.work_exp1,
    'company_address1': company_address1,
    'company_address2': company_address2
}
doc = DocxTemplate("cover.docx")
doc.render(context)
wordPath = 'CoverLetter_' + company_name+'_'+position_name+'.docx'
doc.save(wordPath)
pdfPath = wordPath.replace('.docx', '.pdf')
convert(wordPath, pdfPath)
print("Execution Successful")
