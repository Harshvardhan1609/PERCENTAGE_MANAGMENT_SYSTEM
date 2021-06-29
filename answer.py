import xlsxwriter 
#actual code

#creating excel book
book = xlsxwriter.Workbook('percentage.xlsx')     
sheet = book.add_worksheet()    
row = 4 
column = 7
#entering number of students
number_students = int(input("Please enter number of students\n"))
#Marks entering
session_list =[]
name_list = []
maths_list = []
science_list = []
english_list = []
percent_list = []
grade_list = []

for marks in range(0,number_students):
    name = input("Please enter your name\n")
    name_list.append(name)
    maths = (input("Please enter maths marks\n"))
    mathsli = int(maths)
    maths_list.append(maths)
    science =(input("Please enter science marks\n"))
    scienceli = int(science)
    science_list.append(science)
    english =(input("Please enter english marks\n"))
    englishli = int(english)
    english_list.append(english)
    #code for calculation of precentage
    percent = (mathsli + scienceli + englishli)/3
    percentli = int((mathsli + scienceli + englishli)/3)
    percent_list.append(percentli)
    #Condition checking
    if percent >= 90:
        print(f'Name: {name}') 
        grade = 'A'
        print("Grade: A\n")
        grade_list.append(grade)
    elif percent < 90 and percent > 70 :
        print(f'Name: {name}') 
        grade = 'B'
        print("Grade: B\n")
        grade_list.append(grade)
    elif percent < 70 and percent > 50 :
        print(f'Name: {name}') 
        grade = 'C'
        print("Grade: C\n")
        grade_list.append(grade)
    elif percent < 50 and percent > 30 :
        print(f'Name: {name}') 
        grade = 'D'
        print("Grade: D\n")
        grade_list.append(grade)
    else:
        print(f'Name: {name}') 
        grade = 'E'
        print("Grade: E\n")
        grade_list.append(grade)
    sheet.write(marks,0,marks+1)
    sheet.write(marks,1,name_list[marks])
    sheet.write(marks,2,maths_list[marks])
    sheet.write(marks,3,english_list[marks])
    sheet.write(marks,4,science_list[marks])
    sheet.write(marks,5,percent_list[marks])
    sheet.write(marks,6,grade_list[marks])
    
#greeting messages
book.close()
print("Thanks for using grade calculator")