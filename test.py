# This is a program to process a CSV of classes and infromation about them into a visual
# representation of an excel spreadsheet.
# In order for the program to work, the csv input must have these exact columns:
#  ________________________________________________________________________________________________
# | Department | Course Num | Section Num | Course Name | M | T | W | R | F | Start Time | End Time|
#  ________________________________________________________________________________________________

# Author: Daniel
# Edited: George Small, 2024

import csv
import xlsxwriter
import math 

# This method takes in the time as a parameter and converts it to the corresponding column index in the excel sheet
def time_to_index(time, end='false'):
    # extract minute and hour from time
    mins = time % 100
    time /= 100
    hours = int(time)
    # normalize minutes to a 15 minute mark
    if mins == 20:
        mins += 10
    if mins == 50:
        mins = 0
        hours += 1
    # get the time by the second of the day
    time = hours*3600 + mins*60
    # the first time is 8:30
    initial = 8*3600 + 30*60  #30,6000
    # add 2 for the initial column offset and divide by 900 to find which

    # Bikal: Added ceiling function to get the indices as ints. Else, the printer won't print out class numbers for some timings.
    return math.ceil((time-initial)/900+2) if end=='false' else math.ceil((time-initial)/900+ 1) #added 2 and 1 to make 1 based like excel

def make_days_ary(infile):
    with open(infile, "r") as file:
        reader = csv.reader(file, delimiter=',')

        #print(reader)

        mon = []
        tues = []
        wed = []
        thurs = []
        fri = []

        daysary = [mon, tues, wed, thurs, fri]
        daystrings = ['M','T','W','R','F']

        colorhex = ['#DAEAF2','#D0C0C0','#C0D0C0','#C0C0D0','#F5F6DC','#D3F8E2','#EBE0F2','#ffecd0','#F2E9E0','#D2E0E1']
        it = 0
        prevname = ''
        line_count = 0

        for row in reader:
            #print(row)
            # If this is the first line in the program (column headers)
            if line_count == 0:
                line_count += 1
                print("FIRST")
               
            # If there are not 11 full columns of data in a row, or there is something weird in column 10
            elif len(row)!= 14:
                line_count += 1
                print("Len:", len(row), "content:", row[10])
                print("ELIF")
            else:
                #print("ELSE")
                # Name is Dept Course - Section
                name = str(row[0]) + ' ' + str(row[1])  + '-' + str(row[2]) #SUB, CRSE, SEC
                # Convert times to index in the excel sheet
                startTime = time_to_index(int(row[12]))
                endTime = time_to_index(int(row[13]), 'true')
                #check if the next line is a lab of the class on the previous line

                # If the name is the same as the previous name, this means this one is a lab
                if name != prevname:
                    it+=1
                prevname = name

                classobj = [name, startTime, endTime, colorhex[it%len(colorhex)]]
                print(classobj)
                print()

                # add the classobj to the day array of every day that class occurs on
                for x in range (7,12):
                    if row[x] != '':
                        #print(classobj)
                        daysary[x-7].append(classobj)

                line_count += 1

        return daysary

def write_worksheet(daysary, outfile):
    daystrings = ['M','T','W','R','F']
    # adds an end index to the days array to indicate that a bottom borderline should be drawn
    for obj in daysary:
        obj.append(['end'])

    workbook = xlsxwriter.Workbook(outfile)
    worksheet = workbook.add_worksheet()

    bottom_cell_format = workbook.add_format()
    bottom_cell_format.set_bottom()

    right_cell_format = workbook.add_format()
    right_cell_format.set_right()

    row = 0
    timelist = []
    hour = 9

    # creates the top row indicating times of the classes
    for ind in range(1,54):
        if ind == 2:
            worksheet.write(row, ind-1, '8:30', bottom_cell_format)
        else:
            if ind%4==0:
                worksheet.write(row, ind-1, str(hour) + ':00', bottom_cell_format)
                hour += 1
            else:
                worksheet.write(row, ind-1, '', bottom_cell_format)

    row+=1

    i = 0 # i is flag to show if this is the first obj in the day ary
    j = 0 # j is a flag to show which day it is

    # iterates over every day ary in daysary
    for ary in daysary:
        # iterates over each classobj in each day array
        for obj in ary:
            color_cell_format0 = workbook.add_format({'bg_color':obj[3]}) if len(obj)>1 else ''
            # iterates over each cell in range and writes the proper entry
            for ind in range (1,55):
                if obj[0]=='end':
                    worksheet.write(row, ind-1,'',bottom_cell_format)
                elif i == 0 and ind == 1:
                     worksheet.write(row,ind-1,daystrings[j],right_cell_format)
                elif ind==obj[1]:
                    worksheet.write(row, ind-1, str(obj[0]), color_cell_format0)
                elif ind>obj[1] and ind<=obj[2]:
                    worksheet.write(row, ind-1, str("*"), color_cell_format0)
                elif ind == 1:
                    worksheet.write(row,ind-1,'',right_cell_format)
            row+=1
            i+=1
        i=0
        j+=1

    workbook.close()


def main():
    infile = './Fall23ScheduleEtest.csv'
    outfile = 'course-output-Fall-2023.xlsx'


    # print(infile)
    # print(" ")
    # print(outfile)

    daysary = make_days_ary(infile)
    write_worksheet(daysary, outfile)

if __name__ == "__main__":
    main()