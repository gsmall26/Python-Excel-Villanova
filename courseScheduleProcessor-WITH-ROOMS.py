import csv
import xlsxwriter
import sys
import math

# Author: Daniel
# Edited by George Small, 2024

# This method takes in the time as a parameter and converts it to the corresponding column index in the excel sheet
def time_to_index(time, end='false'):
    mins = time%100
    time /= 100
    hours = int(time)
    if mins == 20:
        mins += 10
    if mins == 50:
        mins = 0
        hours += 1
    time = hours*3600 + mins*60
    initial = 8*3600 + 30*60  #30,6000

    # Bikal: Added math ceiling to round off certain time indices. This was creating a bug of not printing out class names
    return math.ceil((time-initial)/900+2) if end=='false' else math.ceil((time-initial)/900+ 1) #added 2 and 1 to make 1 based like excel

# infile = './Input/SPR20-CSC-Submitted-Updated2CSV.csv'      #default
# outfile = './Output/course-output-v3WITH-ROOMS.xlsx'         #default
infile = './Fall23ScheduleE.csv'
outfile = 'course-output-WITH-ROOMS-Fall-2023.xlsx'

with open(infile, "r") as file:
    reader = csv.reader(file, delimiter=',')

    daystrings = ['M','T','W','R','F']

    colorhex = ['#DAEAF2','#D0C0C0','#C0D0C0','#C0C0D0','#F5F6DC','#D3F8E2','#EBE0F2','#FFECD0','#F2E9E0','#D2E0E1']
    it = 0
    prevname = ''
    line_count = 0

    rooms = {}
    readerlist = list(reader)

    first = 'true'
    for row in readerlist:
        if row[3] not in rooms and row[3]!='ROOM' and len(row[3])!=0:
            rooms[row[3]] = [[],[],[],[],[]]

    for row in readerlist:
        if line_count == 0:
            line_count += 1
        elif len(row)!= 14 or row[3] == '':
            line_count += 1
        else:
            name = str(row[0]) + ' ' + str(row[1])  + '-' + str(row[2])
            startTime = time_to_index(int(row[12]))
            endTime = time_to_index(int(row[13]), 'true')
            roomNum = row[3]

            if name != prevname: #check if the next line is a lab of the class on the previous line
                it+=1
            prevname = name

            classobj = [name, startTime, endTime, colorhex[it%len(colorhex)]]

            # add the classobj to the day array of every day that class occurs on
            for x in range (7,12):
                if row[x] != '':
                    rooms[roomNum][x-7].append(classobj)
                    # print(name, '<- name, roomnum -> ', rooms[roomNum])
            line_count += 1

    # adds an end index to the days array to indicate that a bottom borderline should be drawn
    workbook = xlsxwriter.Workbook(outfile)

    bottom_cell_format = workbook.add_format()
    bottom_cell_format.set_bottom()

    right_cell_format = workbook.add_format()
    right_cell_format.set_right()

    botright_cell_format = workbook.add_format()
    botright_cell_format.set_bottom()
    botright_cell_format.set_right()

    for key in rooms:
        worksheet = workbook.add_worksheet(key)

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

        for day in rooms[key]:
            day.append(['end'])

            if len(day) == 1:
                worksheet.write(row,0,daystrings[j],right_cell_format)
                row+=1

            for obj in day:  #obj refers to each classobj in day array
                color_cell_format0 = '' if len(obj)<4 else workbook.add_format({'bg_color':obj[3]})
                for ind in range (1,54):
                    if obj==['end'] and (ind == 1 or ind ==53):
                        worksheet.write(row, ind-1,'',botright_cell_format)
                    elif obj==['end']:
                        worksheet.write(row, ind-1,'',bottom_cell_format)
                    elif i == 0 and ind == 1:
                        worksheet.write(row,ind-1,daystrings[j],right_cell_format)
                    elif ind==obj[1]:
                        worksheet.write(row, ind-1, str(obj[0]), color_cell_format0)
                    elif ind>obj[1] and ind<=obj[2]:
                        worksheet.write(row, ind-1, str("*"), color_cell_format0)
                    elif ind == 1 or ind==53:
                        worksheet.write(row,ind-1,'',right_cell_format)
                row+=1
                i+=1
            i=0
            j+=1

    workbook.close()