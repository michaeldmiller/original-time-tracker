#import necessary libraries
import matplotlib.pyplot as pyplot
import matplotlib.patches as patches
import numpy as np
import time
import calendar
import colorsys
import copy

from matplotlib import rcParams
from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.section import WD_SECTION
from docx.shared import Inches
from docx.shared import Pt

from openpyxl import Workbook
from openpyxl import load_workbook

def colorCategorySorter(category):
    """Takes activity category and generates its corresponding color
        str->tuple"""
    colorDict = {"sleep": [0, 29, 105], "entertainment": [255, 48, 48], "career": [10, 75, 158],
                  "outdoor": [255, 220, 90], "project": [15, 251, 251], "socialization": [255, 99, 196],
                  "grooming": [166, 255, 123], "culinary": [255, 191, 30], "relaxation": [91, 173, 68],
                  "fitness": [255, 251, 67], "housework": [191, 131, 0], "maintenance": [118, 55, 0],
                  "education": [155, 50, 170], "volunteering": [140, 20, 86], "transit": [218, 218, 218],
                  "caregiving": [95, 10, 10], "other": [80, 180, 155], "missing": [87, 87, 87]}
    intermediateColor = colorDict[category]
    intermediateColorCounter = 0
    while intermediateColorCounter < 3:
        if intermediateColor[intermediateColorCounter] == 0:
            pass
        else:
            intermediateColor[intermediateColorCounter] = intermediateColor[intermediateColorCounter]/255
        intermediateColorCounter +=1
    return intermediateColor


def innerColorCreator(rGBOutput, lightenVal = 0.7):
    """Takes colorCategorySorter RGB Tuple, adds lightened alpha channel
        tup->tup"""
    startR = rGBOutput[0]
    startG = rGBOutput[1]
    startB = rGBOutput[2]
    startHSV = list(colorsys.rgb_to_hsv(startR, startG, startB))
    startHSV[1] = startHSV[1]*lightenVal
    newH = startHSV[0]
    newS = startHSV[1]
    newV = startHSV[2]
    newRGBOutput = colorsys.hsv_to_rgb(newH, newS, newV)
    return newRGBOutput

def hour_and_minute_converter(minutes):
    """takes a number of minutes and converts it into hours and minutes
        int->str"""
    minutes = int(minutes)
    nHours,nMinutes = minutes//60, minutes%60
    nHours = str(nHours)
    nMinutes = str(nMinutes)
    if len(nMinutes) < 2:
        nMinutes = "0"+nMinutes
    timeOutput = nHours+ ":"+ nMinutes
    return timeOutput

def clock_time_to_minute_converter(clock_time):
    """takes a clock time of format HH:MM and converts it to number of
        minutes on a 0-1439 scale
        str->int"""
    clock_time = str(clock_time)
    if len(clock_time) < 5:
        clock_time = "0" + clock_time
    nHours = int(clock_time[0:2])
    nMinutes = int(clock_time[3:5])
    minute_number = (nHours *60) + nMinutes
    return minute_number

def report_day_converter(report_day):
    """takes report day (str in format: YYYY/MM/DD) and converts to
        format: YYYY-MM-DD
        str->str"""
    newReportDay = report_day.replace("/", "-")
    return newReportDay

def date_decrement(date):
    """takes a date and returns the previous date
    str->str"""
    date_intermediate = date.replace("-","")
    date_year_check = date_intermediate[4:8]
    if date_year_check == "0101":
        old_year = int(date_intermediate[0:4])
        new_year = str(old_year - 1)
        date_new_year_output = new_year + "-12-31"
        return date_new_year_output
    date_month_check = date_intermediate[6:8]
    if date_month_check == "01":
        old_month = int(date_intermediate[4:6])
        new_month = str(old_month - 1)
        if len(new_month) == 1:
            new_month = "0" + new_month
        #find new date (this one's a bitch)
        months31 = ["01", "03", "05", "07", "08", "10", "12"]
        months30 = ["04", "06", "09", "11"]

        if new_month in months31:
            new_date = "31"
        if new_month in months30:
            new_date = "30"
        if new_month == "02":
            if int(date_intermediate[0:4])/4 == int(date_intermediate[0:4])//4:
                #print("leap year")
                if int(date_intermediate[0:4])/100 == int(date_intermediate[0:4])//100:
                    new_date = "28"
                else:
                    new_date = "29"
            else:
                new_date = "28"

        date_new_month_output = date_intermediate[0:4] + "-" + new_month + "-" + new_date
        return date_new_month_output
    else:
        old_date = int(date_intermediate[6:8])
        new_date = str(old_date - 1)
        if len(new_date) == 1:
            new_date = "0" + new_date
        date_output = date_intermediate[0:4] + "-" + date_intermediate[4:6] + "-" + new_date
        return date_output

#data acquisition

print("""
This program is a less stable, custom review oriented companion to the
time tracker program. Using the rawData.xlsx file, it will generate aggregate
review reports for long stretches of time. Based on a start date, given in
YYYY-MM-DD, and a number of days (which it will go back from the start date)
it will generate an averaged pie chart, along with a full timeline.
The resulting charts will be outputted as images by the program.

Important to understand about this program is that it is /not/ idiot-proofed,
and edge cases are not handled. This program should be considered unstable until
further notice (2020/07/21)\n""")

#the two inputs are handled. If they request data that does not exist, however,
#the program will crash.

print("""
Please input the ending date for the time frame to be considered by this program:
(format as YYYY-MM-DD)\n""")

while True:
    report_day = str(input())
    try:
        report_day_intermediate = report_day.replace("-"," ")
        report_day_object = time.strptime(report_day_intermediate, "%Y %m %d")
        break
    except:
        print("I'm sorry, that input was not valid\n")
        continue

print("""
Now, enter the number of days back you want the program to go from the end date:
(format as a single integer)\n""")

while True:
    number_of_days = input()
    try:
        number_of_days = int(number_of_days)
        break
    except:
        print("I'm sorry, that input was not valid\n")


print("""
Thank you, the aggregate reports will now be created.""")

#this concludes the user facing portion of the program. Onto calculations.

###the following section is for testing
##report_day = "2020-07-10"
##number_of_days = 2
###/testing section

def add_hours_and_minutes(time1, time2):
    """takes two times, strings of format HH:MM, and adds them together
        str->str"""
    minutes_together = int(time1[3:5]) + int(time2[3:5])
    carried_hours = 0
    if minutes_together >= 60:
        minutes_together -= 60
        carried_hours = 1
    hours_together = int(time1[0:2]) + int(time2[0:2]) + carried_hours
    new_minutes = str(minutes_together)
    new_hours = str(hours_together)
    new_time = new_hours + ":" + new_minutes
    return new_time


#open rawData.xlsx
wb = load_workbook("rawData.xlsx")

#calculate sheet names
sheet_name_counter_limit = copy.deepcopy(number_of_days)
sheet_name_counter = 0
active_sheet_names = [report_day,]
current_loop_day = report_day

while sheet_name_counter < sheet_name_counter_limit - 1:
    current_loop_day = date_decrement(current_loop_day)
    active_sheet_names.append(current_loop_day)
    sheet_name_counter += 1



#begin mega line chart

#if this line is not here the whole program breaks, fyi.
active_sheet_names.reverse()

block_master_list = []
day_list = []
worksheet_number = 1

#aggregate all blocks in the dates in question
for worksheet_name in active_sheet_names:
    ws = wb[worksheet_name]
    for row in range(3, ws.max_row + 1):
        if ws.cell(row, 1).value is None:
            break
        new_block = [ws.cell(row,1).value, ws.cell(row,2).value,
                    ws.cell(row,3).value, worksheet_number]
        day_list.append(new_block)
    block_master_list.append(day_list)
    day_list = []
    worksheet_number += 1

#convert clock times to minutes on the 0-1439 scale, multiplied by the
#date number, and then remove the date number to be same format as
#chronologicalActivities from time_tracker

for day_block in block_master_list:
    #print(day_block)
    for activity_block in day_block:
        #print(activity_block)
        activity_block[0] = clock_time_to_minute_converter(activity_block[0]) #+ ((activity_block[3]*1440)-1440)
        activity_block[1] = clock_time_to_minute_converter(activity_block[1]) #+ ((activity_block[3]*1440)-1440)
        activity_block.pop(3)


#create first chart, the mega line chart
figure1, ax1 = pyplot.subplots()
ax1.plot()

#set chart dimensions
ax1.set_xlim(-25, 1465)

chart1_ylim = 0.25*number_of_days + 0.76
ax1.set_ylim(0,chart1_ylim)

#set the aspect ratio
ax1.set_aspect(aspect=300)

#make the axes invisible
ax1.axis('off')

#add labels
hourLabelCounter = 0
hourLabel = 0

hourLabelX = 0
timeLabelDict = {0:'12:00 AM', 1: "1:00 AM", 2: "2:00 AM", 3: "3:00 AM", 4: "4:00 AM",
                 5: '5:00 AM', 6: "6:00 AM", 7: "7:00 AM", 8: "8:00 AM", 9: "9:00 AM",
                 10: '10:00 AM', 11: "11:00 AM", 12: "12:00 PM", 13: "1:00 PM", 14: "2:00 PM",
                 15: '3:00 PM', 16: '4:00 PM', 17: '5:00 PM', 18: '6:00 PM', 19: '7:00 PM',
                 20: '8:00 PM', 21: '9:00 PM', 22: '10:00 PM', 23: '11:00 PM', 24:'12:00 AM'}

label_x_limit = 25

#set font scalars
if number_of_days > 10: 
    day_font_size_scalar = 0.10*(number_of_days-10)
    date_font_size_scalar = 0.10*(number_of_days-10)
else:
    day_font_size_scalar = 0
    date_font_size_scalar = 0

hour_height = 0.75 + 0.25*number_of_days
while hourLabelCounter < 25:
    ax1.add_patch(patches.Rectangle((hourLabelX, 0.25), 1.2, hour_height, facecolor = 'k', edgecolor = None))
    pyplot.text(hourLabelX, 0.19, s=timeLabelDict[hourLabelCounter],
                horizontalalignment='center', fontsize = 2.5-day_font_size_scalar)
    hourLabelX += 60
    hourLabelCounter += 1 


ax1.add_patch(patches.Rectangle((-25, 0.18), 1490, 0.0035, facecolor = 'k', edgecolor = None))    
ax1.add_patch(patches.Rectangle((-25, 0.22), 1490, 0.0035, facecolor = 'k', edgecolor = None))

#print(block_master_list)

label_day_number = 0
data_y_level = 0.25 * number_of_days + 0.260
#add data rectangles
for day_block in block_master_list:
    #print("day block")
    for final_activity_block in day_block:
        block_category = final_activity_block[2]
        block_color = colorCategorySorter(block_category)
        ax1.add_patch(patches.Rectangle((final_activity_block[0], data_y_level), int(final_activity_block[1]-final_activity_block[0]+1), 0.2,
                                        facecolor = block_color, edgecolor = None))

    #add date
    label_date = active_sheet_names[label_day_number]
    pyplot.text(hourLabelX, data_y_level+0.075, s="%s" %label_date, horizontalalignment='left', fontsize = 4-date_font_size_scalar)
    label_day_number +=1
    data_y_level -= 0.25
    #print(data_y_level)

#create legend

#add date
pyplot.text(1597, 0.07, s="Report Covering: %s - %s" %(active_sheet_names[0], active_sheet_names[number_of_days-1]), horizontalalignment='right', fontsize = 4-date_font_size_scalar)

#hour_label_day_number = 0
#hour_label_date = active_sheet_names[hour_label_day_number]
#pyplot.text(hourLabelX, 0.10, s="%s" %hour_label_date, horizontalalignment='center', fontsize = 4/number_of_days)


#create legend
pyplot.text(-25, 0.07, s="Legend: ", horizontalalignment='left', fontsize = 4-day_font_size_scalar)
categoryList = ["sleep", "maintenance", "career", "project", "fitness",
                 "outdoor", "entertainment", "relaxation", "socialization",
                 "grooming", "culinary", "housework", "education", "volunteering",
                 "caregiving", "transit", "missing"]

legendXCoord = 60
legendYCoord = 0.10
for legendProtoEntry in categoryList:
    legendColor = colorCategorySorter(legendProtoEntry)
    legendTitle = legendProtoEntry
    if legendTitle == "outdoor":
        legendTitle = "Outdoor"
    ax1.add_patch(patches.Rectangle((legendXCoord, legendYCoord), 15, 0.05,
                                    facecolor = legendColor, edgecolor = None))
    pyplot.text(legendXCoord+20, legendYCoord+0.005, s=legendTitle.title(),
                horizontalalignment='left', fontsize = 3.25-day_font_size_scalar)
    legendYCoord -= 0.07
    #if reached bottom of column, start next column
    if round(legendYCoord,2) == -0.04:
        legendXCoord += 125
        legendYCoord = 0.10

#save plot 1    
pyplot.savefig("day_report", dpi=800+(25*number_of_days), orientation = 'landscape', bbox_inches='tight')





#create plot 2

#create second plot
figure2, ax2 = pyplot.subplots()

wedgeWidth = 0.6

#create array containing the duration of all experience blocks
#print(chronologicalActivities)

flat_list =[]
for day_block in block_master_list:
    for event in day_block:
        flat_list.append(event)

#print(flat_list, "\n\n")

categorizedActivities = sorted(flat_list, key = lambda entry:entry[2])
#print(categorizedActivities)

#pull out category order
identifiedActivities = []
for entry in categorizedActivities:
    if entry[2] not in identifiedActivities:
        identifiedActivities.append(entry[2])

#create Activities list divorced from start and end time, only duration
durationActivities = []
for block in categorizedActivities:
    eventDuration = block[1]-block[0] + 1
    durationActivities.append([eventDuration, block[2]])
    
#create duration array
#separate durations by category

currentSet = []
durationProtoList = []
currentCategory = durationActivities[0][1]
durationActivitiesLoopCounter = 0

for pair in durationActivities:  
    if pair[1] == currentCategory:
        currentSet.append(pair[0])
    else:
        durationProtoList.append(currentSet)
        currentSet = []
        currentSet.append(pair[0])
    
    currentCategory = pair[1]
    durationActivitiesLoopCounter +=1
    if durationActivitiesLoopCounter == len(durationActivities):
        durationProtoList.append(currentSet)

#next step in creating array: find the largest number of entries in
#a single category

maxNumRepeats = 1
for listing in durationProtoList:
    if len(listing) > maxNumRepeats:
        maxNumRepeats = len(listing)

#following step: add junk 0's to categories with less entries
#(this makes the matrix 'regular' permitting matrix algebra
# and ensures matplotlib won't flip out)

durationArray = durationProtoList
for listing in durationArray:
    while True:
        if len(listing) < maxNumRepeats:
            listing.append(0)
        else:
            break

#sort in descending order by highest duration category
currentLocaterInList = 0
for nextListing in durationArray:
    nextListing.append(identifiedActivities[currentLocaterInList])
    currentLocaterInList += 1
    totalDuration = sum(nextListing[0:len(nextListing)-1])
    nextListing.append(totalDuration)

durationArray = sorted(durationArray, key = lambda entry:entry[-1], reverse = True)
#get durations for use in table
durations = []
for entry in durationArray:
    durations.append(entry[-1])

durationArrayCounter = 0
identifiedActivities = []

for listing in durationArray:
    identifiedActivities.append(listing[-2])
    listing = listing[:-2]
    durationArray[durationArrayCounter] = listing
    durationArrayCounter+=1


#print(durationArray)
durationArray = np.array(durationArray)

#set outer ring's colors, creating custom cmap
outerRingColors = []
for category in identifiedActivities:
    outerRingColors.append(colorCategorySorter(category))
outerRingColors = np.array(outerRingColors)

#set inner ring's colors, creating custom cmap
##innerRingColors = []
##for category2 in identifiedActivities:
##    currentColor = colorCategorySorter(category2)
##    tempCounter = 0
##    lightenVal = 0.6
##    while tempCounter < maxNumRepeats:
##        innerRingColors.append(innerColorCreator(currentColor, lightenVal))
##        if lightenVal >= 0.25:
##            lightenVal -= 0.2
##        tempCounter += 1
##
##innerRingColors = np.array(innerRingColors)


#create the labels

pieChartLabels = copy.deepcopy(identifiedActivities)
pieChartLabelCounter = 0
while pieChartLabelCounter < len(pieChartLabels):
    pieChartLabels[pieChartLabelCounter] = pieChartLabels[pieChartLabelCounter].capitalize()
    pieChartLabelCounter +=1
    
explode = []
for item in pieChartLabels:
    explode.append(0.2)
explode = tuple(explode)

#create pie chart outer ring
ax2.pie(durationArray.sum(axis=1), labels=None, radius=1,
        startangle = 90, colors=outerRingColors, autopct = None,
        pctdistance =0.85, explode = None,
        wedgeprops=dict(width=wedgeWidth, edgecolor='w'))
#create pie chart inner ring
#ax2.pie(durationArray.flatten(), radius=1-wedgeWidth, colors=innerRingColors,
#        startangle = 90,
#        wedgeprops=dict(width=wedgeWidth, edgecolor='w'))
#ax2.set_aspect("equal")

duration_array_copy = copy.deepcopy(durationArray)

#modify the total times into an array compatible with zip()
total_time_list = []
for value in duration_array_copy.sum(axis=1):
    new_value = hour_and_minute_converter(value)
    total_time_list.append(new_value)
total_time_array = np.array(total_time_list)

#modify the average times into an array compatible with zip()
average_time_list = []
for value  in duration_array_copy.sum(axis=1):
    average_daily_minutes = value/number_of_days
    new_value = hour_and_minute_converter(average_daily_minutes)
    average_time_list.append(new_value)
average_time_array = np.array(average_time_list)

#create equal length legend labels

#find maximum character length of a legend label
pie_chart_label_max_length = 1
for item in pieChartLabels:
    label_length = len(item)
    if label_length > pie_chart_label_max_length:
        pie_chart_label_max_length = label_length

#add spaces to all the other labels to make them all the same length
legend_category_labels = copy.deepcopy(pieChartLabels)
legend_category_labels_final = []
for category in legend_category_labels:
    category = category +","
    while len(category) < pie_chart_label_max_length + 2:
        category = category + " "
    legend_category_labels_final.append(category)

rcParams['font.family'] = 'monospace'
#rcParams['font.sans-serif'] = ['Tahoma']

ax2.legend(loc = 'upper left',
           labels =['%s %1.1f%%, %s, %s' %(l, s, t, u) for l,s,t,u in zip(legend_category_labels_final,(durationArray.sum(axis=1)/durationArray.sum()*100), total_time_array, average_time_array)],
           bbox_to_anchor=(-0.3, .8), bbox_transform=pyplot.gcf().transFigure)

pyplot.text(1.25, -1.25, s="Report Covering: %s - %s" %(active_sheet_names[0], active_sheet_names[number_of_days-1]), horizontalalignment='right', fontsize = 7.5)
pyplot.savefig("pie_report", dpi=500, orientation = 'landscape', bbox_inches='tight')
