#!python3 
# monthlyPlanHack.py - a simplie tool to generate xlsx monthly 
# plans for my stupid job for jerks

import openpyxl
import random

newsOutlet = ['BBC', 'Vice News', 'New York Times',
                'New Yorker', 
                'BBC World', 'CNN Online', 'San Francisco Chronicle']

podcastList = ['NPR', 'This American Life', 'Radiolab', 'In our time']

scienceOutlet = ['National Geographic', 'Motherboard', 'PLOS Science Journal',
                'BBC Science', 'BBC Science']         


youtubeChanel = ['Vice', 'Vox', 'Motherboard', 'Journeyman Pictures',
                'Science Friday']

authors = ['Tolkien', 'H.P. Lovecraft', 'Henry Thoreaux', '']



# columns is 18, with filler spaces 

# there are 6 rows, or  A-F

# open the template, that NEEDS to be in the same directory
wb = openpyxl.load_workbook('monthlyPlanTemplate.xlsx')
# open the sheet object
sheet = wb.get_sheet_by_name('Hoja1')

# input

# save as title

# toplines and titles
sheet['B1'] = input('Enter month date: ')
print('Month: ' + str(sheet['B1'].value))
sheet['D1'] = input('Enter company name: ')
print('Company name: ' + str(sheet['D1'].value))
sheet['F1'] = input('Enter group name: ')
print('Group name: ' + str(sheet['F1'].value))

# B 8-15 needs weekend recap + week discussion + 2 random
# B 14 needs oral exam + 1 random
# B 15 needs written quiz and 1 random

sheet['B8'] =  ' Weekend Recap / '  + 'Begin monthly Unit' + str(random.choice(podcastList)) # delete this!!!
sheet['B9'] =  ' Week Discussion /' + 'Listening Exercise and Discussion'
sheet['B10'] = ' Weekend Recap / '  + 'Continue Unit / Short Documentary / Short discussion'
sheet['B11'] = ' Week Discussion /' + 'Science article reading exercise and short Science Documentary'
sheet['B12'] = ' Weekend Recap / '  + 'News article and Political discussion/analysis'
sheet['B13'] = ' Week Discussion /' + 'Continue Unit / Science or Technology Documentary and Discussion'
sheet['B14'] = ' Weekend Recap / '  + 'Quick oral Exam / news article analysis '
sheet['B15'] = ' Week Discussion /' + 'Unit activity and Written Quiz'


# C 8-15 needs weekend recap + week discussion + 2 verb desc random
# C 14 needs oral exam + 1 verb desc random
# C 15 needs written quiz and 1 verb desc random

sheet['C8'] =  ''
sheet['C9'] =  ''
sheet['C10'] = ''
sheet['C11'] = ''
sheet['C12'] = ''
sheet['C13'] = ''
sheet['C14'] = ''
sheet['C15'] = ''

# D 8-15 needs random 'pg number' or media

sheet['D8'] =  'pg. XXX'
sheet['D9'] =  
sheet['D10'] = 'pg. XXX'
sheet['D11'] = 
sheet['D12'] = 
sheet['D13'] = 'pg. XXX'
sheet['D14'] = 
sheet['D15'] = 'pg. XXX'

# E 8-15 needs a 'put into practice'

sheet['E8'] =  ''
sheet['E9'] =  ''
sheet['E10'] = ''
sheet['E11'] = ''
sheet['E12'] = ''
sheet['E13'] = ''
sheet['E14'] = ''
sheet['E15'] = ''

# F 8-15 needs corresponding sources

sheet['F8'] =  ''
sheet['F9'] =  ''
sheet['F10'] = ''
sheet['F11'] = ''
sheet['F12'] = ''
sheet['F13'] = ''
sheet['F14'] = ''
sheet['F15'] = ''


# write the files out
planName = input('Name the file.. i.e. Prudential May 2016.. :')

print('file saved as: ' + planName)

wb.save(planName + '.xlsx')





