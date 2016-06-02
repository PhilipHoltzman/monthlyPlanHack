#!python3 
# monthlyPlanHack.py - a simple tool to generate xlsx monthly 
# plans for my stupid job for jerks


# TODO : 


# generate random links for column F for random sources

# Weekly random day generators, further obfuscating my generated lesson plans

# randomize key areas to further obfuscate

# figure out how to draw those border/ underlines ->  http://openpyxl.readthedocs.io/en/default/styles.html#introduction





import openpyxl
from openpyxl.drawing.image import Image
from openpyxl.styles import PatternFill, Border, Side, Alignment, Protection, Font
import random



# media sources
newsOutlet = ['BBC', 'Vice News', 'New York Times',
                'New Yorker', 'Iberian Times',
                'BBC World', 'CNN Online', 'San Francisco Chronicle', 'Reddit World News']

podcastList = ['NPR', 'This American Life', 'Radiolab', 'In our time', 'Sword and Scale', 'Science Friday']

scienceOutlet = ['National Geographic', 'Motherboard', 'PLOS Science Journal',
                'BBC Science', 'BBC Science' 'Reddit Science']         

youtubeChanel = ['Vice', 'Vox', 'Motherboard', 'Journeyman Pictures',
                'Science Friday', 'NYT Channel', 'Broadly', 'Debate Squared']

authors = ['Tolkien', 'H.P. Lovecraft', 'Henry Thoreaux', 'Edgar Allen Poe', 'George RR Martin', 'Isaac Asimov']

# topics

scienceTopic = ['Genetic Engineering',
				'Climate Change',
				'Space Travel',
				'Disease',
				'Robotics',
				'Marine Biology',
				'Chemistry',
				'Pollution',
				'Medical Technology']

politicalTopic =   ['Presidential Election',
					'Border Control',
					'Drug Policy',
					'Mexican Policy',
					'Middle East Conflict',
					'Refugee Situation in Europe',
					'Racism and Hate groups',
					'Womens Rights',
					'Political Corruption']

technologyTopic =  ['Cybersecurity',
					'Digital Privacy',
					'Social Media',
					'Emergent Technologies',
					'Encryption',
					'Establishing an online presence',
					'Protecting Children online',
					'History of the Internet',
					'Online resources for teaching yourself']

themeTopic = ['','','','','','','','','']

lit = ['short story', 'poem']

# temporary choices

tNewsOutlet = str(random.choice(newsOutlet))
tNewsOutlet2 = str(random.choice(newsOutlet))
tPodcastList = str(random.choice(podcastList))
tScienceOutlet = str(random.choice(scienceOutlet))
tScienceOutlet2 = str(random.choice(scienceOutlet))
tYoutubeChanel = str(random.choice(youtubeChanel))
tYoutubeChanel2 = str(random.choice(youtubeChanel))
tAuthors = str(random.choice(authors))

tScienceTopic = str(random.choice(scienceTopic))
tPoliticalTopic = str(random.choice(politicalTopic))
tThemeTopic = str(random.choice(themeTopic))
tTechnologyTopic = str(random.choice(technologyTopic))

tLit = str(random.choice(lit))

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
sheet['B13'] = ' Week Discussion /' + 'Continue Unit / Technology Documentary and Discussion'
sheet['B14'] = ' Weekend Recap / '  + 'Quick oral Exam / news article analysis '
sheet['B15'] = ' Week Discussion /' + 'Unit activity and Written Quiz'


# C 8-15 needs weekend recap + week discussion + 2 verb desc random
# C 14 needs oral exam + 1 verb desc random
# C 15 needs written quiz and 1 verb desc random
# Exposure/Grammar

sheet['C8'] =  'Begin new coursework, with a focus on exploring key terms'
sheet['C9'] =  'Listen to a podcast, with pauses inbetween, debating the topics in between'
sheet['C10'] = 'Work a bit on the current unit, watch a short documentary, and brief discussion'
sheet['C11'] = 'Analyze a Science article, and understand the difference between reports and papers and enjoy a science documentary'
sheet['C12'] = 'Read a news article, and have an objective news and political analysis'
sheet['C13'] = 'Work on Unit, and then watch a Technology Documentary, followed by a discussion'
sheet['C14'] = 'Conduct a quick Oral Exam, followed by an analysys of a short story or poem'
sheet['C15'] = 'Finish the weeks Unit Activity and conduct the monthly written quiz'

# D 8-15 needs random 'pg number' or media

sheet['D8'] =  'pg. XXX'
sheet['D9'] =   tPodcastList + ' podcast'
sheet['D10'] = 'pg. XXX'
sheet['D11'] =  tYoutubeChanel + ' science short documentary'
sheet['D12'] =  tNewsOutlet + ' news article'
sheet['D13'] = 'pg. XXX'
sheet['D14'] = tNewsOutlet2 + ' news article analysis'
sheet['D15'] = 'pg. XXX'

# E 8-15 needs a 'put into practice'

sheet['E8'] =  'Students will read and discuss passages in the book, ask questions, and review key business terms' # randomize these in future
sheet['E9'] =  'Listen to the ' + tPodcastList + ' on ' + tScienceTopic   # randomize these tTopics
sheet['E10'] = 'Writing exercise on current Unit and watch a documentary on ' + tPoliticalTopic 
sheet['E11'] = 'Review an article from ' + tScienceOutlet + ' about ' + tScienceTopic
sheet['E12'] = 'Read an article from ' + tNewsOutlet + ' regarding ' + tThemeTopic 
sheet['E13'] = 'After working on Unit, watch a quick ' + tYoutubeChanel + ' on ' + tTechnologyTopic
sheet['E14'] = 'Conduct a short Oral Exam and then analyze a ' + tLit + ' by ' + tAuthors 
sheet['E15'] = 'Complete the Monthly activity and review the months work, and short written quiz' # randomize these in future

# F 8-15 needs corresponding sources

# sources variables

ytLinksArray = ['https://youtu.be/yQjove0nzss',
				'https://youtu.be/vw4ybw4yv45',
				'https://youtu.be/wvtywtwrtyy',
				'https://youtu.be/3lkl4mmmgo6',
				'https://youtu.be/zxcb79bssbn',
				'https://youtu.be/mlen3769b77',
				'https://youtu.be/bbmvi3ifuv5',
				'https://youtu.be/nuwhj6vls80',
				'https://youtu.be/wnkvhkjbb2k',
				'https://youtu.be/jss0qqv6wwj',
				'https://youtu.be/nbiuyd5vlk0',
				'https://youtu.be/dbiegitfv4s',
				'https://youtu.be/fbiuguie3vv',
				'https://youtu.be/oiybbb8vk0s',]

ytLink1 = str(random.choice(ytLinksArray))
ytLink2 = str(random.choice(ytLinksArray))

sheet['F8'] =  'Advanced Business English'
sheet['F9'] =  tPodcastList + ' Podcast'
sheet['F10'] = ytLink1
sheet['F11'] = tScienceOutlet + ' article'
sheet['F12'] = tNewsOutlet + ' article'
sheet['F13'] = ytLink2
sheet['F14'] = tAuthors + ' poem'
sheet['F15'] = 'Advanced Business English and Teach English Feel Good worksheet'

# outlines 



# insert ahpla logo

logo = Image('ahplalogo.png')
sheet.add_image(logo,'E27')



# write the files out
planName = input('Name the file.. i.e. Prudential May 2016.. :')

print('file saved as: ' + planName)

wb.save(planName + '.xlsx')





