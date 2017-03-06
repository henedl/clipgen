import gspread
import os, sys
import subprocess
from datetime import datetime
from oauth2client.service_account import ServiceAccountCredentials

# Constants 
REENCODING = False
FILEFORMAT = '.mp4'
VERSIONNUM = '0.2.6'
DEBUGGING  = False

# What is this?
# This script will help quickly cut out video snippets in user reserach videos, based on researcher's timestamps in a spreadsheet!

# TODO
# Quality of life:
# 	- Allow for comma separation and not just + for multiple timestamps in the same cell
# 	- Settings, at least the ability to change fileformat, reencoding and other overarching options, without editing code
#	- Project names with spaces: attempt to resolve files as both "xystonPike" and "xyston_pike"?
#	- Autocompletion for name list stuff?
#	- There are many other types of name list interactions possible, perhaps generate a numerically sorted list and allow selection by just giving a number?
#	- Case insensitivity for sheet name entry
#	- Stop requiring spaces between multiple clips (a + should be enough)
#	- Sanity checking video length (videos longer than X seconds are not likely, so ping the user before cutting - or in post cut presentation?)
# Programming stuff:
# 	- Probably break out input-parsing to a separate method (just pass it a list of what to accept, what to not accept and error messages?)
# 	- It would be much faster to just dump all the contents of the sheet in to a list and work with that (also supported by gspread)
# 	- We should check for file name length, max 255 (including suffix)
# 	- Cleaner variable names (even for things relating to iterators - ipairList, really?)
#	- Also, naming consistency, currently there is some camel case and some underscores, etc
#	- Make sure that the script does its work on the sheet named "data set" (+ support for selecting other tabs)
#	- Except ffmpeg crashes/errors
#	- Except connection timeouts (probably except gspread.exceptions.HTTPError on line 29?)
#	- Logging of which timestamps are discarded
#	- Gracefully handle connection timeouts (look into gspread code)
#	- Debug mode (with multiple levels?) throughout the code
#	- Upgrade to Python 3
#	- gspread should reauthorize (either after an extended time when it fails out, or whenever a new loop begins)
#	- Refactor try statements to be smaller
# Batch improvements:
# 	- Implement the special character to select only one video to be rendered, out of several
# 	- Add support for special tokens like * for starred video clip (this can be added to the dict as 'starred' and then read in the main loop)
# 	- Start using the meta field for checking which issues are already processesed and what the grouping is
#	- Generate all positive moments as a batch call
#	- Generate all issues of a certain category as a batch call?
# Major new features:
# 	- GUI
#	- Cropping and timelapsing! For example generate a timelapse of the minimap in TWY or EU.

# Goes through sheet, bundles values from timestamp columns and descriptions columns into tuples.
def generate_list(sheet, mode, type='Default'):
	numRow = sheet.row_count
	numCol = sheet.col_count
	p = sheet.find('Participants') # Find pariticpant listing coords.
	m = sheet.find('Meta') # Find the meta tag coords.
	s = sheet.find('Summary')
	times = []

	# TODO 
	# Add more processing of the title, split out the study number, and project name.
	studyName = sheet.cell(1, 1).value # Find the title of the study, assuming top left in sheet.
	studyName = studyName[0:studyName.find('Data set')-1] # Cut off the stuff we don't want.
	print 'Beginning work on {0}.'.format(studyName)
	
	# Just some name formatting, after we announced everything up top.
	studyName = studyName.lower() 
	studyName = studyName.replace('study ', 'study')
	studyName = studyName.replace(' ', '_') # Replace whitespace with underscore.
	# It should now look like this: 'thundercats_study5'

	# Figure out how many users we have in the sheet (assumes every user is indicated by a 'PXX' identifier)
	numUsers = 0
	userList = sheet.row_values(p.row+1)
	for j in range(0, numCol-p.col):
		if len(userList[j]) > 0:
			if userList[j][0] == 'P':
				numUsers += 1
	print 'Found {0} users in total, spanning columns {1} to {2}.'.format(numUsers, p.col, numUsers+p.col)

	if mode == 'batch':
		latestCategory = ''
		if type == 'Default':
			for j in range(p.col, p.col + numUsers):	
				for i in range(p.row + 2, numRow - p.col):
					if sheet.cell(i, m.col).value == 'T':
						latestCategory = sheet.cell(i, s.col).value
						print '+ Found category \'{0}\' on line {1}.'.format(latestCategory, i)
					val = sheet.cell(i, j)
					if val.value is None:
						# Discard empty cells.
						pass
					elif val.value == '':
						# Discard empty cells.
						pass
					else:
						issue = { 'cell': val, 'desc': sheet.cell(i, s.col).value, 'study': studyName, 'participant': sheet.cell(j, i).value, 'category': latestCategory }
						times.append(issue)
						print '+ Found timestamp: {0}'.format(val.value)
		elif type == 'Positive':
			pass
	elif mode == 'line':
		# This mode generates videos for a single line/row number.
		latestCategory = ''
		while True:
			try:
				lineSelect = int(raw_input('\nWhich issue (row number only)?\n>> '))
			except ValueError:
				# TODO
				# This should not be set up this way, make it loop
				lineSelect = int(raw_input('\nTry again. Integer only.\n>> '))

			print 'Issue titled: {0}'.format(sheet.cell(lineSelect, s.col).value)
			yn = raw_input('Is this the correct issue? y/n\n>> ')
			if yn == 'y':
				break
			else:
				pass

		# Go up until we find the category we are headed under.
		while latestCategory == '':
			try:
				for i in range(lineSelect, p.row, -1):
					if sheet.cell(i, m.col).value == 'T':
						latestCategory = sheet.cell(i, s.col).value
						print '+ Found category \'{0}\' on line {1}.'.format(latestCategory, i)
						break # Exit the for loop so we don't keep going up.
			except IndexError:
				break

		for j in range(p.col, p.col + numUsers):
			val = sheet.cell(lineSelect, j)
			if val.value is None:
				# Discard empty cells.
				pass
			elif val.value == '':
				# Discard empty cells.
				pass
			else:
				issue = { 'cell': val, 'desc': sheet.cell(lineSelect, s.col).value, 'study': studyName, 'participant': sheet.cell(j, i).value, 'category': latestCategory }
				times.append(issue)
				print '+ Found timestamp: {0}'.format(val.value.replace('\n',' '))
	elif mode == 'range':
		# This mode generates videos for all issues found in a range or span of issues.
		latestCategory = ''
		while True:
			try:
				startLineSelect = int(raw_input('\nWhich starting line (row number only)?\n>> '))
				endLineSelect = int(raw_input('\nWhich ending line (row number only)?\n>> '))
			except ValueError:
				lineSelect = int(raw_input('\nTry again. Integer only.\n>> '))
			print 'Lines selected: {0}'.format(sheet.cell(lineSelect, s.col).value)
			yn = raw_input('Is this correct? y/n\n>> ')
			if yn == 'y':
				break
			else:
				pass
		for j in range(p.col, p.col + numUsers):
			for i in range(startLineSelect, endLineSelect):
				if sheet.cell(i, m.col).value == 'T':
					latestCategory = sheet.cell(i, s.col).value
					print '+ Found category \'{0}\' on line {1}.'.format(latestCategory, i)
				val = sheet.cell(i, j)
				if val.value is None:
					# Discard empty cells.
					pass
				elif val.value == '':
					# Discard empty cells.
					pass
				else:
					issue = { 'cell': val, 'desc': sheet.cell(i, s.col).value, 'study': studyName, 'participant': sheet.cell(j, i).value, 'category': latestCategory }
					times.append(issue)
					print '+ Found timestamp: {0}'.format(val.value)
	elif mode == 'select':
		# This mode generates a list of non-completed issues and lets user select from those.
		pass

	return times

# Takes a string, returns a double digit number
def double_digits(number):
	try:
		if int(number) < 10:
			return '0' + number
		else:
			return number
	except TypeError:
		# If we can't typecast, we give up
		return number

def filesize(size, precision=2):
    suffixes = ['B','KB','MB','GB','TB']
    suffixIndex = 0
    while size > 1024 and suffixIndex < 4:
        suffixIndex += 1 
        size = size / 1024.0
    return '%.*f%s'%(precision, size, suffixes[suffixIndex])

# Appends an incremeneted number to the end of files that already exist.
def check_filename(filename):
	step = 1
	while True:
		if os.path.isfile(filename):
			if step < 2:
				suffixPos = filename.find(FILEFORMAT)
				filename = filename[0:suffixPos] + '-' + str(step) + FILEFORMAT
			else:
				dashPos = filename.rfind('-')
				filename = filename[0:dashPos] + '-' + str(step) + FILEFORMAT
			step += 1
		else:
			break

	return filename

def clean_issue(issue):
	timeStamps = []
	unparsedTimes = issue['cell'].value.lower().split()
	
	# Using own iterator here, instead of letting the for-loop set this up. Otherwise we can't manually advance the iterator (we need to step twice
	# which continue won't do.)
	lines = iter(range(0,len(unparsedTimes)))
	for i in lines:
		unparsedTimes[i] = unparsedTimes[i].strip()
		if unparsedTimes[i] == '':
			pass
		elif unparsedTimes[i].find('interview') != -1:
			issue['interview'] = True
			# The reason we use i+1 everywhere in this block is because of us doing the advancing at the end. Should probably still work if we moved the next() up top here.
			if unparsedTimes[i+1].find('-') >= 0:
				if unparsedTimes[i+1][unparsedTimes[i+1].find('-')-1].isdigit():
					timePair = unparsedTimes[i+1][0:unparsedTimes[i+1].find('-')], unparsedTimes[i+1][unparsedTimes[i+1].find('-')+1:]
					timeStamps.append(timePair)
			elif unparsedTimes[i+1].find(':') >= 0: 
				if unparsedTimes[i+1][unparsedTimes[i+1].find(':')-1].isdigit():
					timePair = unparsedTimes[i+1], '00:00:00' # We add the zero time so that we will later fire the add_duration for this timestamp
					timeStamps.append(timePair)
			next(lines, None)
			continue
		elif unparsedTimes[i].find('-') >= 0:
			if unparsedTimes[i][unparsedTimes[i].find('-')-1].isdigit():
				# Slice the timestamp until the dash, and then from after the dash.
				timePair = unparsedTimes[i][0:unparsedTimes[i].find('-')], unparsedTimes[i][unparsedTimes[i].find('-')+1:]
				timeStamps.append(timePair)
		elif unparsedTimes[i].find(':') >= 0:
			if unparsedTimes[i][unparsedTimes[i].find(':')-1].isdigit():
				timePair = unparsedTimes[i], '00:00:00' # We add the zero time so that we will later fire the add_duration for this timestamp
				timeStamps.append(timePair)
		else:
			pass

	issue['times'] = timeStamps

	# Are there other characters that will mess up file names? If so, add them here.
	# TODO: This should be reasonable to do with a dictionary/list loop instead of multiple replaces
	issue['desc'] = issue['desc'][ issue['desc'].rfind(']')+1: ].strip()
	issue['desc'] = issue['desc'].replace('/','-')
	issue['desc'] = issue['desc'].replace('?','_')
	issue['desc'] = issue['desc'].replace('\\','-')
	issue['desc'] = issue['desc'].replace('\"','')
	issue['desc'] = issue['desc'].replace('\'','')
	issue['desc'] = issue['desc'].replace('.','')
	issue['desc'] = issue['desc'].replace('<','')
	issue['desc'] = issue['desc'].replace('>','')
	issue['desc'] = issue['desc'].replace('|','')

	return issue

def ffmpeg(inputfile, outputfile, startpos, outpos, reencode):
	# TODO
	# Protect against negative duration
	# Protect against videos that have an outtime beyond base video length
	# Protect against unreasonably long clips

	# DEBUG
	# Just makes the clip a minute long if we didn't get an in-time
	if outpos == '00:00:00':
		outpos = add_duration(startpos)

	duration = get_duration(startpos, outpos)

	if duration < 0:
		print 'Can\'t work with negative duration for videos, exiting.'
		sys.exit(0)

	print 'Cutting {0} from {1} to {2}.'.format(inputfile, startpos, outpos)
	if not reencode:
		subprocess.call(['ffmpeg', '-y', '-loglevel', '16', '-ss', startpos, '-i', inputfile, '-t', str(duration), '-c', 'copy', '-avoid_negative_ts', '1', outputfile])
	else:
		# If we do this, we will re-encode the video, but resolve all issues with with iframes early and late.
		subprocess.call(['ffmpeg', '-y', '-loglevel', '16', '-ss', startpos, '-i', inputfile, '-t', str(duration), outputfile])
	print '+ Generated video \'{0}\' successfully.\n File size: {1}\n Expected duration: {2} s\n'.format(outputfile,filesize(os.path.getsize(outputfile)), duration)

# Returns the duration of a clip as seconds
def get_duration(intime, outtime):
	duration = 0
	try:
		intimeDatetime = datetime.strptime(intime,'%H:%M:%S')
		outtimeDatetime = datetime.strptime(outtime,'%H:%M:%S')
	except ValueError as e:
		print '* Timestamp formatting error was caught.'
		print e
		try:
			intimeDatetime = datetime.strptime(intime,'%H:%M:%S.%f')
			outtimeDatetime = datetime.strptime(outtime,'%H:%M:%S.%f')
		except ValueError as e:
			print '* Further timestamp formatting error was caught, exiting.'
			print '* Timestamp formats need to match each other.'
			print e
			sys.exit(0)

	hDelta = (outtimeDatetime.hour - intimeDatetime.hour)*60*60
	mDelta = (outtimeDatetime.minute - intimeDatetime.minute)*60
	sDelta = (outtimeDatetime.second - intimeDatetime.second)
	duration = hDelta + mDelta + sDelta

	return duration

# Just adds a minute
def add_duration(intime):
	intimeDatetime = datetime.strptime(intime,'%H:%M:%S')
	if intimeDatetime.minute == 59:
		return double_digits(str(intimeDatetime.hour+1)) + ':00:' + double_digits(str(intimeDatetime.second))
	else:	
		return double_digits(str(intimeDatetime.hour)) + ':' + double_digits(str(intimeDatetime.minute+1)) + ':' + double_digits(str(intimeDatetime.second))

# Comma-separated list of all accessible Google Spreadsheets
def get_alldocs(connection):
	docs = []
	for doc in connection.openall():
		docs.append(doc.title)
	return ', '.join(docs)

def main():
	# Change working directory to place of python script.
	os.chdir(os.path.dirname(os.path.abspath(__file__)))
	print '-------------------------------------------------------------------------------'
	print 'Welcome to clipgen v{1}, for use by Paradox User Research\n\nWorking directory: {0}\nPlace video files and the oauth.json file in this directory.'.format(os.getcwd(), VERSIONNUM)
	
	# Remember that documents need to be shared to the email found in the json-file for OAuth-ing to work.
	# Each user of this program should also have their own, unique json-file (generate this on the Google Developer API website).
	scope = ['https://spreadsheets.google.com/feeds']
	try:
		credentials = ServiceAccountCredentials.from_json_keyfile_name('oauth.json', scope)
	except IOError as e:
		print e
		print 'Could not find credentials (oauth.json).'
		# TODO
		# Here we could have an interactive method that asks the user for the right directory to work in. Same for video files (would require some new code)
		sys.exit(0)
	try:
		gc = gspread.authorize(credentials)
	except gspread.AuthenticationError as e:
		print e
		print 'Could not authenticate.'
		sys.exit(0)
	
	inputFileFails = 0

	while True:
		inputName = raw_input('\nPlease enter the index, name, URL or key of the spreadsheet (\'all\' for list,    \'new\' for list of newest, \'last\' to immediately open latest):\n>> ')
		try:
			if inputName[:4] == 'http':
				worksheet = gc.open_by_url(inputName).sheet1
				break
			elif inputName[:3] == 'all':
				docList = get_alldocs(gc).split(',')
				print '\nAvailable documents:'
				for i in range(len(docList)):
					print '{0}. {1}'.format(i+1, docList[i].strip())
			elif inputName[:3] == 'new':
				docList = get_alldocs(gc).split(',')
				print '\nNewest documents:'
				for i in range(3):
					print '{0}. {1}'.format(i+1, docList[i].strip())
			elif inputName[:4] == 'last':
				latest = get_alldocs(gc).split(',')[0]
				worksheet = gc.open(latest)
  				break
  			elif inputName[0].isdigit():
  				i = int(inputName)-1
  				worksheet = gc.open(get_alldocs(gc).split(',')[i].strip())
  				break
			elif inputName.find(' ') == -1:
				worksheet = gc.open_by_key(inputName).sheet1
				break
			else:
				worksheet = gc.open(inputName).sheet1
				break
		except gspread.SpreadsheetNotFound:
			inputFileFails += 1
			if inputFileFails <= 1 or inputFileFails >= 3:
				print '\nDid not find spreadsheet. Please try again.'
			else:
				print '\n###############################################################################'
				print 'Remember that you need to share the spreadsheet you want to parse. Share it with the user listed in the json-file (value of client_email).'
				print '\nThis needs to be done on a per-document basis.'
				print '\nAvailable documents: {0}'.format(get_alldocs(gc))
				print '###############################################################################\n'

	print 'Connected to Google Drive! Using Sheet: {0}'.format(worksheet.title)
	inputModeFails = 0

	while True:
		while True:
			inputMode = raw_input('\nSelect mode: (b)atch, (r)ange or (l)ine\n>> ')
			try:
				if inputMode[0] == 'b' or inputMode == 'batch':
					timesList = generate_list(worksheet, 'batch')
					break
				elif inputMode[0] == 'l' or inputMode == 'line':
					timesList = generate_list(worksheet, 'line')
					break
				elif inputMode[0] == 'r' or inputMode == 'range':
					timesList = generate_list(worksheet, 'range')
					break
				#elif inputMode[0] == 's' or inputMode == 'select':
				#	timesList = generate_list(worksheet, 'select')
				#	break
				elif inputMode == 'karl':
					plogo()
			except IndexError as e:
				inputModeFails += 1

		print '\n* ffmpeg is set to never prompt for input and will always overwrite.\n  Only warns if close to crashing.\n'
		videosGenerated = 0

		try:
			ipairPos = worksheet.find('Interview pairs')
			ipairList = worksheet.cell(ipairPos.row+1, ipairPos.col).value.split(',')
		except gspread.exceptions.CellNotFound:
			pass
		
		for i in range(0, len(timesList)):
			# timesList is a list containing issues, one per index
			# issues are dicts that hold:
			# - cell 			Full gspread Cell object (row, col, value)
			# - desc 			String, summary description of the issue
			# - study 			String, name of the study
			# - participant 	String, participant ID (without prefix)
			# - times 			List, contains one timestamp pair (as a tuple) per index
			# - interview 		
			# Note that the 'times' entry in the dict is generated during the clean_issue method call.

			timesList[i] = clean_issue(timesList[i])
			for j in range(0,len(timesList[i]['times'])):
				vidIn, vidOut = timesList[i]['times'][j]
				vidName = check_filename('[Study ' + filter(str.isdigit, timesList[i]['study']) + '][' + timesList[i]['category'] + '] ' + timesList[i]['desc'] + FILEFORMAT)
				if 'interview' in timesList[i]:
					#print '{0},{1},{2}'.format(i,j,timesList[i]['interview'])
					for k in range(0,len(ipairList)):
						if ipairList[k].find(timesList[i]['participant']) >= 0:
							#print 'Participant match {0},{1}'.format(ipairList[k],timesList[i]['participant'])
							ipairList[k] = ipairList[k].replace('+', 'p')
							ipair = ipairList[k]
							#print ipair
						#print '{0},{1},{2}'.format(ipairList[k],timesList[i]['participant'],ipairList[k].find(timesList[i]['participant']))
					if timesList[i]['interview']:
						baseVideo = timesList[i]['study'] + '_interview_p' + ipair  + FILEFORMAT
					else:
						baseVideo = timesList[i]['study'] + '_p' + timesList[i]['participant']  + FILEFORMAT
				else:
					baseVideo = timesList[i]['study'] + '_p' + timesList[i]['participant']  + FILEFORMAT
				ffmpeg(inputfile=baseVideo, outputfile=vidName, startpos=vidIn, outpos=vidOut, reencode=REENCODING)
				videosGenerated += 1

		if not REENCODING:
			print '* No re-encoding done, expect:\n- inaccurate start and end timings\n- lossy frames until first keyframe\n- bad timecodes at the end\n'
		else:
			pass
		print 'All done, created {0} videos!\nFiles are in {1}\n'.format(videosGenerated, os.getcwd())
		yn = raw_input('Continue working (y) or quit the program (n)? y/n\n>> ')
		if yn == 'n':
			break
		else:
			pass

def plogo():
	print '                                          ;.\n	                                  ###:   ######@.\n	                                  ;###   #########\n 	                          .###;    ###   #########@\n	                          +####\'   ###;  ##########\n	                           #####   @##@  ##########    .\n 	                     \'      #####  @###,\'#########@    #@\'\n	                    +##+     ##### ###############@   ######\n	                    +###;    \'#####################  #######\n 	                     ####     ##############################@\n	                      ####   \'###############################\'\n	                       ####.,################################@\n 	                ;#,    :#####################################\n	               @####;  :################,    .@############@\n	               #######@###############+         ###########\n 	                ;#####################           ##########    ,\n	                 :###################@           ###########+@###;\n	                 \'####################           +################\n 	                 \'########, ##########:          #################\'\n	                 #########  @##########      @#########@ #########:\n	                :#########  ;#########;     @##########  .####+.,\';\n 	                @#########+ ##########     \'##########+   @###\n	                \'######@#############      ###########    .###\'\n	                 ###@    \'#####@\'+#;      .###########     ####\'\n 	                :###       ###@            ###########     ######@\n	          ,##@::###@       @##+            @##########@\'   ########\n	          @#########,      ####\'           \'#######################\n 	          ###########;     #####@           #############+@#######@\n	          .###########     +###@##@         #@\'########+    ######,\n	           ###########       ## \'###        #\' #\'@#####     ,#####\n 	           ####\'@###,        ##  @##;      :#, #+ ####:      #####\n	           ###   ;##         ##   ##.      @#\' #@ ,###       @#####\n	           ##+    ##\' :     \'###  ##       ##  ##  @##       @#####@\n 	          .##@    #####.     +######,     \'##  @#   ##       @#####\'\n	           ###,  ######+      @###@##     +##  ,#.  :@       @#####\n	           @###########\'      @#@#\'.#     ###   #@           #####@\n 	           ,###########,      @\' #@       ###  ,##           #####\'\n	           .###@  #####       ,\' ,#       ####,###\'         +#####\'\n	            ##@   #####                   #########         ######\n 	            +\'   @####@                @######+   :        @######\'\n	               @######@              \'######@            ;#######;\n	             \'########@             @######             #########\n 	             \'#########            .@,#####             ########,\n	              @########.             :####.             #######@\n	               @#######@             @####              ######+\n 	                :#######             #. @@             ,###:\n	                  ######\'            \'  :@             ####,\n	                   #####\'                @            #####+\n 	                   #####\'                           \'######\'\n	                    ####:                         \'#######\' \n	                    ####                         #######@\n 	                    ####                        ,######:\n	                    ###:                     ,\',#####@\n	                   \'##@                    ,########:\n	                   @#@                   \'######  \' \n 	                                       \'@######\n	                                        @##@@.'
	print '\n	 /$$$$$$$                                    /$$\n 	| $$__  $$                                  | $$                    \n 	| $$  \ $$ /$$$$$$   /$$$$$$  /$$$$$$   /$$$$$$$  /$$$$$$  /$$   /$$\n 	| $$$$$$$/|____  $$ /$$__  $$|____  $$ /$$__  $$ /$$__  $$|  $$ /$$/\n 	| $$____/  /$$$$$$$| $$  \__/ /$$$$$$$| $$  | $$| $$  \ $$ \  $$$$/ \n 	| $$      /$$__  $$| $$      /$$__  $$| $$  | $$| $$  | $$  >$$  $$ \n 	| $$     |  $$$$$$$| $$     |  $$$$$$$|  $$$$$$$|  $$$$$$/ /$$/\  $$\n 	|__/      \_______/|__/      \_______/ \_______/ \______/ |__/  \__/'

if __name__ == '__main__':
    try:
    	main()
    except KeyboardInterrupt:
    	print '\nInterrupted by user'
    	try:
    		sys.exit(0)
    	except SystemExit:
    		os._exit(0)