import gspread
import os, sys, string
import subprocess
from datetime import datetime
from oauth2client.service_account import ServiceAccountCredentials

# Constants 
REENCODING = False
FILEFORMAT = '.mp4'
VERSIONNUM = '0.2.7'
SHEET_NAME = 'data set'
DEBUGGING  = False

# What is this?
# This script will help quickly cut out video snippets in user reserach videos, based on researcher's timestamps in a spreadsheet!

# TODO
# Quality of life:
# 	- Settings, at least the ability to change fileformat, reencoding and other overarching options, without editing code
#	- Timestamp cleaning can't handle this situation: "H:M:S-H:M:S+interview H:M:S" because of spaces between interview and subsequent timestamp, but not before.
#	- Timestamp cleaning doesn't handle: " +H:M:S" either, strip + prefixes?
#	- Timestamp cleaning can't handle: "H:M:S + interview"
#	- Command to open the current Sheet in Chrome from the commandline?
#	- Created composite videos with clips from multiple participants?
#	- Title/ending cards?
#	- Being able to select multiple non-continous lines
#	- Add ability to target only one cell. Proposed syntax "P01.11". Should also be batchable, i.e. "P01.11 + P03.11 + P03.09". Should be available directly at (current) mode select stage.
# Programming stuff:
#	- Command line arguments to run everything from a prompt instead of interactively.
# 	- It would be much faster to just dump all the contents of the sheet into a list and work with that (easily supported by gspread)
# 	- Cleaner variable names (even for things relating to iterators - ipairList, really?)
#	- Also, naming consistency, currently there is some camel case and some underscores, etc
#	- Logging of which timestamps are discarded
#	- Expand debug mode (with multiple levels?)
#	- Upgrade to Python 3?
#	- Refactor try statements to be smaller
#	- Support other data formats (Excel, CSV) - would need to re-write parsing backend and refactor code heavily
# Batch improvements:
# 	- Implement the special character to select only one video to be rendered, out of several
# 	- Add support for special tokens like * for starred video clip (this can be added to the dict as 'starred' and then read in the main loop)
# 	- Start using the meta fields for checking which issues are already processesed and what the grouping is
# Major new features:
# 	- GUI
#	- Cropping and timelapsing! For example generate a timelapse of the minimap in TWY or EU.

# Goes through sheet, bundles values from timestamp columns and descriptions columns into tuples.
def generate_list(sheet, mode, type='Default'):
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
	studyName = studyName[0:studyName.find('study')].replace(' ', '') + '_' + studyName[studyName.find('study'):]
	studyName = studyName.replace(' ', '_') # Replace any leftover whitespace with underscore.
	# It should now look like this: 'thundercats_study5'

	# Figure out how many users we have in the sheet (assumes every user is indicated by a 'PXX' identifier)
	numUsers = 0
	userList = sheet.row_values(p.row+1)
	for j in range(0, sheet.col_count - p.col):
		if len(userList[j]) > 0:
			if userList[j][0] == 'P':
				numUsers += 1
	print 'Found {0} users in total, spanning columns {1} to {2}.'.format(numUsers, p.col, numUsers+p.col)

	if mode == 'batch':
		latestCategory = ''
		passedOverTitle = False
		if type == 'Default':
			for i in range(p.row + 2, sheet.row_count - p.col):
				if not passedOverTitle:
					if sheet.cell(i, m.col).value == 'T':
						latestCategory = sheet.cell(i, s.col).value
						print '+ Found category \'{0}\' on line {1}.'.format(latestCategory, i)
						passedOverTitle = True
					elif not passedOverTitle:
						latestCategory = get_category(sheet, i, p.row, m.col, s.col)	
				for j in range(p.col, p.col + numUsers):
					val = sheet.cell(i, j)
					if val.value is None:
						# Discard empty cells.
						pass
					elif val.value == '':
						# Discard empty cells.
						pass
					else:
						issue = { 'cell': val, 'desc': sheet.cell(i, s.col).value, 'study': studyName, 'participant': sheet.cell(p.row+1, j).value, 'category': latestCategory }
						times.append(issue)
						print '+ Found timestamp: {0}'.format(val.value)
		else:
			category = raw_input('Which category would you like to work in?\n>> ')
			times = generate_category(sheet, p, m, s, numUsers, studyName, category)
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

			print '\nIssue titled: {0}\n'.format(sheet.cell(lineSelect, s.col).value)
			yn = raw_input('Is this the correct issue? y/n\n>> ')
			if yn == 'y':
				break
			else:
				pass

		latestCategory = get_category(sheet, lineSelect, p.row, m.col, s.col)

		for j in range(p.col, p.col + numUsers):
			val = sheet.cell(lineSelect, j)
			if val.value is None:
				# Discard empty cells.
				pass
			elif val.value == '':
				# Discard empty cells.
				pass
			else:
				issue = { 'cell': val, 'desc': sheet.cell(lineSelect, s.col).value, 'study': studyName, 'participant': sheet.cell(p.row+1, j).value, 'category': latestCategory }
				times.append(issue)
				print '+ Found timestamp: {0}'.format(val.value.replace('\n',' '))
	elif mode == 'range':
		# This mode generates videos for all issues found in a range or span of issues.
		while True:
			try:
				startLineSelect = int(raw_input('\nWhich starting line (row number only)?\n>> '))
				endLineSelect = int(raw_input('\nWhich ending line (row number only)?\n>> '))
			except ValueError:
				startLineSelect = int(raw_input('\nTry again. Starting line (row number only)?\n>> '))
				endLineSelect = int(raw_input('\nTry again. Ending line (row number only)?\n>> '))
			print 'Lines selected: {0} to {1}'.format(sheet.cell(startLineSelect, s.col).value, sheet.cell(endLineSelect, s.col).value)
			yn = raw_input('Is this correct? y/n\n>> ')
			if yn == 'y':
				break
			else:
				pass
		
		passedOverTitle = False
		for i in range(startLineSelect, endLineSelect+1):
			if not passedOverTitle:
				if sheet.cell(i, m.col).value == 'T':
					latestCategory = sheet.cell(i, s.col).value
					print '+ Found category \'{0}\' on line {1}.'.format(latestCategory, i)
					passedOverTitle = True
				elif not passedOverTitle:
					latestCategory = get_category(sheet, i, p.row, m.col, s.col)
			for j in range(p.col, p.col + numUsers):
				val = sheet.cell(i, j)
				if val.value is None:
					# Discard empty cells.
					pass
				elif val.value == '':
					# Discard empty cells.
					pass
				else:
					issue = { 'cell': val, 'desc': sheet.cell(i, s.col).value, 'study': studyName, 'participant': sheet.cell(p.row+1, j).value, 'category': latestCategory }
					times.append(issue)
					print '+ Found timestamp: {0}'.format(val.value)
	elif mode == 'select':
		# This mode generates a list of non-completed issues and lets user select from those.
		pass

	return times

def generate_category(sheet, p, m, s, numUsers, studyName, category):
	# TODO
	# Case-insensitive category matching.
	times = []
	catCell = sheet.find(category)
	if sheet.cell(catCell.row, m.col).value == 'T':
		print '+ Found category \'{1}\' on line {0}.'.format(catCell.row, category)
		for i in range(catCell.row+1, sheet.row_count - p.col):
			for j in range(p.col, p.col + numUsers):
				if sheet.cell(i, m.col).value != 'T':
					if DEBUGGING: print '! {0}'.format(sheet.cell(i, j))
					val = sheet.cell(i, j)
					if val.value is None:
						# Discard empty cells.
						pass
					elif val.value == '':
						# Discard empty cells.
						pass
					else:
						issue = { 'cell': val, 'desc': sheet.cell(i, s.col).value, 'study': studyName, 'participant': sheet.cell(p.row+1, j).value, 'category': category }
						times.append(issue)
						print '+ Found timestamp: {0}'.format(val.value)
				else:
					if DEBUGGING: print '! Encountered other category, stopping category batch call'
					return times

def get_category(sheet, startingRow, pRow, mCol, sCol):
	category = ''
	while category == '':
		try:
			for i in range(startingRow, pRow, -1):
				if sheet.cell(i, mCol).value == 'T':
					category = sheet.cell(i, sCol).value
					print '+ Found category \'{0}\' on line {1}.'.format(category, i)
					break # Exit the for loop so we don't keep going up.
		except IndexError:
			break
	return category

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
			filename = check_filename_length(filename, step)
			break
	return filename

def check_filename_length(filename, step=1):
	if len(filename) > 255:
		if step > 1:
			if DEBUGGING: print '! Filename was longer than 255 chars ({0}, length {1})'.format(filename, len(filename))
			filename = filename[0:255-(1+len(str(step))+len(FILEFORMAT))] + '-' + str(step) + FILEFORMAT
		else:
			filename = filename[0:255-(len(FILEFORMAT))] + FILEFORMAT
	return filename

def clean_issue(issue):
	timeStamps = []
	unparsedTimes = issue['cell'].value.lower().split()
	if unparsedTimes == issue['cell'].value:
		unparsedTimes = unparsedTimes.split('+').split(',')
	
	# Using own iterator here, instead of letting the for-loop set this up. Otherwise we can't manually advance the iterator (we need to step twice
	# which continue won't do.)
	lines = iter(range(0,len(unparsedTimes)))
	issue['interview'] = []

	for i in lines:
		if DEBUGGING: print '! Cleaning timestamp {0}'.format(unparsedTimes[i])
		unparsedTimes[i] = unparsedTimes[i].strip().rstrip(',').rstrip('-')
		if unparsedTimes[i] == '':
			pass
		elif unparsedTimes[i].find('interview') != -1:
			issue['interview'].append(len(timeStamps))
			# The reason we use i+1 everywhere in this block is because of us doing the advancing at the end. Should probably still work if we moved the next() up top here.
			# TODO this goes out of index if we do i+1 and there is only one occurence/timestamp available to check
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
	issue['desc'] = issue['desc'].replace('\\','-')
	issue['desc'] = issue['desc'].replace('/','-')
	issue['desc'] = issue['desc'].replace('?','_')
	issue['desc'] = issue['desc'].translate({ord(i):None for i in '\"\'.><|:'})
	
	return issue

def ffmpeg(inputfile, outputfile, startpos, outpos, reencode):
	# TODO
	# Protect against videos that have an outtime beyond base video length

	# DEBUG
	# Just makes the clip a minute long if we didn't get an in-time
	if outpos == '00:00:00':
		outpos = add_duration(startpos)

	duration = get_duration(startpos, outpos)

	if duration < 0:
		print 'Can\'t work with negative duration for videos, exiting.'
		sys.exit(0)
	elif duration > 60*5:
		yn = raw_input('This video is over 5 minutes long, do you want to still generate it? (y/n)\n>> ')
		if yn == 'n':
			return None

	print 'Cutting {0} from {1} to {2}.'.format(inputfile, startpos, outpos)
	if DEBUGGING:
		print '! Debugging enabled, not attempting to call ffmpeg or output any files.'
	else:
		try:
			if not reencode:
				subprocess.call(['ffmpeg', '-y', '-loglevel', '16', '-ss', startpos, '-i', inputfile, '-t', str(duration), '-c', 'copy', '-avoid_negative_ts', '1', outputfile])
			else:
				# If we do this, we will re-encode the video, but resolve all issues with with iframes early and late.
				subprocess.call(['ffmpeg', '-y', '-loglevel', '16', '-ss', startpos, '-i', inputfile, '-t', str(duration), outputfile])
			print '+ Generated video \'{0}\' successfully.\n File size: {1}\n Expected duration: {2} s\n'.format(outputfile,filesize(os.path.getsize(outputfile)), duration)
		except WindowsError as e:
			print '\n! ERROR ffmpeg could not successfully run.\n  clipgen returned the following error:\n  {0}\n'.format(e)

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
	if DEBUGGING: print '! Debug mode is ON. Several limitations apply and more things will be printed.'
	# Remember that documents need to be shared to the email found in the json-file for OAuth-ing to work.
	# Each user of this program should also have their own, unique json-file (generate this on the Google Developer API website).
	scope = ['https://spreadsheets.google.com/feeds',
	 		 'https://www.googleapis.com/auth/drive']
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
		inputName = raw_input('\nPlease enter the index, name, URL or key of the spreadsheet (\'all\' for list, \'new\' for list of newest, \'last\' to immediately open latest):\n>> ')
		try:
			if inputName[:4] == 'http':
				# In case user copies a URL, we can handle that.
				worksheet = gc.open_by_url(inputName).worksheet(SHEET_NAME)
				break
			elif inputName[:3] == 'all':
				# Lists all Sheets, prefixed by a number.
				docList = get_alldocs(gc).split(',')
				print '\nAvailable documents:'
				for i in range(len(docList)):
					print '{0}. {1}'.format(i+1, docList[i].strip())
			elif inputName[:3] == 'new':
				# Typing 'new' shows the three latest Sheets (handy in case we have dozens of Sheets later).
				docList = get_alldocs(gc).split(',')
				print '\nNewest documents:'
				for i in range(3):
					print '{0}. {1}'.format(i+1, docList[i].strip())
			elif inputName[:4] == 'last':
				# This is equivalent to opening the Sheet numbered 1 in the 'all' list.
				latest = get_alldocs(gc).split(',')[0]
				worksheet = gc.open(latest).worksheet(SHEET_NAME)
  				break
  			elif inputName[0].isdigit():
  				# If user enters a number, we open the Sheet of that number from the 'all' list.
  				i = int(inputName)-1
  				worksheet = gc.open(get_alldocs(gc).split(',')[i].strip()).worksheet(SHEET_NAME)
  				break
			elif inputName.find(' ') == -1:
				# If user has entered text that has no spaces (and hasn't been caught as a number, per above) we try to open it as a GID key.
				worksheet = gc.open_by_key(inputName).worksheet(SHEET_NAME)
				break
			else:
				# As we have free text entry, we match it to a Sheet name (regardless of case) and then open that Sheet.
				inputName = inputName.strip().lower()
				docList = get_alldocs(gc).split(',')
				for i in range(len(docList)):
					if docList[i].strip().lower() == inputName:
						worksheet = gc.open(docList[i]).worksheet(SHEET_NAME)
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

	print 'Connected to Google Drive!'
	inputModeFails = 0

	while True:
		while True:
			inputMode = raw_input('\nSelect mode: (b)atch, (r)ange, (c)ategory or (l)ine\n>> ')
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
				elif inputMode[0] == 'c' or inputMode == 'cat' or inputMode == 'category':
					timesList = generate_list(worksheet, 'batch', 'category')
					break
				elif inputMode == 'positive':
					timesList = generate_list(worksheet, 'batch', 'Positive')
					break
				#elif inputMode[0] == 's' or inputMode == 'select':
				#	timesList = generate_list(worksheet, 'select')
				#	break
				elif inputMode == 'karl':
					plogo()
			except (IndexError, gspread.HTTPError) as e:
				inputModeFails += 1
				try:
					gc = gspread.authorize(credentials)
				except gspread.AuthenticationError as e:
					print e
					print 'Could not authenticate.'
					sys.exit(0)

		print '\n* ffmpeg is set to never prompt for input and will always overwrite.\n  Only warns if close to crashing.\n'
		videosGenerated = 0

		for i in range(0, len(timesList)):
			# timesList is a list containing issues, one per index
			# issues are dicts that hold:
			# - cell 			Full gspread Cell object (row, col, value)
			# - desc 			String, summary description of the issue
			# - study 			String, name of the study
			# - participant 	String, participant ID (without prefix)
			# - times 			List, contains one timestamp pair (as a tuple) per index
			# - interview 		List, contains indices of timestamps that are from interviews
			# - category 		String, category heading found over issue
			# Note that the 'times' entry in the dict is generated during the clean_issue method call.

			timesList[i] = clean_issue(timesList[i])
			for j in range(0,len(timesList[i]['times'])):
				vidIn, vidOut = timesList[i]['times'][j]
				vidName = check_filename('[Study ' + filter(unicode.isdigit, timesList[i]['study']) + '][' + timesList[i]['category'] + '] ' + timesList[i]['desc'] + FILEFORMAT)

				if timesList[i]['interview'].count(j) > 0:
					if DEBUGGING: print '! Timestamp had interview'
					baseVideo = timesList[i]['study'] + '_interview_' + timesList[i]['participant'] + FILEFORMAT
				else:
					baseVideo = timesList[i]['study'] + '_' + timesList[i]['participant']  + FILEFORMAT
				
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