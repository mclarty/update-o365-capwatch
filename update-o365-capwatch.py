import csv
import re

### DEFINE VARIABLES HERE ###

# Match member contacts with email addresses ending in @flwg.cap.gov

emailpattern = r'@flwg\.cap\.gov'

### END VARIABLES ###

### START FUNCTIONS ###

# Format the member's organization to populate the Company field

def formatUnitName(org):
	return f"{org['Region']}-{org['Wing']}-{org['Unit']} {org['Name']}"

# Format the member's organization to populate the Display Name field

def formatDisplayUnit(org):
	str = org['Region']

	if (org['Scope'] != 'REGION'):
		str += f" {org['Wing']}WG"

		if (org['Scope'] != 'WING'):
			if (org['Scope'] == 'GROUP'):
				group = re.search('(GROUP|GP) (.?)', org['Name'])
				str += f" GP{group[2]}"

			if (org['Scope'] != 'GROUP'):
				grouporg = orgs[org['NextLevel']]
				if (grouporg['Scope'] != 'GROUP'):
					str += f" {org['Wing']}{org['Unit']}"
				else:
					group = re.search('(GROUP|GP) (.?)', grouporg['Name'])
					str += f" GP{group[2]} {org['Wing']}{org['Unit']}"

	return str

# Format the member's Display Name field

def formatDisplayName(user):
	middleinitial = f" {user['middleinitial'].upper()}"

	if (len(user['middleinitial']) > 0):
		middleinitial += " "

	if ('physicaldeliveryofficename' in user):
		officesymbol = f"/{user['physicaldeliveryofficename']}"
	else:
		officesymbol = ""

	return f"{user['sn'].upper()}, {user['givenname'].upper()}{middleinitial}{user['grade']} CAP {user['displayunit']}{officesymbol}"

# Format the member's duty position title and office symbol
#
# Business Logic: 
# 1. Duty position must equal member's unit of assignment
# 2. Primary duty positions are used first, in sort order by priority defined in office-symbols.csv
# 3. Assistant duty positions are used if no primary duty position is found, by same sort order

def formatDutyPosition(member, asst):
	candidatepositions = []
	if (asst == True):
		asstflag = 1
	else:
		asstflag = 0

	for row in dutypos:
		if (row['CAPID'] == member['CAPID'] and row['ORGID'] == member['ORGID'] and int(row['Asst']) == asstflag):
			row['FAS'] = officesymbol[row['Duty']][row['Lvl']]
			row['SortOrder'] = officesymbol[row['Duty']]['SortOrder']
			candidatepositions.append(row)

	if (len(candidatepositions) > 0):
		candidatepositions = sorted(candidatepositions, key=lambda x: x['SortOrder'])
		return candidatepositions[0]
	else:
		return None

### END FUNCTIONS ###

# Read in the organization file to populate an orgs dictionary

with open('capwatch/Organization.txt') as orgfile:
	orgreader = csv.DictReader(orgfile)
	orgs = {}
	for row in orgreader:
		orgs[row['ORGID']] = row

# Read in the member contact file to populate a users dictionary with members that have address in eServices matching emailpattern

with open('capwatch/MbrContact.txt') as contactfile:
	contactreader = csv.DictReader(contactfile)
	users = {}
	for row in contactreader:
		if (re.search(emailpattern, row['Contact'], re.IGNORECASE)):
			users[row['CAPID']] = {}
			users[row['CAPID']]['email'] = row['Contact'].strip().lower()

# Read in the member file to populate a members dictionary

with open('capwatch/Member.txt') as memberfile:
	memberreader = csv.DictReader(memberfile)
	members = {}
	for row in memberreader:
		members[row['CAPID']] = row

# Read in the office symbols file to populate an office symbols dictionary

with open('office-symbols.csv') as officesymbolfile:
	officesymbolreader = csv.DictReader(officesymbolfile)
	officesymbol = {}
	for row in officesymbolreader:
		if (row['Duty'] not in officesymbol):
			officesymbol[row['Duty']] = {}
			officesymbol[row['Duty']]['SortOrder'] = row['SortOrder']
		levels = row['Lvl'].split(",")
		for level in levels:
			officesymbol[row['Duty']][level] = row['OfficeSymbol']

# Read in the duty position file to populate a duty positions dictionary

with open('capwatch/DutyPosition.txt') as dutyposfile:
	dutyposreader = csv.DictReader(dutyposfile)
	dutypos = []
	for row in dutyposreader:
		dutypos.append(row)

# Iterate the users dictionary to populate attributes from the members dictionary and formatting functions

for capid, row in users.items():

	# Set user name attributes

	users[capid]['givenname'] = members[capid]['NameFirst']
	users[capid]['sn'] = members[capid]['NameLast']
	if (len(members[capid]['NameMiddle']) > 0):
		users[capid]['middleinitial'] = members[capid]['NameMiddle'][0]
	else:
		users[capid]['middleinitial'] = ""

	# Set member grade for use in creating display name

	if (members[capid]['Type'] == "CADET"):
		users[capid]['grade'] = "CADET"
	else:
		users[capid]['grade'] = members[capid]['Rank']

	# Set member company based on unit of assignment

	users[capid]['company'] = formatUnitName(orgs[members[capid]['ORGID']])

	# Temporary unit name for use in creating display name

	users[capid]['displayunit'] = formatDisplayUnit(orgs[members[capid]['ORGID']])

	# Set member title and office symbol

	title = formatDutyPosition(members[capid], False)
	if (title is None):
		title = formatDutyPosition(members[capid], True)
		if (title):
			title['Duty'] = f'Assistant {title['Duty']}'
	if (title):
		users[capid]['title'] = title['Duty']
		users[capid]['physicaldeliveryofficename'] = title['FAS']

	# Call function to set display name

	users[capid]['displayname'] = formatDisplayName(users[capid])

	# Delete unused attributes
	
	del users[capid]['displayunit']

# Output the users dictionary display names for validation
# TODO: Write the output to a CSV file in a format to be imported to Office 365 via PowerShell

print(users)