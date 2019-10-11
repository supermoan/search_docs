# search_docs.py
# Andre Moan <andre.moan@hi.no>

# This script searches through the contents of all word documents in a given directory and all
# its subdirectories, looking for matches to a set of user-defined strings. This can be used to
# identify all word documents that contain one of the specified strings. Matching is done
# with regular expressions, so it is quite flexible.

# Dependencies: To work with binary documents (doc files), the script depends on antiword, so that
# must be installed on your system and available to the user running the script. 

### SETTINGS
# SEARCH_PATH: Root directory from which to start searching. Subdirectories will also be searched.
SEARCH_PATH = "/Users/a21997/Documents/Bifangstestimering/Data/pelagisk/"

# SEARCH_STRINGS: This is a list of regular expressions that'll each be matched against the
# contents of each scanned document.
SEARCH_STRINGS = ["[h|k]val", "kn.l", "spekkh[o|u]gg[a|e]r", "nise", "spring[a|e]r", "[h|k]vitnos", "[h|k]vitskjeving",
    "delfin", "sj.pattedyr", "sel", "steinkobbe", "havert", "gr.nnlands?sel", "ringsel","klappmyss","storkobbe"]

# PRINT_PATH: When a match is found, print just the filename or the complete path? Set to false
# to only print the filename
PRINT_PATH = False

### DO NOT EDIT BELOW THIS LINE

import os
import zipfile
import re
import subprocess

# read_docx: unzips the specificed docx file and returns the text contents
def read_docx(filename):
	doc = zipfile.ZipFile(filename)
	content = doc.read("word/document.xml").decode("utf-8")
	content = re.sub("<(.|\n)*?>", "", content) # remove XML tags
	return(content)

# read_doc: executes antiword and returns the output
def read_doc(filename):
	p = subprocess.run(["antiword", "-m", "UTF-8.txt", filename], universal_newlines=True, stdout=subprocess.PIPE)
	return(p.stdout)
	
# search: determines the file type (doc or docx) for the file given and then calls the
# appropriate handler. Returns a list of all match results for the patterns in SEARCH_LIST
def search(filename):
	type = filename.split(".")[-1] # todo: use endswith?
	if type == "doc": 
		x = read_doc(filename)
	elif type == "docx":
		x = read_docx(filename)

	x = str(x).lower()

	matches = [0] * len(SEARCH_STRINGS)

	for i in range(0, len(SEARCH_STRINGS)):
		if re.search(SEARCH_STRINGS[i], x):
			matches[i] = 1
	return(matches)


files = []
matches = []

# Make a list of all the doc/docx files in the search path and subdirectories
for r, d, f in os.walk(SEARCH_PATH):
	for file in f:
		if '~$' not in file and ('docx' in file or 'doc' in file):
			files.append(os.path.join(r, file))

print("")
print("There are %d files in search path (including subfolders)..." % len(files))

# Scan each file and note down the filenames of any files that contained one or more matches.
for file in files:
	match = search(file)
	if sum(match) > 0:
		matches.append([file, match])

# Print feedback back to the user
if len(matches) == 0:
	print("Sorry, no matches!")
else:
	print("Found %d documents(s) containing matches!" % len(matches))
	print("")
	print(" #   %-40s %s" % ("MATCHES", "FILENAME"))
	for index, (filename, match) in enumerate(matches):
		z = [SEARCH_STRINGS[i] for i, m in enumerate(match) if m>0]
		z = ', '.join(z)[0:39]
		if PRINT_PATH == True:
			print(" %02d  %-40s %s" % (index+1, z, filename))
		else:
			print(" %02d  %-40s %s" % (index+1, z, os.path.basename(filename)))
	print("")


