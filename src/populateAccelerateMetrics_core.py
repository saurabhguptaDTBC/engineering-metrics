import sys
from sheetWriter_kb import *
if __name__ == '__main__':
#    if len( sys.argv ) > 1:
#        lvProgramMode == sys.argv[1]#
#
    book = openWorkbook()    
    #populateStories(book,lvProgramMode)
    populateStories(book)
    #populateReleases(book)
    saveWorkbook(book)