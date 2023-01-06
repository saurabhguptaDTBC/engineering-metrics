from helperFunctions import ifnull, resetSheet, requestHelper
import sys
#from sheetWriter import *

if __name__ == '__main__':
    if len( sys.argv ) > 1:
       if(sys.argv[1]) == "kanban":
            from config_core import *
            from sheetWriter_kb import * 
            
            book = openWorkbook()    
            populateStories(book)
            saveWorkbook(book)
    else:
        from config_digital import *
        from sheetWriter import *

        book = openWorkbook()    
        resetSheet(gvSheetNameStories,book,"Sprint")
        resetSheet(gvSheetNameReleases,book,"Sprint")
        populateStories(book)
        populateBugs(book)
        populateReleases(book)
        saveWorkbook(book)