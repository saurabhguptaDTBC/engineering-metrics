import sys
#from sheetWriter import *

if __name__ == '__main__':
    if len( sys.argv ) > 1:
       if(sys.argv[1]) == "kanban":
            from config_core import *
            from sheetWriter_kb import *        
       else:
            from config_digital import *
            from sheetWriter import *
    else:
        from config_digital import *
        from sheetWriter import *

    book = openWorkbook()    
    #populateStories(book,lvProgramMode)
    populateStories(book)
    #getPersonMetrics()
    #populateReleases(book)
    saveWorkbook(book)