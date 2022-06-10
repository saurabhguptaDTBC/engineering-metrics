from sheetWriter import *

if __name__ == '__main__':
    book = openWorkbook()
    populateStories(book)
    populateReleases(book)
    saveWorkbook(book)