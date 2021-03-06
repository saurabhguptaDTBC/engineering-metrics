from distutils.command.config import config
from config import *
import openpyxl
from asyncore import loop
from helperFunctions import ifnull, resetSheet, requestHelper
# vIterations=set()
# vReleases=set()

def openWorkbook():
   return openpyxl.load_workbook(gvPath)

def saveWorkbook(book):
    book.save(gvPath)

def populateStories(book):
    resetSheet(gvSheetNameStories,book)
    for loopProjects in gvProjects:
        for loopTeams in gvTeams:
            sheet = book[gvSheetNameStories]
            URL="https://drivetime.tpondemand.com/api/v1/UserStories?where=(Team.Name%20eq%20%27BC%20Digital%20"+loopTeams+"%27)and(Project.ID%20eq%20"+ str(loopProjects) +")&include=[Id,Name,Effort,Project,Iteration[Name],Team[Name],Feature[Name],LeadTime,CycleTime,Release,TeamIteration[Id],EntityState[Name]]&append=[Bugs-count]&take=1000&orderbydesc=Effort&format=json"
            while URL != "":
                # req = requests.get(URL+gvTPToken)
                # text_data= req.text
                # json_dict= json.loads(text_data)
                json_dict= requestHelper(URL)
                iterationReleaseList=[]
                vIterationRelease=""
                releaseCount=0
                for item in json_dict['Items']:
                # The following block is added to get unique releases in iteration : excel wasn't giving me an easy way to do this
                    if ifnull(item['Release'],"null",'Id') != "null" and ifnull(item['TeamIteration'],"null",'Id') != "null":
                        vIterationRelease=str(ifnull(item['Release'],"null",'Id'))+":"+str(ifnull(item['TeamIteration'],"null",'Id'))
                        if vIterationRelease not in iterationReleaseList:
                            iterationReleaseList.append (vIterationRelease)   
                            releaseCount=1
                        else:
                            releaseCount=0
                    else:
                        releaseCount=0
                    row = [item['Id'], item['Name'],item['Effort'],ifnull(item['Project'],"null",'Name'),ifnull(item['Team'],"null",'Name'),ifnull(item['Feature'],"null",'Name'), item['LeadTime'],item['CycleTime'], ifnull(item['Release'],"",'Id'),ifnull(item['TeamIteration'],"",'Id'),ifnull(item['EntityState'],"null",'Name'),item['Bugs-Count'],releaseCount]
                    sheet.append(row)
                    # if item['Release'] is not None:
                        # vReleases.add(item['Release']['Id'])
                    # if item['TeamIteration'] is not None:
                        # vIterations.add(item['TeamIteration']['Id'])
                try:
                    URL=json_dict['Next']
                except:
                    URL=""



def populateReleases(book):
    resetSheet(gvSheetNameReleases,book)
    sheet=book[gvSheetNameReleases]
    # for i in vReleases:
    #     sheet = book[vSheetNameReleases]
    #     URL="https://drivetime.tpondemand.com/api/v1/Releases/"+str(i)+"?include=[Name,EndDate,Effort,Owner]&format=json&dateformat=iso"
    #     json_dict= requestHelper(URL)
    #     if json_dict != "Error":
    #         row = [json_dict['Id'],json_dict['Name'],json_dict['EndDate'],json_dict['Effort'],json_dict['Owner']['FullName']]
    #         print(i)
    #         sheet.append(row)
    for loopProjects in gvProjects:
        URL="https://drivetime.tpondemand.com/api/v1/Releases?where=(Projects.ID%20eq%20"+str(loopProjects)+")&include=[Name,EndDate,Effort,Owner]&format=json&dateformat=iso&take=1000"
        json_dict= requestHelper(URL)
        if json_dict != "Error":
            for item in json_dict['Items']:
                row = [item['Id'],item['Name'],item['EndDate'],item['Effort'],item['Owner']['FullName']]
                sheet.append(row)