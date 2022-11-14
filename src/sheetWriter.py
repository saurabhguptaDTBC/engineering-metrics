from distutils.command.config import config
from config_digital import *
import openpyxl
from asyncore import loop
from helperFunctions import ifnull, resetSheet, requestHelper
from datetime import *
# vIterations=set()
# vReleases=set()
import csv

def openWorkbook():
   return openpyxl.load_workbook(gvPath)

def saveWorkbook(book):
    book.save(gvPath)

def getTeamIterationDetails(iterationNumber):
    URL="https://drivetime.tpondemand.com/api/v1/TeamIteration/"+str(iterationNumber)+"?format=json&dateformat=iso&include=[Id,Name,StartDate,EndDate,Effort,Team[Name]]"
    return requestHelper(URL)

def getPersonMetrics():
    f= open('VenuAndTeam.csv', 'w', newline='') 
    writer = csv.writer(f)
    dataHeader=['Id','Name','Effort','Project','Team','Feature','State','TeamIteration','EndDate','Developer','Type']
    writer.writerow(dataHeader)
    # Sushma URL="https://drivetime.tpondemand.com/api/v1/Assignments?where=(GeneralUser.Id%20eq%20684)&format=json&dateformat=iso&take=1000"
    URL="https://drivetime.tpondemand.com/api/v1/Assignments?where=(GeneralUser.Id%20in%20(678,686,685))&format=json&dateformat=iso&take=1000"
    json_dict= requestHelper(URL)
    i=1
    for item in json_dict['Items']:
        URL2="https://drivetime.tpondemand.com/api/v1/UserStories?where=(ID%20eq%20"+str(item['Assignable']['Id'])+")&format=json&dateformat=iso&include=[Id,Name,Effort,Project,Iteration[Name],Team[Name],Feature[Name],LeadTime,CycleTime,Release,TeamIteration[Id],EntityState[Name],TeamIteration,EndDate]"
        json_dict2= requestHelper(URL2)
#        print(str(item['Assignable']['Id']) )
        for item2 in json_dict2['Items']:
            if item2['EndDate'] is not None:
                lvDate=item2['EndDate'][0:10]
            else:
                lvDate=""
            data = [str(item['Assignable']['Id']) ,item['Assignable']['Name'] , str(item2['Effort']) ,ifnull(item2['Project'],"null",'Name') , ifnull(item2['Team'],"null",'Name') ,ifnull(item2['Feature'],"null",'Name'),ifnull(item2['EntityState'],"null",'Name'),ifnull(item2['TeamIteration'],"null",'Name'),lvDate,ifnull(item['GeneralUser'],"null",'FullName'),'Story']
            writer.writerow(data)
        URL3="https://drivetime.tpondemand.com/api/v1/Tasks?where=(ID%20eq%20"+str(item['Assignable']['Id'])+")&format=json&dateformat=iso&include=[Id,Name,Team,Effort,Project[Name],TeamIteration[Name],EntityState[Name],EndDate]"
        json_dict3= requestHelper(URL3)
        for item3 in json_dict3['Items']:
            if item3['EndDate'] is not None:
                lvDate=item3['EndDate'][0:10]
            else:
                lvDate=""
            data = [str(item['Assignable']['Id']) , item['Assignable']['Name'] , str(item3['Effort']) , ifnull(item3['Project'],"null",'Name'), ifnull(item3['Team'],"null",'Name') ,"null" , ifnull(item2['EntityState'],"null",'Name'),ifnull(item3['TeamIteration'],"null",'Name'),lvDate,ifnull(item['GeneralUser'],"null",'FullName'),'Task']
            writer.writerow(data)
            

def populateStories(book):
    resetSheet(gvSheetNameStories,book,"Sprint")
    for loopProjects in gvProjects:
        for loopTeams in gvTeams:
            sheet = book[gvSheetNameStories]
            URL="https://drivetime.tpondemand.com/api/v1/UserStories?where=(Team.Name%20eq%20%27"+loopTeams+"%27)and(Project.ID%20eq%20"+ str(loopProjects) +")and(CreateDate%20gt%20%272022-01-01%27)&include=[Id,Name,Effort,Project,Iteration[Name],Team[Name],Feature[Name],LeadTime,CycleTime,Release,TeamIteration[Id],EntityState[Name]]&append=[Bugs-count]&take=1000&orderbydesc=Effort&format=json&dateformat=iso"
            while URL != "":
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
                    if item['TeamIteration'] is not None:
                        lvIterationObject=getTeamIterationDetails(item['TeamIteration']['Id'])
                        lvIterationName=lvIterationObject['Name']
                        lvIterationStartDate=lvIterationObject['StartDate'][0:10]
                        lvIterationEndDate=lvIterationObject['EndDate'][0:10]
                        lvPivot = lvIterationName + " : " + lvIterationStartDate + " - " + lvIterationEndDate
                    else:
                        lvIterationName=""
                        lvIterationStartDate=""
                        lvIterationEndDate=""
                        lvPivot =""
                       # if item['EndDate'] is not None:
                       #     lvWeek=datetime.fromisoformat(item['EndDate']).isocalendar()[1]
                       # else:
                       #     lvWeek=""

                    row = [item['Id'], item['Name'],item['Effort'],ifnull(item['Project'],"null",'Name'),ifnull(item['Team'],"null",'Name'),ifnull(item['Feature'],"null",'Name'), item['LeadTime'],item['CycleTime'], ifnull(item['Release'],"",'Id'),ifnull(item['TeamIteration'],"",'Id'),ifnull(item['EntityState'],"null",'Name'),item['Bugs-Count'],releaseCount,lvIterationName,lvIterationStartDate,lvIterationEndDate, lvPivot ]
                    sheet.append(row)
                    # if item['Release'] is not None:
                        # vReleases.add(item['Release']['Id'])
                    #if item['TeamIteration'] is not None
                        # vIterations.add(item['TeamIteration']['Id'])
                    #    lvIterationObject=getTeamIterationDetails(item['TeamIteration']['Id'])
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