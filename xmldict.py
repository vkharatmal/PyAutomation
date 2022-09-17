#import xmltodict
from os import walk

class test:
    def __init__(self,paths):
    
        path = paths
        #path = 'CZ2_02052022_00000_vkharatmal_ATK-FG-DLY-SHIRE_19Jul2022-10-25-07.xml'
        paths=paths.split('\\')[2]
        print(paths)
        xmldata = '<root>' + open(path,'r').read() + '</root>'
        dict_data=xmltodict.parse(xmldata)
        #print(dict_data['root']["ns0:TimeCardData"]["TimeEntries"]["TimeEntry"])
        self.sourceid=paths.split('_')[4]
        self.opco=paths.split('_')[0]
        self.tester=paths.split('_')[3]
        self.originalfileid=dict_data['root']["ns0:TimeCardData"]["TimeEntries"]["OriginalFileId"]
        self.originalfilepath=dict_data['root']["ns0:TimeCardData"]["TimeEntries"]["OriginalFilePath"]+dict_data['root']["ns0:TimeCardData"]["TimeEntries"]["OriginalFileId"]
        self.htmlpath=dict_data['root']["ns0:TimeCardData"]["TimeEntries"]["PathToDerivedSupportingDocument"]
        print(self.originalfilepath)
        temp=self.htmlpath.split('\\')[4]
        self.generatedhtmlpath=".\Supporting_docs"+"\\" + temp
        self.filetype=dict_data['root']["ns0:TimeCardData"]["TimeEntries"]["OriginalFileType"]
        self.supp_docpath=dict_data['root']["ns0:TimeCardData"]["TimeEntries"]["PathToDerivedSupportingDocument"]
        #print(type(dict_data['root']["ns0:TimeCardData"]["TimeEntries"]["TimeEntry"]))
        self.timecoll_type=dict_data['root']["ns0:TimeCardData"]["TimeEntries"]["TimeEntry"][0]["TimeCollectionType"]
        self.weekenddate=dict_data['root']["ns0:TimeCardData"]["TimeEntries"]["TimeEntry"][0]["PAY_END_DT"]
        #print('Timeentry is in list')
        
        #return sourceid,opco,tester,originalfileid,originalfilepath,htmlpath,generatedhtmlpath,filetype,supp_docpath,timecoll_type,weekenddate
    
def xmldataValues(paths):
    return test(paths)
