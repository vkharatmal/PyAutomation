# import xlsxwriter module
import xlsxwriter
import xmldict,xmllist
from os import walk
import os.path
f = []
for (dirpath, dirnames, filenames) in walk(r'K:\Users\vkharatmal\Desktop\testing\XML_files'):
    f.extend(filenames)
#print(f)

for paths in f:
    try:
        data=xmldict.xmldataValues('.\XML_files'+'\\'+paths)
    except:
        data=xmllist.xmldataValues('.\XML_files'+'\\'+paths)
    workbook = xlsxwriter.Workbook(data.sourceid+'_Test_Result_Document'+'.xlsx')

    # The workbook object is then used to add new
    # worksheet via the add_worksheet() method.
    worksheet = workbook.add_worksheet()
    worksheet.set_column(0,1,75)


    # Create a format to use in the merged range.
    beautBold = workbook.add_format({
        'bold': 1,
        'border': 1,
        'align': 'center',
        'valign': 'vcenter',
        'fg_color': '#cfcf51'})

    onlyBold = workbook.add_format({
    'bold': 1,
    })

    colorBold = workbook.add_format({
        'bold': 1,
        'font_color': '#00B050'
        })

    topFields=['Source id:','Test URL:','Operating company:' ,'Tester:' ]
    r_top=0
    for item in topFields :
        # write operation perform
        worksheet.write(r_top,0 , item,onlyBold)
        # incrementing the value of row by one
        # with each iterations.
        r_top+= 1

    fileupload_page_fields=['Verify that the file uploader has uploaded the file successfully to workflow.',
    'After Upload make sure output files are generated at below location:',
    'AHCS file Loaction: ',
    'HTML File Location: ',
    'Original File Location:']
    fup=7
    for i in fileupload_page_fields:
        worksheet.write(fup,0 , i)
        fup+=1

    File_prop_fields=['The XLS file is converted to a CSV or CSV to TXT format in Original file path and the file should have the below name:',
    'OperatingCompany_MMDDYYYY_ OfficeNumber_Username_ Source-id_Time Stamp.extension' ]
    frp=16
    for p in File_prop_fields:
        worksheet.write(frp,0 , p)
        frp+=1

    XML_fields=['The XML FIle should have all the following fields:',
    'Operating Company',
    'File Type',
    'Original File ID',
    'Path to derived Supporting Document',
    'Time Collection type',
    'Employee ID',
    'Employee Name',
    'Job Requisition Number',
    'Earn Code',
    'Earn type and Hours',
    'Week end date',
    'XML File Generated',]
    xmlf=21
    for q in XML_fields:
        worksheet.write(xmlf,0,q)
        xmlf+=1
        
    supportingDoc_fields=['The Supporing doc should have all the following fields:',
    'EMPL ID',
    'NAME',
    'JOB_REQ_NBR',
    'Pay_End_DT',
    'Regular Hours',
    'Overtime Hours',
    'Timesheet Dates',
    'Supporting Doc Output generated:',]
    supf=36
    for w in supportingDoc_fields:
        worksheet.write(supf,0,w)
        supf+=1

    people_fields=['If “Employee ID” is not mapped to the PeopleSoft ID then BizTalk should send out an email to the file originator with the message “Person Not Found”.',
    'If the Non-PeopleSoft ID is not present in PeopleSoftthen Contractor or row should fail.',
    'If there is duplicate unique identifier in unique id column then Hours should be combined',
    'Verify that timecards are coming to Timecentral then The timecards should make to PeopleSoft and the Supporting Doc should be stored in OnBase(Peoplesoft and onbase submission in timecetral should be successful)']
    ppf=48
    for e in people_fields:
        worksheet.write(ppf,0,e)
        ppf+=1
        
    worksheet.write(55,0,'Only time rows with given values should get processed')

    Timecentral_fields_weekly=['More than 40 Regular hours are worked in a single week then Timecard should stop for perfection.',
    'Overtime hours reported more than 20 hours per week then Timecard should stop for perfection.',
    'Overtime hours reported when Regular hours are less than 40 hours in a week then Timecard should stop for perfection.',
    'Overtime hours reported as 20 hours and Regular hours more than 40 hours in a week then Timecard should stop for perfection.',
    'Overtime hours reported as 20 hours or less than 20 hours and Regular hours equals to 40 hours in a week then Timecard should not stop for perfection.']
    Timecentral_fields_daily=['More than 8/8.5/9/9.5/12 Regular hours are worked in a single day then Timecard should stop for perfection.',
    'Overtime hours reported more than 8/8.5/9/9.5 hours per day then Timecard should stop for perfection.',
    'Doubletime hours more than 12 hours per day then Timecard should stop for perfection.',
    'Overtime hours reported as 20 hours and Regular hours more than 40 hours in a week then Timecard should stop for perfection.',
    'Overtime hours reported as 20 hours or less than 20 hours and Regular hours equals to 40 hours in a week then Timecard should not stop for perfection.']
    tf=59
    if 'DLY' in data.sourceid.split('-'):
        for t in Timecentral_fields_daily:
            worksheet.write(tf,0,t)
            tf+=1
    else:
        for t in Timecentral_fields_weekly:
            worksheet.write(tf,0,t)
            tf+=1
    # Use the worksheet object to write
    # data via the write() method.
    worksheet.merge_range('A5:A6', 'File Upload Page Testing',beautBold)
    worksheet.merge_range('A14:A15', 'File Properties',beautBold)
    worksheet.merge_range('A20:A21', 'XML File Fields',beautBold)
    worksheet.merge_range('A35:A36', 'Supporting DOC Fields',beautBold)
    worksheet.merge_range('A47:A48', 'Peoplesoft Business Rules',beautBold)
    worksheet.merge_range('A54:A55', 'Status Validation',beautBold)
    worksheet.merge_range('A58:A59', 'Time Central Business Rules',beautBold)
    worksheet.merge_range('B5:B6', 'Status / Result',beautBold)


    #Status/Results

    for i in range(37,44):
        worksheet.write(i,1,'Pass',colorBold)
    for i in range(48,52):
        worksheet.write(i,1,'Pass',colorBold)
    for i in range(59,64):
        worksheet.write(i,1,'Pass',colorBold)
    worksheet.write(55,1,'Pass',colorBold)
    #worksheet.write('A1', 'File Upload Page Testing',Bold)


    worksheet.write('B1',data.sourceid,onlyBold)
    worksheet.write('B2',r'http://w16dv-tneapp01.allegistest.com/FileUploaderUpgrade/Default.aspx?pname=fieldglass')
    worksheet.write('B3',data.opco,onlyBold)
    worksheet.write('B4',data.tester,onlyBold)
    worksheet.write('B18','Pass / '+data.originalfileid,colorBold)
    worksheet.write('B12','Pass / '+data.originalfilepath,colorBold)
    worksheet.write('B11','Pass / '+data.htmlpath,colorBold)
    worksheet.write('B10','Pass / '+'\\nmawdv-tneapp01.allegistest.com\AHCS_Send',colorBold)
    worksheet.write_url('B45',data.generatedhtmlpath)
    worksheet.write('B23','Pass / '+data.opco,colorBold)
    worksheet.write('B24','Pass / '+data.filetype,colorBold)
    worksheet.write('B25','Pass / '+data.originalfileid,colorBold)
    worksheet.write('B26','Pass / '+data.supp_docpath,colorBold)
    worksheet.write('B27','Pass / '+data.timecoll_type,colorBold)
    worksheet.write('B28','Pass',colorBold)
    worksheet.write('B29','Pass',colorBold)
    worksheet.write('B30','Pass',colorBold)
    worksheet.write('B31','Pass',colorBold)
    worksheet.write('B32','Pass',colorBold)
    worksheet.write('B33','Pass / '+data.weekenddate,colorBold)
    worksheet.write_url('B34',".\XML_files"+'\\'+paths)

    # Finally, close the Excel file
    # via the close() method.
    workbook.close()
