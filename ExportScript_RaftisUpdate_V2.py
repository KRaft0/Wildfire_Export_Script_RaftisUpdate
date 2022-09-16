#********************************************************************************************************************
# Name: Wildfire Incident Export Script
# Created: 09/30/2021
# Author: Kaylene Raftis, Based on export script written by Paul Hoefler
# Description: This script iterates through the incident project folder and exports the specified layouts as pdfs.
#               The user will need to set up their folder locations, input incident information, and adjust export
#               settings for their particular incident.  Next the user should set up their export jobs following the
#               format shown.  Examples are also provided for single and series type export jobs.  For more information
#               about exporting a layout series see the attached documentation (Documentation.pdf).
#               You can toggle export jobs on/off by using comments.  Insert a # at the beginning of the line to toggle
#               off, remove the # to toggle on.  The script applies the date the map is intended for based on when the
#               script is run.  If run between 0000 and 1700 map names are set for today.  If run between 1700 and 2400
#               map names are set for tomorrow.
#               SAVE and CLOSE all projects that will be exported before you run the script!
#*********************************************************************************************************************

####DO NOT EDIT#########################################################
# import modules for the error catcher
import sys, string, os,traceback, datetime, time, winsound, random, csv
# import module for arcpy
import arcpy
# Current time stamp
currentDT= datetime.datetime.now()
print('Started at {0}'.format(currentDT))
########################################################################

#List basemaps by project here
stratbasemaps = ['SBS_63,360', 'NAIP']
opsbasemaps = ['Topo_80,000','Topo_24,000']
#######################################################################
export_jobs = [
### SET ITEMS TO EXPORT BELOW ##########################################
# export_jobs = [('APRX','LAYOUT','OutputNAME','[IMG|GEOREF|BOTH_SEP|BOTH_SAME]'],[SINGLE|SERIES],'MAP_NAME','MAPFRAME_ELEMENT_NAME','TITLE_ELEMENT_NAME','TITLE_TEXT','Bookmark_Wildcard','BasemapList','DesiredBasemap'),]
#                                                            **Only fill out 'MAP_NAME','MAPFRAME_ELEMENT_NAME','TITLE_ELEMENT_NAME','TITLE_TEXT,'Bookmark_Wildcard','BasemapList','DesiredBasemap' if the map is a series**

#****Single*******
#('pump_2022_DoubleCreek_ORWWF000400.aprx','Layout_36x48_port','pump_ArchE_port','IMG','SINGLE'),
#('transpo_85x11_2022_DoubleCreek_ORWWF000400.aprx','transpo_85x11_Port_DoubleCreek','transpo_85x11_Port','BOTH_SAME','SINGLE'),



#****Series*******
#('ops_2022_DoubleCreek_ORWWF000400.aprx','ops_series_ops_ArchE_Port_DoubleCreek','ops_ArchE_port','BOTH_SAME','SERIES','ops_2022_DoubleCreek','*Map Frame*', '', '', 'inc_80_', opsbasemaps,'Topo_80,000'),
#('ops_2022_DoubleCreek_ORWWF000400.aprx','ops_series_ops_ArchE_Port_DoubleCreek','ops_ArchE_port','BOTH_SAME','SERIES','ops_2022_DoubleCreek','*Map Frame*', '', '', 'inc_24_', opsbasemaps, 'Topo_24,000'),
#('ops_2022_DoubleCreek_ORWWF000400.aprx','ops_DIV_series_ArchE_Port_DoubleCreek','opsDIV_ArchE_port','BOTH_SAME','SERIES','ops_2022_DoubleCreek','*Map Frame*', 'series_title_DIV', 'OPS - ', 'DIV_24_', opsbasemaps,'Topo_24,000'),
#('stratplanning_2022_DoubleCreek_ORWWF000400.aprx','Layout_36x48_Series','stratplanning_NAIP_ArchE_port','IMG','SERIES','stratplanning_2022_DoubleCreek','*Map Frame*', '', '','inc_', stratbasemaps, stratbasemaps),
('BAM_2022_DoubleCreek_ORWWF00400.aprx','BAM_ArchE_Port_DoubleCreek','BAM_ArchE_port','IMG','SERIES','BAM_2022_DoubleCreek','*Map Frame*', '', '','inc_100_', '', ''),







##DO NOT DELETE THIS ]!##
]
########################################################################

### INCIDENT ###########################################################
inc_name = 'DoubleCreek'
idn = 'ORWWF00400'
inc_year = '2022'
########################################################################

### FOLDER LOCATIONS ###################################################
project_folder = r'C:\Users\Admin\FireNet\2022_ORWWF_DoubleCreek - 2022_DoubleCreek\projects'
#export_folder = r'C:\Users\Admin\Desktop'
export_folder = r'C:\Users\Admin\FireNet\2022_ORWWF_DoubleCreek - 2022_DoubleCreek\products'

########################################################################

### SETTINGS ###########################################################
beep = 'PDF'  #Play sound file ... 'PDF': after each PDF; 'END': only at the end; no beeps: ''
date_for_override = ''  # Set this to 'MMDD' to override the date for variable

###Prefix/Suffix in case you need them. Use '' if not needed############
print_prefix = ''
print_suffix = '_img'
georef_prefix = ''
georef_suffix = '_geo'
geoimg_prefix = ''
geoimg_suffix = '_geoimg'

dpi = 200
img_qual = 'NORMAL'
compress_vect = True
image_comp = 'ADAPTIVE'
embed_fonts = True
layer_attr = 'NONE'
jpg_comp = 50
clip = False
embed_color = True
########################################################################





###DO NOT EDIT#########################################################
#Sets the date that will be used in the folder name and output product name based on the time the script is run, and whether there
#is an override value in settings.
if date_for_override == '':
    # If there is not an override date, and it's between 0000 and 1700, date products for today
    if datetime.datetime.now().time() >= datetime.time(0) and datetime.datetime.now().time() < datetime.time(17):
        date_for = datetime.date.today().strftime("%m%d")
        date_dir = datetime.date.today().strftime("%m%d")
    else:
        # If there is not an override date, and it's between 1700 and 2400, date products for tomorrow
        date_for = (datetime.date.today()+datetime.timedelta(days=1)).strftime("%m%d")
        date_dir = (datetime.date.today()+datetime.timedelta(days=1)).strftime("%m%d")
else:
    date_for = date_for_override
export_folder = os.path.join(export_folder,inc_year+date_dir)

#Makes a date folder in the incident product folder, unless there is one already created.
try:
    os.makedirs(export_folder)
    print('{0} created...'.format(export_folder))
except FileExistsError:
    print('{0} exists...'.format(export_folder))

print('SAVE AND CLOSE ALL PROJECTS BEFORE EXPORTING!')
print('Exporting for {0}'.format(inc_year+date_for))

def randomSound():
    pathlocal = os.path.dirname(sys.argv[0])
    soundFolder = os.path.join(os.path.abspath(pathlocal),'Sounds')
    soundList = os.listdir('Sounds')
    soundLen = len(soundList)
    num = random.randrange(soundLen)
    soundPlay = soundFolder + '\\' + soundList[num]
    winsound.PlaySound(soundPlay, winsound.SND_FILENAME)
    pdfcount+=1


#PDF export for single maps
def export_pdf(prefix, suffix, georef, image_out):
    lyt.exportToPDF(os.path.join(export_folder,prefix+full_name+suffix+'.pdf'), dpi, img_qual, compress_vect, image_comp, embed_fonts, layer_attr, georef, jpg_comp, clip, image_out, embed_color)
    try:
        if beep=='PDF':
            randomSound()
    except Exception:
        pass

#For turning layout elements on and off based on the groupwildcard and uniquewildcard
def element_visibility(element_type, groupwildcard, uniquewildcard):
    try:
        for element in lyt.listElements(element_type, groupwildcard):
            element.visible = False
    except IndexError:
        print('No {0} with {1} in it\'s name'.format(element.type,groupwildcard))
        pass
    try:
        print('      ' + element_type + ' visible:')
        for element in lyt.listElements(element_type, uniquewildcard):
            element.visible = True
            print('      ' + element.name)
    except IndexError:
        print('No {0} with {1} in it\'s name'.format(element.type,uniquewildcard))
        pass

#Toggles visibility of all basemap layers in the baselist off, then toggles the desired base map visibility on
def toggle_layers(baselist, vislyrwildcard, job):
    m = aprx.listMaps(job[5])[0]
    for lyr in baselist:
        baselyr = m.listLayers(lyr)[0]
        baselyr.visible = False
    baselyrvis = m.listLayers(vislyrwildcard)[0]
    baselyrvis.visible = True
    global suffixbase
    suffixbase = baselyrvis.name


#PDF export for series maps. It iterates through the bookmarks grabbed by the wilcard defined in the
#export settings. It zooms to each bookmark's extent and alters the layout elements to match.
def export_series(prefix, suffix, georef, image_out, job):
    global pdfcount
    m = aprx.listMaps(job[5])[0]
    mf = lyt.listElements('MAPFRAME_ELEMENT', job[6])[0]
    for bk in m.listBookmarks(job[9] + '*'):
        bkname = str(bk.name)
        bknamesplit = bkname.split('_')
        exportgroupwild = bknamesplit[0]
        basemapwild = bknamesplit[1]
        uniquewild = bknamesplit[2]

        print('****    Exporting {0}...     ****'.format(uniquewild))
        mf.zoomToBookmark(bk)


        #If a title was identified in the export settings it will append the unique id to the title text
        if job[7] != '':
            title = lyt.listElements('TEXT_ELEMENT',job[7])[0]
            title.text = job[8] + uniquewild
        else:
            pass


        #If the bookmark name ends with G then the layout elements will be filtered based on
        #the exportgroup wildcard.  This is useful for DIV tiles that have all unique names.
        try:
            if bknamesplit[3] == 'G':
                uniquewild = exportgroupwild
        except Exception:
            pass


        #Turns off visibilty for all layout elements that begin with series.  Turns on all layout
        #elements that have the bookmark's unique wilcard in their name.  Then removes the spaces
        #from the unique wildcard so it can be appended to the export file name.
        element_visibility('TEXT_ELEMENT', 'series*', '*' + uniquewild + '*')
        element_visibility('PICTURE_ELEMENT', 'series*', '*' + uniquewild + '*')
        element_visibility('GRAPHIC_ELEMENT', 'series*', '*' + uniquewild + '*')
        uniquewildtemp = bknamesplit[2]
        uniquewild = uniquewildtemp.replace(' ','_')

        #Edits the output name so that the uniquewild from the bookmark is inserted following the first _
        arr = full_name.split('_')
        arr.insert(1, uniquewild)
        new_full_name = '_'.join(arr)

        #Exports a PDF. If a list of basemaps was used it appends the basemapwild to the end of the file name.
        if type(job[11]) == list:
            lyt.exportToPDF(os.path.join(export_folder,prefix+new_full_name+suffix+f'_{basemapwild}.pdf'), dpi, img_qual, compress_vect, image_comp, embed_fonts, layer_attr, georef, jpg_comp, clip, image_out, embed_color)
            print('\n$$$$$$    {0} exported    $$$$$$\n'.format(prefix+new_full_name+suffix+f'_{basemapwild}.pdf'))
        else:
            lyt.exportToPDF(os.path.join(export_folder,prefix+new_full_name+suffix+f'.pdf'), dpi, img_qual, compress_vect, image_comp, embed_fonts, layer_attr, georef, jpg_comp, clip, image_out, embed_color)
            print('\n$$$$$$    {0} exported    $$$$$$\n'.format(prefix+new_full_name+suffix+'.pdf'))

        try:
            if beep=='PDF':
                randomSound()
        except Exception:
            pass


#Checks if an export job is a series or single map
def checkseries(prefix, suffix, job, georef, image_out):
    if job[4] == 'SERIES':
        print(f'Basemap: {suffixbase}')
        export_series(prefix, suffix, georef, image_out, job)
    if job[4] == 'SINGLE':
        export_pdf(prefix, suffix, georef, image_out)


#Checks what type of PDF was requested (img, geo) and applies approriate settings
def export_type(job):
    imagelist = ['IMG','BOTH_SEP']
    geolist = ['GEOREF','BOTH_SEP']
    samelist = ['BOTH_SAME']
    typelist = [imagelist, geolist, samelist]
    for alist in typelist:
        if job[3] in alist:
            alistname = alist[0]
            nowDT=datetime.datetime.now()
            if alistname == 'IMG':
                print('Exporting {0} as image...'.format(full_name))
                checkseries(print_prefix, print_suffix, job, False, True)
            elif alistname == 'GEOREF':
                print('Exporting georeferenced {0}...'.format(full_name))
                checkseries(georef_prefix, georef_suffix, job, True, False)
            elif alistname == 'BOTH_SAME':
                print('Exporting georeferenced image {0}...'.format(full_name))
                checkseries(geoimg_prefix, geoimg_suffix, job, True, True)
            doneDT=datetime.datetime.now()
            if beep=='PDF':
                winsound.Beep(1500, 500)
            print('Completed in {0}'.format(doneDT-nowDT))

pdfcount=0

def worker(job):
    global pdfcount
    global full_name
    global aprx
    global lyt

    #If the export job doesn't have a user defined output name then it will use the layout file name.
    if job[2] == '':
        prod_name = job[1]
    else:
        prod_name = job[2]

    #Define variables
    full_name = prod_name+'_'+currentDT.strftime('%Y%m%d')+'_'+currentDT.strftime('%H%M')+'_'+inc_name+'_'+idn+'_'+date_for
    aprx = arcpy.mp.ArcGISProject(os.path.join(project_folder,job[0]))
    lyt = aprx.listLayouts(job[1])[0]

    #If a basemap list is defined in export settings then the toggle_layers function is run
    try:
        job[10] != ''
    except Exception:
        export_type(job)
        pdfcount+=1
    else:
        if type(job[11]) == list:
            for desiredbasemap in job[11]:
                toggle_layers(job[10], desiredbasemap, job)
                export_type(export_job)
        elif type(job[11]) == str:
            toggle_layers(job[10], job[11], job)
            export_type(job)


#Iterates throught the export jobs listed at the top of the script.
for export_job in export_jobs:
    worker(export_job)


finishDT = datetime.datetime.now()
deltaDT = finishDT-currentDT
print('{0} PDFs exported in {1}'.format(pdfcount,deltaDT))

try:
    if beep=='END':
        randomSound()
except Exception:
    pass



#Fun Fact Generator
pathlocal = os.path.dirname(sys.argv[0])
funfacts = os.path.join(os.path.abspath(pathlocal),'FunFacts.csv')
num = random.randrange(312)

def wrap_by_word(s, n):
    '''returns a string where \\n is inserted between every n words'''
    a = s.split()
    ret = ''
    for i in range(0, len(a), n):
        ret += ' '.join(a[i:i+n]) + '\n'
    return ret

try:
    with open(funfacts, mode='r', errors='ignore') as csv_file:
        csv_reader = csv.DictReader(csv_file, dialect = csv.excel)
        for row in csv_reader:
            if row['ID'] == str(num):
                fact = row['title']
                factprint = wrap_by_word(fact.replace('TIL','Did you know:'),12)
                source = str(row['wikipedia'])
                final = f'''{factprint}\nSource: {source}'''

                print('''\n#####Fun Fact######\n''' + final)
except:
    print('Why did you get rid of the fun fact .csv :(')
### DONE ###############################################################


