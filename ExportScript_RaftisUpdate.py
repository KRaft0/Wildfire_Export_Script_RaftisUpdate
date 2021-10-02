#********************************************************************************************************************
# Name: Wildfire Incident Export Script            
# Created: 09/30/2021          
# Author: Kaylene Raftis, Based heavily on export script written by Paul Hoefler
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
export_jobs = [
########################################################################


### SET ITEMS TO EXPORT BELOW ##########################################
# export_jobs = [('APRX','LAYOUT','OutputNAME','[PRINT|GEOREF|BOTH]'],[SINGLE|SERIES],'MAP_NAME','MAPFRAME_ELEMENT_NAME','TITLE_ELEMENT_NAME','TITLE_TEXT'),]
#                                                                     **Only fill out 'MAP_NAME','MAPFRAME_ELEMENT_NAME','TITLE_ELEMENT_NAME','TITLE_TEXT'if the map is a series**

#****Single*******
#('pilot_2021_DevilsKnob_ORUPF000450_ArcPro_2_7.aprx','pilot_18x24_port_DevilsKnob_ORUPF000450','pilot_18x24_port_','PRINT','SINGLE'),
('progression_PanunTools_2021_DevilsKnob_ORUPF000450_ArcPro_2_7.aprx','prog_36x48_port_DevilsKnobComplex_ORUPF000450','progression_port_','GEOREF','SINGLE'),


#****Series*******
('DIV_Tiles_2021_DevilsKnob_ORUPF000450_ArcPro_2_7.aprx','Ops_DIV_Map_Packet','Ops_DIV_Map_Packet_18x24_land','PRINT','SERIES','Ops_DivTiles_2021_DevilsKnob','*Main*','Title','OPS - '),
#('Repair_2021_DevilsKnob_ORUPF000450_ArcPro_2_7.aprx','Repair_Tiles','SuppressionRepair_Tiles_ArchC_land','GEOREF','SERIES','RepairTiles_2021_DevilsKnob','*Main*','Title','Suppression Repair '),


##DO NOT DELETE THIS ]!##
]  
########################################################################

### INCIDENT ###########################################################
inc_name = 'DevilsKnob'
idn = 'ORUPF000450'
inc_year = '2021'
########################################################################

### FOLDER LOCATIONS ###################################################
project_folder = r'C:\Users\Admin\FireNet\2021_ORUPF_DevilsKnob - 2021_DevilsKnob\projects'
export_folder = r'C:\Users\Admin\FireNet\2021_ORUPF_DevilsKnob - 2021_DevilsKnob\products'
########################################################################

### SETTINGS ###########################################################
beep = 'END'  #Play sound file ... 'PDF': after each PDF; 'END': only at the end; no beeps: ''
date_for_override = ''  # Set this to 'MMDD' to override the date for variable

###Prefix/Suffix in case you need them. Use '' if not needed############
print_prefix = ''
print_suffix = '_img'
georef_prefix = ''
georef_suffix = '_geo'

dpi = 150
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
try:
  #Your variables, environment settings, and entire script go into the
  #try statement for error testing.  If an error is found you drop to the
  #except statement and are given a complete error message.

    #Sets the date that will be used in the folder name and output product name based on the time the script is run, and whether there
    #is an override value in settings.
    if date_for_override == '':
        # If there is not an override date, and it's between 0000 and 1700, date products for today
        if datetime.datetime.now().time() >= datetime.time(0) and datetime.datetime.now().time() < datetime.time(12):
            date_for = datetime.date.today().strftime("%m%d")
            date_dir = datetime.date.today().strftime("%m%d")
        elif datetime.datetime.now().time() >= datetime.time(12) and datetime.datetime.now().time() < datetime.time(17):
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
        print("{0} created...".format(export_folder))
    except FileExistsError:
        print("{0} exists...".format(export_folder))

    print("SAVE AND CLOSE ALL PROJECTS BEFORE EXPORTING!")
    print("Exporting for {0}".format(inc_year+date_for))
    
    #Create the two functions that export based on whether the layout is a series or not.
    def export_pdf(prefix, suffix, georef, image_out):
        lyt.exportToPDF(os.path.join(export_folder,prefix+full_name+suffix+".pdf"), dpi, img_qual, compress_vect, image_comp, embed_fonts, layer_attr, georef, jpg_comp, clip, image_out, embed_color)
    
    def export_series(prefix, suffix, georef, image_out):
        for bk in m.listBookmarks():
            mf.zoomToBookmark(bk)
            title.text = export_job[8] + str(bk.name)
            lyt.exportToPDF(os.path.join(export_folder,prefix+full_name+suffix+f'_{bk.name}.pdf'), dpi, img_qual, compress_vect, image_comp, embed_fonts, layer_attr, georef, jpg_comp, clip, image_out, embed_color)

    pdfcount=0

    #Iterates throught the export jobs listed at the top of the script.  
    for export_job in export_jobs:
        #If the export job doesn't have a user defined output name then it will use the layout file name.
        if export_job[2] == '':
            prod_name = export_job[1]
        else:
            prod_name = export_job[2]
            
        #Define variables    
        full_name = prod_name+"_"+currentDT.strftime("%Y%m%d")+"_"+currentDT.strftime("%H%M")+"_"+inc_name+"_"+idn+"_"+date_for
        aprx = arcpy.mp.ArcGISProject(os.path.join(project_folder,export_job[0]))
        lyt = aprx.listLayouts(export_job[1])[0]
        
        #First sorts out the export jobs based on print/georef/both.  Then directs to the appropriate function based on 
        #whether the export job is a single map or series.
        if export_job[3] == 'PRINT' or export_job[3] == 'BOTH':
            nowDT=datetime.datetime.now()
            print("Exporting {0} for print...".format(full_name))
            if export_job[4] == 'SERIES':
                m = aprx.listMaps(export_job[5])[0]
                mf = lyt.listElements('MAPFRAME_ELEMENT',export_job[6])[0]
                title = lyt.listElements('TEXT_ELEMENT',export_job[7])[0]
                export_series(print_prefix, print_suffix, False, True)
            if export_job[4] == 'SINGLE':
                export_pdf(print_prefix, print_suffix, False, True)
            doneDT=datetime.datetime.now()
            if beep=='PDF':
                winsound.Beep(1500, 500)
            print('Completed in {0}'.format(doneDT-nowDT))
            pdfcount+=1
        if export_job[3] == 'GEOREF' or export_job[3] == 'BOTH':
            nowDT=datetime.datetime.now()
            print("Exporting georeferenced {0}...".format(full_name))
            if export_job[4] == 'SERIES':
                m = aprx.listMaps(export_job[5])[0]
                mf = lyt.listElements('MAPFRAME_ELEMENT',export_job[6])[0]
                title = lyt.listElements('TEXT_ELEMENT',export_job[7])[0]
                export_series(georef_prefix, georef_suffix, True, False)
            if export_job[4] == 'SINGLE':
                export_pdf(georef_prefix, georef_suffix, True, False)
            doneDT=datetime.datetime.now()
            if beep=='PDF':
                winsound.Beep(1500, 500)
            print('Completed in {0}'.format(doneDT-nowDT))
            pdfcount+=1

    finishDT = datetime.datetime.now()
    deltaDT = finishDT-currentDT
    if beep=='END':
        winsound.PlaySound('meow.wav', winsound.SND_FILENAME)
        #winsound.Beep(1500, 500)
    print('{0} PDFs exported in {1}'.format(pdfcount,deltaDT))
    
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
        with open(funfacts, mode='r', errors="ignore") as csv_file:
            csv_reader = csv.DictReader(csv_file, dialect = csv.excel)
            for row in csv_reader:
                if row['ID'] == str(num):
                    fact = row['title']
                    factprint = wrap_by_word(fact.replace('TIL','Did you know'),12)
                    source = str(row['wikipedia'])
                    final = f'''{factprint}\nSource: {source}'''
                    
                    print('''\n#####Fun Fact######\n''' + final)
    except:
        print('Why did you get rid of the fun fact .csv :(')
### DONE ###############################################################


except:
#imported modules produce error messaging from both ArcPy and Python.
  import sys, traceback #if needed
  # Get the traceback object
  tb = sys.exc_info()[2]
  tbinfo = traceback.format_tb(tb)[0]
  # Concatenate information together concerning the error into a message string
  pymsg = "PYTHON ERRORS:\nTraceback info:\n" + tbinfo + "\nError Info:\n" + str(sys.exc_info()[1])
  msgs = "ArcPy ERRORS:\n" + arcpy.GetMessages(2) + "\n"
  # Return python error messages for use in script tool or Python Window
  arcpy.AddError(pymsg)
  arcpy.AddError(msgs)
  # Print Python error messages for use in Python / Python Window
  print (pymsg + "\n")
  print (msgs)

