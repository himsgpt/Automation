# -*- coding: utf-8 -*-
"""
This is a temporary script file.
"""

import win32com.client
# import time
import logging
import os
os.getcwd()
import Config.global_var as gvar

# =============================================================================
# logging definition
# =============================================================================
logger = logging.getLogger(__name__)
logger.setLevel(logging.DEBUG)
logging.basicConfig(filename=gvar.LOG_FILE,
                    format='%(asctime)s -- %(levelname)s -- %(message)s', 
                    level = logging.DEBUG,datefmt='%Y-%m-%d %H:%M:%S')

# =============================================================================
# Defining new directories
# =============================================================================
def dir_creation(Pathname,newname):
    newdir = os.path.join(Pathname,newname)
    if not os.path.exists(newdir):
        os.makedirs(newdir)
        print("new directory created")
    else: 
        print("directory already exist")
    return newdir

# =============================================================================
# Refreshing function
# =============================================================================

def refresh_dasboards(Pathname,filename,newdir,newfile):
    '''
    Refreshing the called file and making amends to the worksheets
    ----------
    Returns none but updates and saves a copy of the file
    -------
    '''
    xlapp = win32com.client.DispatchEx("Excel.Application")
    # Show Excel. While this is not required, it can help with debugging
    xlapp.DisplayAlerts = True
    xlapp.Visible = True
    wb = xlapp.Workbooks.Open(Pathname+ '/'+ filename + '.xlsx')
    wb.RefreshAll()
    # time.sleep(5)
    # this will actually wait for the excel workbook to finish updating
    xlapp.CalculateUntilAsyncQueriesDone()
    try:
        ws = wb.Worksheets('ReadMe')
        ws.Range('C2').value = gvar.newperiod
        ws.Range('C4').value = gvar.week1
        ws.Range('C5').value = gvar.week2
    except:
        print("No Readme present")
    wb.SaveCopyAs(newdir+'/'+newfile+'.xlsx')
    wb.Save()
    # wb.close()
    xlapp.Quit()
    del wb
    del xlapp

    
# =============================================================================
# Running Author, Authen & Error dashboards for all acquirers
# =============================================================================

for i in (gvar.acquirers):
    if (gvar.Author_ref=='Y'):
        print("******** Refreshing Author *************** ")
        filename = i+"_Author_Dashboard"
        newfile = i+"_Author_Dashboard_"+gvar.newperiod 
        newdir = dir_creation(gvar.AuthorPathName,gvar.newperiod)
        print("Refreshing-----> ",filename)
        try:
            refresh_dasboards(gvar.AuthorPathName,filename,newdir,newfile)
            logger.info("Refreshed Author file " + newfile)
        except Exception as e:
            print("Unable to refresh Author file ")
            logger.critical("Unable to refresh Author file " + filename + '\n' + str(e))
    
    if (gvar.Error_ref=='Y'):
        print("******** Refreshing Error *************** ")
        filename = i+"_Error_View"
        newfile = i+"_Error_View_"+gvar.newperiod
        newdir = dir_creation(gvar.ErrorPathName,gvar.newperiod)
        print("Refreshing-----> ",filename)
        try:
            refresh_dasboards(gvar.ErrorPathName,filename,newdir,newfile)
            logger.info("Refreshed error file " + newfile)
        except Exception as e:
            print("Unable to refresh Error file ")
            logger.critical("Unable to refresh Error file " + filename + '\n'+str(e))
    
    if (gvar.Authen_ref=='Y'):
        print("******** Refreshing Authen *************** ")
        filename = i+"_Authen_Dashboard"
        newfile = i+"_Authen_Dashboard_"+gvar.newperiod
        newdir = dir_creation(gvar.AuthenPathName,gvar.newperiod)
        print("Refreshing-----> ",filename)
        try:
            refresh_dasboards(gvar.AuthenPathName,filename,newdir,newfile)
            logger.info("Refreshed Authen file " + newfile)
        except Exception as e:
            print("Unable to refresh Authen file ")
            logger.critical("Unable to refresh Authen file " + filename +'\n'+ str(e))
            
    if (gvar.Samples_ref=='Y'):
        print("******** Refreshing Samples *************** ")
        filename = i+"_Error_Sample"
        newfile = i+"_Error_Sample_"+gvar.newperiod
        newdir = dir_creation(gvar.SamplesPathName,gvar.newperiod)
        print("Refreshing-----> ",filename)
        try:
            refresh_dasboards(gvar.SamplesPathName,filename,newdir,newfile)
            logger.info("Refreshed Sample file " + newfile)
        except Exception as e:
            print("Unable to refresh Sample file ")
            logger.critical("Unable to refresh Sample file " + filename +'\n'+ str(e))
            
    # else:
    #     logger.critical("No refresh done")
    

