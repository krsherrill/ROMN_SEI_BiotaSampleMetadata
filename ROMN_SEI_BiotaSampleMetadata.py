###################################
# ROMN_SEI_BiotaSampleMetadata.py
###################################
# Description:
# Script creates the Biota metadata information for the Rhithron Biota Lab.  After all the Streams field sheet information has been entered in the Stream database this script is
# used to run the routine to define the biota sample metadata summary. This Sample Metadata summary is used by the Rhithron lab to define relevant metadata content and is used in lab calculations.
# The lab will be unable to process biota data until this information is received.  Sample metadata is created for the MacroInverts, Periphyton, Chlorophyll, and Ash Free Dry Mass biota components.

#Code performs the Following Routines:

#1) Excecutes the 'dbo.BioSampleMetadata_RunAll' stored procedure in the Streams SQL Database. Procedure creates the following metadata tables in SQL Server:
#dbo.SampleMetadataAFDM, dbo.SampleMetadataBenthos, dbo.SampleMetadataChloraA and dbo.SampleMetadataPeriphyton.<br>
#2) Exports the above table records by defined year to a defined output directory and defined excel file output name

# Dependencies:
# Python version 3.9
# Packages: Pandas and pyodbc
# Script Name: ROMN_SEI_BiotaSampleMetadata.py

# Python/Conda environment - py39
# Created by: Created by Kirk Sherrill - Data Manager Rock Mountain Network - I&M National Park Service
# Date Created: November 30th, 2021

#######################################
# Start of Parameters requiring set up.
#######################################

#Define Field Season Being Processed
inYear = 2022

#ROMN Streams SQL Server path - leaving null for security
streamsServer = "xxxxxxx\\yyy"
#Streams Database Name
streamsDB = "ROMN_SEI"

#Directory Information
outputFolder = r'C:\ROMN\Monitoring\Streams\Data\DataGathering\Lab\Rhithron\Python\SampleMetadata\2022' #Folder for the output Data Package Products
workspace = r'C:\ROMN\Monitoring\Streams\Data\DataGathering\Lab\Rhithron\Python\SampleMetadata\2022' #Workspace Folder

outName = "NPS_ROMN_SEI_" + str(inYear) + "_BioSampleMetadata"   #output name for export excel file - note yyyymmdd will alos be added as suffix

#######################################
## Below are paths which are hard coded:
#######################################
#Import Required Libraries

import pyodbc
import time
import pandas as pd
import sys
#from datetime import date
#import shutil
import os
#from zipfile import ZipFile
import traceback
from openpyxl import load_workbook
##################################


#################################################

# Function to Get the Date/Time
def timeFun():
    from datetime import datetime
    b = datetime.now()
    messageTime = b.isoformat()
    return messageTime

def main():
    try:
        # Run the [summary].[EventSiteMetadataAll2nd] stored procedure in ROMN_SEI - SQL Server
        # This stored procedure updates the 'summary.EventSiteMetadata' table used to define Site/Event Metadata
        outVal = runStoredProcedure("[summary].[EventSiteMetadataAll2nd]")
        if outVal == "Success function":
            print("Success - Function runStoredProcedure - [summary].[EventSiteMetadataAll2nd]")
        else:
            print(
                "WARNING - Function runStoredProcedure - [summary].[EventSiteMetadataAll2nd] - failed - Exiting Script")
            exit()

        # Run the [dbo].[BioSampleMetadata_RunAll] stored procedure in ROMN_SEI - SQL Server
        outVal = runStoredProcedure("[dbo].[BioSampleMetadata_RunAll]")
        if outVal == "Success function":
            print("Success - Function runStoredProcedure - [dbo].[BioSampleMetadata_RunAll]")
        else:
            print("WARNING - Function runStoredProcedure - [dbo].[BioSampleMetadata_RunAll] - failed - Exiting Script")
            exit()

        # Make output Directory if it doesn't exist
        if os.path.exists(outputFolder):
            print("Directory Exists:" + outputFolder)
        else:
            os.mkdir(outputFolder)
            print("Created Directory:" + outputFolder)

            # Define Tables to be exported
        processList = ['AFDM', 'Benthos', 'ChlorA', 'Periphyton']
        rowRange = range(0, len(processList))

        for row in rowRange:

            inTable = str(processList[row])

            # Define Query
            outVal = defineQueryFun(inTable, inYear)
            if outVal[0] != "Success function":
                print("WARNING - Function defineQueryFun failed - for - " + inTable + " - Exiting Script")
            else:
                print("Success - Function defineQueryFun - for -" + inTable)
                outQuery = outVal[1]

            outVal = connect_to_SSMS(outQuery)
            if outVal[0] != "Success function":
                print("WARNING - Function connect_to_SSMS failed - for - " + inTable + " - Exiting Script")
            else:
                print("Success - Function connect_to_SSMS - for - " + inTable)
                outDf = outVal[1]

                # Export to Excel
            outFile = outputFolder + "\\" + outName + "_" + time.strftime("%Y%m%d") + ".xlsx"

            if row == 0:  # Create Export Excel
                with pd.ExcelWriter(outFile, engine='openpyxl') as writer:
                    outDf.to_excel(writer, index=False, sheet_name=inTable)
                    # writer.close()
                    # writer.save()
            else:  # Export to Existings Excel and then add to new worksheet

                book = load_workbook(outFile)
                writer = pd.ExcelWriter(outFile, engine='openpyxl')
                writer.book = book

                outDf.to_excel(writer, index=False, sheet_name=inTable)
                writer.save()
                writer.close()

        messageTime = timeFun()
        scriptMsg = "Successfully processed SEI Biota Metadata - ROMN_SEI_BiotaSampleMetadata.py - " + messageTime
        print(scriptMsg)
        logFile = open(logFileName, "a")
        logFile.write(scriptMsg + "\n")
        logFile.close()
        print("Successfully processed SEI Biota Metadata")
    except:

        messageTime = timeFun()
        scriptMsg = "WARNING ERROR processing SEI Biota Metadata - ROMN_SEI_BiotaSampleMetadata.py - " + messageTime
        print(scriptMsg)
        logFile = open(logFileName, "a")
        logFile.write(scriptMsg + "\n")
        logFile.close()
        traceback.print_exc(file=sys.stdout)


# Function to connect SSMS database and export query to dataframe
# Return dataFrame
def connect_to_SSMS(query):
    import pyodbc
    # Connect to DB
    #conn = pyodbc.connect('Driver={SQL Server};Server=INP2300SQL01\\NTWK;Database=ROMN_SEI;Trusted_Connection=yes;')
    connectionString = 'Driver={SQL Server};Server=' + streamsServer + ';Database=' + streamsDB + ';Trusted_Connection=yes;'
    conn = pyodbc.connect(connectionString)
    # Perform Query and Export to Dataframe
    queryDf = pd.read_sql(query, conn)

    conn.close()

    # display(queryDf)
    return "Success function", queryDf
    # return queryDf


# Function to Run the Stored Procedure  - "dbo.BioSampleMetadata_RunAll"
# input: inStoredProc = name of stored procedure to be Executed
def runStoredProcedure(inStoredProc):
    try:

        # Connect to SQL
        conn = pyodbc.connect('Driver={SQL Server};Server=INP2300SQL01\\NTWK;Database=ROMN_SEI;Trusted_Connection=yes;')

        # Define Stored Procedure
        sql = 'exec ' + inStoredProc + ''

        cursor = conn.cursor()
        cursor.execute(sql)
        cursor.close()
        conn.commit()

        return "Success function"

    except pyodbc.Error as err:
        sqlstate = err.args[0]
        print('Exception' + sqlstate)
        print("Error on runStoredProcedure Function")
        traceback.print_exc(file=sys.stdout)
        return "Failed function - 'runStoredProcedure'"

    finally:
        conn.close()


# Define the Query to pass to SQL Server
def defineQueryFun(inTable, inYear):
    try:

        strInYear = str(inYear)

        if inTable == "AFDM":
            outQuery = "SELECT [ProjectID], [SampleID],[BottleCode],[Sample_Station_Name],[Sample_STORET_Station_ID],[Sample_Client_ID]\
                ,[Event ID],[Sample_Date_Collected],[Sample_Number_Jars],[Sample_Notes],[Sample_Lat],[Sample_Lon],[PPHYTON_SamplerArea_cm2]\
                ,[PPHYTON_Number_Erosional],[PPHYTON_Number_Depositional],[PPHYTON_TotalCompVol_ml],[PPHYTON_IDVol_ml],[FiltVol_Chlra_ml]\
                ,[PPHYTON_FiltVol_AFDM_ml],[PPHYTON_TotalAreaSampled_cm2] FROM [ROMN_SEI].[dbo].[SampleMetadataAFDM] as DS INNER JOIN\
                summary.EventSiteMetadata as MD ON DS.[Event ID] = MD.EventName WHERE MD.Year >= " + strInYear + " ORDER BY DS.Sample_STORET_Station_ID,\
                DS.Sample_Date_Collected"

        elif inTable == "Benthos":
            outQuery = "SELECT [ProjectID],[SampleID],[BottleCode],[Sample_Station_Name],[Sample_STORET_Station_ID],[Event ID],[Sample_Client_ID]\
                ,[Sample_Date_Collected],[Sample_Number_Jars],[Sample_Notes],[Sample_Lat],[Sample_Lon],[Sample_Collection_Procedure_ID]\
                ,[BENTHOS_SampleType],[BENTHOS_MeshSize_um],[Bags Benthos Soft],[Bags Benthos Hard],[NumBugSubSamples],[BENTHOS_SamplerArea_m2]\
                ,[BENTHOS_TotalAreaSampled_m2] FROM [ROMN_SEI].[dbo].[SampleMetadataBenthos] as DS INNER JOIN summary.EventSiteMetadata\
                as MD ON DS.[Event ID] = MD.EventName WHERE MD.Year >= " + strInYear + " ORDER BY DS.Sample_STORET_Station_ID, DS.Sample_Date_Collected"

        elif inTable == "ChlorA":
            outQuery = "SELECT [ProjectID],[SampleID],[BottleCode],[Sample_Station_Name],[Sample_STORET_Station_ID],[Sample_Client_ID],[Event ID]\
                ,[Sample_Date_Collected],[Sample_Number_Jars],[Sample_Notes],[Sample_Lat],[Sample_Lon],[PPHYTON_SamplerArea_cm2],[PPHYTON_Number_Erosional]\
                ,[PPHYTON_Number_Depositional],[PPHYTON_TotalCompVol_ml],[PPHYTON_IDVol_ml],[FiltVol_Chlra_ml],[PPHYTON_FiltVol_AFDM_ml],[PPHYTON_TotalAreaSampled_cm2]\
                FROM [ROMN_SEI].[dbo].[SampleMetadataChlorA] as DS INNER JOIN summary.EventSiteMetadata as MD ON DS.[Event ID] = MD.EventName\
                WHERE MD.Year >= " + strInYear + " ORDER BY DS.Sample_STORET_Station_ID, DS.Sample_Date_Collected"

        elif inTable == "Periphyton":
            outQuery = "SELECT [ProjectID],[SampleID],[BottleCode],[Sample_Station_Name],[Sample_STORET_Station_ID],[Event ID],[Sample_Client_ID],[Sample_Date_Collected]\
                ,[Sample_Number_Jars],[Sample_Notes],[Sample_Lat],[Sample_Lon],[PPHYTON_SamplerArea_cm2],[PPHYTON_Number_Erosional],[PPHYTON_Number_Depositional]\
                ,[PPHYTON_TotalCompVol_ml],[PPHYTON_IDVol_ml],[FiltVol_Chlra_ml],[PPHYTON_FiltVol_AFDM_ml],[PPHYTON_TotalAreaSampled_cm2]\
                FROM [ROMN_SEI].[dbo].[SampleMetadataPeriphyton] as DS INNER JOIN summary.EventSiteMetadata as MD ON DS.[Event ID] = MD.EventName\
                WHERE MD.Year >= " + strInYear + " ORDER BY DS.Sample_STORET_Station_ID, DS.Sample_Date_Collected"

        else:
            print("In Table - '" + inTable + "'- not defiend unable to process")
            return "Failed Function - 'process_VerifyDatasets'"

        return "Success function", outQuery

    except:

        messageTime = timeFun()
        print("Error on defineQuery Function - ") + messageTime
        traceback.print_exc(file=sys.stdout)
        return "Failed function - 'defineQuery'"


if __name__ == '__main__':

    # Write parameters to log file ---------------------------------------------
    ##################################
    # Checking for working directories
    ##################################

    if os.path.exists(workspace):
        pass
    else:
        os.makedirs(workspace)

    logFileName = workspace + "\\logFile_" + outName + '.txt'

    # Check if logFile exists
    if os.path.exists(logFileName):
        pass
    else:
        logFile = open(logFileName, "w")  # Creating index file if it doesn't exist
        logFile.close()

    # Analyses routine ---------------------------------------------------------
    main()
