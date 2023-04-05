# ROMN_SEI_BiotaSampleMetadata.py
Rocky Mountain Network Stream Biota Metadata Script for Rhithron Lab. 

## Description
Script creates the Biota metadata information for the Rhithron Biota Lab.  After all the Streams field sheet information has been entered in the Stream database this script is used to run the routine to define the biota sample metadata summary. This Sample Metadata summary is used by the Rhithron lab to define relevant metadata content and is used in lab calculations. The lab will be unable to process biota data until this information is received.  Sample metadata is created for the MacroInverts, Periphyton, Chlorophyll, and Ash Free Dry Mass biota components.

## Routines
1) Excecutes the 'dbo.BioSampleMetadata_RunAll' stored procedure in the Streams SQL Database. Procedure creates the following metadata tables in SQL Server:
dbo.SampleMetadataAFDM, dbo.SampleMetadataBenthos, dbo.SampleMetadataChloraA and dbo.SampleMetadataPeriphyton.<br>
2) Exports the above table records by defined year to a defined output directory and defined excel file output name.

