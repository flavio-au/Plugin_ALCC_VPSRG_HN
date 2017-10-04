# Plugin_ALCC_VPSRG_HN
Eclipse plug-in for harvesting plan metrics for H&amp;N VicRapidPlan project

This is a plug-in running on Eclipse, where you need to have a patient open.
Just copy the plug-in file (Plugin_ALCC_VPSRG_HN.esapi.dll) anywhere, and lunch from Eclipse  
tools>script… (select directory where  script is)
The script searches for the structures appearing on the excel worksheet (VPSRG Head and Neck case-tracking sheet.xls) 
with logic intended to be as broad as possible but anyhow directed to pick the correct choices at ALCC.
When more than 1 structure is found matching the criteria, it pops-up a dialog for choosing the correct one.
For the intermediate and low PTVs it searches their corresponding dose levels on their names: 
PTV 56 gets its 100% dose level = 56 Gy.
At the end it produces (overwrites) a file VPSRG_HN_[pat ID].txt at c:\temp
For inserting the data in excel, the excel workbook needs to be saved with macro enable ( .xlsm) and the visual basic module 
imported into it (needs “developer” main tab enabled on options>customize ribbon).
developer>visual basic> (have the workbook selected on context window) > import file (select the Import_VPSRG_HN.bas file)
Before running the macro select the FIRST cell (Col. A) of the row were you want the data imported, 
then you have to run the macro (Alt + F8 shows the macros) “Import_VPSRG_HN” and there it is!

Hope it will be useful!
