# Plugin_ALCC_VPSRG_HN
Eclipse plug-in for harvesting plan metrics for H&amp;N VicRapidPlan project

This is a plug-in running on Eclipse, where you need to have a patient open.
Just download the plug-in file (/bin/release/Plugin_ALCC_VPSRG_HN.esapi.dll) anywhere, and lunch from Eclipse  
tools>script… (select directory where script is)
The script searches for the structures appearing on the excel worksheet (VPSRG Head and Neck case-tracking sheet V3.xlsm)
with logic intended to be as broad as possible but anyhow directed to pick the correct choices at ALCC.
When more than 1 structure is found matching the criteria, it pops-up a dialog for choosing the correct one.
For the intermediate and low PTVs it searches their corresponding dose levels on their names: 
PTV 56 gets its 100% dose level = 56 Gy.
At the end it produces (overwrites) a file VPSRG_HN_[pat ID].txt at c:\temp whith the data 
and a file VPSRG_HN_[pat ID]_SelectedStructures.txt with a summary of wich structures where selected (for QA...)
The source code is available, please if you pretend to customize it on github, just fork the project.
For inserting the data in excel, select the any cell OF THE ROW were you want the data imported, 
then you have to run the macro (Alt + F8 shows the macros) “Import_VPSRG_HN” and there it is!
The macro saves the workbook before running (as undo is not available) just in case.
For comparing data with Vic constrains, just run macro compare_values.
The selection logic is presented on file LogicOfStrSelection.docx

Hope it will be useful!
