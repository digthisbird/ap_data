# ap_data
libraries: 
pandas
re
argparse


This program takes a csv at the command line and checks the data against prescribed standards. Requires precise headers to run. 

Headers:
Course	/ ItemSequence	/ Intended Form /	Form Deadline	/ Batch	/ Item Purpose /	Item Type /	Section  /	
"Topic (Sequence)" / "Topic (Label)"	/ Skill	/ Subskill	/ Complexity	/ Author Name	/ Author ID	/ 
Date Assigned	/ Date Due	/ Date Submitted	/ Content Reviewer	/ Graphics Status	/ Item Status	/ "Notes / Comments"

Uses pandas xlsx writer to output a new workbook with results in multiple sheets.
