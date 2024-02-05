This is a project I did for my company, Energizer Holdings, where we perform tests on hundreds of thousands of batteries per year. These batteries get labeled, then sorted into boxes grouped by test category. Writing the necessary information on every one of these boxes is tedious, time-consuming, and prone to errors. I wrote this macro and python script to scrape all the data on the submission form, clean the data, duplicate the rows depending on how many batteries can fit in the box, and print the information out onto generated labels. 






INSTRUCTIONS
This program intends to greatly reduce the repetition in labeling the boxes of newly sorted cells and to reduce errors. 
-To start, copy the entire label_macro.txt file. 
-Open the submission form you want to print labels for. If you have the "Developer" tab in excel, click on that. 
	-If not, right-click on the ribbon ("File, Home, Insert..."), click "Customize Ribbon". 
	-On the right side of the window, scroll down to find the "Developer" box and click it. Hit okay, then click on the Developer tab. 
-Click the left-most button labeled "Visual Basic". 
-Find the "File" tab at the top left of the window and click the down-arrow right below it. Then, click "Module". 
-You should get a blank page. Paste the text of label_macro.txt into here. Then click the "play" button at the top. 
-A window will pop up. Select "Create_Lables" then click "Run"


STRUCTURE
When you run this macro, it will scrape data off of the cover page, then create a separate CSV file for each test grouping (performance, safety, etc). 
These CSVs will save to a designated file location. The macro will then call a python script. This python script loads in all of these CSVs and reformats
everything into a final CSV that can easily be loaded into the DYMO software and printed off as a batch of labels. This final CSV will save to the same location
as the other CSVs, titled FINISHED_LABELS.csv


IMPORTANT
-the macro text looks for the python file in a certain location. If you move the script, you must also change the filepath in the macro (found near the top).
-If you want to move the location of the project, you must change the destination of the individual CSVs, found near the top of the macro. 
	-You must then change the filepath locations in the PYTHON script, as this script looks for these CSVs. 
-If you want the final CSV to end up in a different place, that can be changed in the python script, also found near the top. 


DEPENDENCIES
For this to work, the computer must have python installed. It must also have pandas and numpy. 
	-if python is already on there, install the other two with "pip install pandas" and "pip install numpy", written into the powershell command line. 
