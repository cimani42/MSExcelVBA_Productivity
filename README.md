# MSExcelVBA_Productivity
Writing Code in MS Excel VBA to help productivity with repetitive tasks conducted whilst at work.


Working in an Accounting/Finance envrionment, I often use MS Excel as a tool to facilitate the work I do. 
Although I have preset templates - Workbooks - created to help work efficiently, I have began to explore manupulating Macros and writing code for repetitive tasks. Reducing the time spent on the task. 

In this commit, I have written three functions; all of which contribute to creating a new directory - for a given file location - saved as the current date.
The Macros combine using in-built fuctions given in MS Excel, as well as user input from the Active Worksheet, and cell, selected. 
(To prevent giving away sensitive information, default locations have been written)


**Sub CreateFolder()**
A simple macro that calls the other written macros in the indended order. 
 _This macro was used to practice calling macros outside of the intended scope._
 _Moreover, this macro is useful as the **sub CurrentPeriod** will be used in multiple MS Workbooks._
 
 **Sub CurrentPeriod()**
 Accounting and Finance functions split thier periods into relevant time periods within, and beyond, their fiscal period. This macro uses If, ElseIf and Else syntax
 with bool conditionals to return what financial period the current day falls within. The code will output a string to a specific Worksheet and Cell due to the   templates.
  _Many of Workbooks require this feature. Either being daily or monnthly. Due to this, is the reason why I have broken the macros into three._
  _There were no issues in creating this macro. Just making sure to be correct with the entry and consisten throughout. Testing for bugs, or unintentional outputs helped
  me with this._
 _As I am new to wrting in VBA, findind documenation for the syntax was/is the hardest, as I tried to understand the best method to input data. Using the Worksheet or 
 functions already exisiting in MS Excel. For example, initially the date was referenced in a given cell. However, I have opted to use the Date function, having 
 read documentation on this. 
 
 **Sub MakeDateDirectory()**
 This is code written to create a new directory, saving to a specific file path. The intention is for this to be dynamic so that it can save over each period in the
 fiscal year. The new directory created will be saved as the current date. Where the file has already been created. A duplicate should not be recreated. 
  _Finding documenation to create the formatting to align with how the folders currently exists was difficult. However, I managed to find out how to set the format
  having to refactro this later on. (A concept I understood through my use of python).
  The main trouble with this macro was adding the variables to the folder path. I would follow the accepted convention outlined in the documetation using "", & and
  \. I kept on receieving the same error message. Although I did not know why? I ended up deleting the line and typing it manually, opposed to manipulating the 
  copied address. This resovled the issue, so I am still unaware._
  
  _For this particular code, I could improve on it by ensuring that files with the same date but different formatting are not created._
  
