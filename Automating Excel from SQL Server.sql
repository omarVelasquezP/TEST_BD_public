      
  
Automating Excel from SQL Server
By Wayne Sheffield, 2008/12/05 
Total article views: 23588 | Views in the last 30 days: 781 
  Rate this |   Join the discussion |   Briefcase 
You get a new email… from your boss. It says: Hey Wayne. I get the data that's in the attached Excel spreadsheet from 
these 5 reports (which he then lists). Can you automate this so that this spreadsheet will be updated weekly with those 
values? Note that it needs to add new information to the end of the spreadsheet, not just replace the data. 
Oh yeah… can you have everything formatted like it currently is? When you look at the spreadsheet, there are many rows 
of data (one per week), and many columns. Some columns are text, some numbers, and some are calculations. 

Some of the text is left-justified, some center, some right-justified. Some numbers have no decimals, some have one or two. 
And some are percentages.

So, you get to work. You get a procedure together that gathers the information. But when you use the OpenRowset method 
to insert the data into the spreadsheet, there is no formatting. So you decide to investigate whether you can get T-SQL 
to perform Excel Automation to do all the work for you. (Note – this is not to say that there aren't other, better ways 
to do this. It's just what you decided to do.)

So, what is Excel Automation? Simply put, it's having one application (in our case, SQL Server) essentially 
"drive" Excel, using built-in Excel properties and methods. In SQL Server, this is accomplished by use of the 
sp_OA stored procedures.

Obviously, if SQL Server is going to drive Excel, then Excel needs to be installed on the server that it's running from.

The first thing that SQL needs to do is to open up an instance of Excel. The code to do that is:

declare @xlApp integer, @rs integer
execute @rs = dbo.sp_OACreate 'Excel.Application', @xlApp OUTPUT

So what we done is to start up the excel application. The variable @xlApp is a handle to the application.

I have found it useful to set the Excel Properties "ScreenUpdating" and "DisplayAlerts" to false.

ScreenUpdating turned off will speed up the code, and you won't be looking at it anyway. 
DisplayAlerts turned off will prevent prompts requiring a response from appearing; Excel will use the default response. 
These are set by:

execute @rs = master.dbo.sp_OASetProperty @xlApp, 'ScreenUpdating', 'False'
execute @rs = master.dbo.sp_OASetProperty @xlApp, 'DisplayAlerts', 'False'
Now we need to get a handle to the open workbooks. The code to do that is:

declare @xlWorkbooks integer
execute @rs = master.dbo.sp_OAMethod @xlApp, 'Workbooks', @xlWorkbooks OUTUT

Now we have a decision to make. Are we going to open an existing spreadsheet, or make a new one?

To open an existing one:

declare @xlWorkbook integer
execute @rs = master.dbo.sp_OAMethod @xlWorkbooks, 'Open', @xlWorkBook OUTPUT, 'C:\Myspreadsheet.xls'
To add a new workbook:

declare @xlWorkBook integer
execute @rs = master.dbo.sp_OAMethod @xlWorkBooks, 'Add', @xlWorkBook OUTPUT, -4167
(The –4167 is the value of the constant xlWBATWorksheet, which specifies to add a new worksheet).

Now that we have a handle to the workbook, we have to get a handle to the worksheet:

declare @xlWorkSheet integer
execute @rs = master.dbo.sp_OAMethod @xlWorkBook, 'ActiveSheet', @xlWorkSheet OUTPUT
Now we have to find out the last row. Thankfully, Excel can tell us that:

declare @xlLastRow integer
execute @rs = master.dbo.sp_OAGetProperty @xlWorkSheet, 'Cells.SpecialCells(11).Row', @xlLastRow OUTPUT
If you want to, you can also get the last column:

declare @xlLastColumn integer
execute @rs = master.dbo.sp_OAGetProperty @xlWorkSheet, 'Cells.SpecialCells(11).Column', @xlLastColumn OUTPUT
After all of this setup work, we're finally ready to start putting the data into the spreadsheet.

First, you need to get a handle to the cell:

declare @xlCell integer
set @LastRow = @LastRow + 1
execute master.dbo.sp_OAGetProperty @xlWorkSheet, 'Cells', @xlCell OUTPUT, @LastRow, 1
Now we put the data into that cell:

execute @rs = master.dbo.sp_OASetProperty @xlCell, 'Value', @Value
If you want to format that cell:

Execute @rs = master.dbo.sp_OASetProperty @xlCell, 'NumberFormat', '0%'
0% sets it to be a percentage with no decimals. 
0.0% sets it to be a percentage with one decimal. 
'd-mmm' set it to be a character date in the format "10 Oct" 
'mm-dd-yyyy' sets it to be a character date in this good old format. 
'mm-dd-yyyy hh:mm:ss' sets it to be a character date in the date/time format. 
'$#,##0.00' sets it to be a number with the currency symbol, 2 decimal points, and at least one whole number. Numbers would be separated by comma at every third number. 
Setting font settings are a little harder:

Declare @objProp varchar(200)
Set @objProp = 'Font.Bold'
Execute @rs = master.dbo.sp_OASetProperty @xlCell, @objProp, 'True'
(You can underline it by using Font.Underline)

One big note: everything that you have a pointer to needs to be destroyed at some point in time. 
So, before you move on to a new cell, you need to:

execute @rs = master.dbo.sp_OADestroy @xlCell
Now you need to save and close the file and close Excel:

Declare @FileName varchar(100)
Set @FileName = 'C:\MyNewExcelSpreadsheet.xls'
execute @rs = master.dbo.sp_OAMethod @xlWorkBook 'SaveAs', null, @FileName, -4143
(The –4143 is the file format constant to save the file as.)

execute @rs = master.dbo.sp_OAMethod @xlWorkBook, 'Close'
execute @rs = master.dbo.sp_OAMethod @xlApp, 'Quit'
Several other things you can do:

To change the name of the workbook:
execute @rs = master.dbo.sp_OASetProperty @xlWorkBook, 'Title', 'My workbook name'

To change the name of the sheet:
execute @rs = master.dbo.sp_OASetProperty @xlWorkSheet, 'Name', 'My sheet name'

To get the format of an existing cell:
execute @rs = master.dbo.sp_OAGetProperty @xlCell, 'NumberFormat', @Value OUTPUT

To get the value of an existing cell:
execute @rs = master.dbo.sp_OAGetProperty @xlCell, 'Value', @Value OUTPUT

If you want to automatically size all of the columns to be the width of the widest data:

execute @rs = master.dbo.sp_OAMethod @xlWorkSheet, 'Columns.AutoFit'

Finally, I did say earlier that all pointers need to be destroyed:

execute @rs = master.dbo.sp_OADestroy @xlWorkSheet
execute @rs = master.dbo.sp_OADestroy @xlWorkBook
execute @rs = master.dbo.sp_OADestroy @xlWorkBooks
execute @rs = master.dbo.sp_OADestroy @xlApp
If you want to use a formula, set the value of the cell to the formula, ie: '=sum(a4.a50)' or '=(+a4+a5)/a6'. Note that the equal sign must be the first character to signify a formula.

Notice that in all of the sp_OA procedure calls, I put the result of the call into the variable @rs. This can be evaluated to return many errors:

If @rs <> 0 execute master.dbo.sp_OAGetErrorInfo @Object, @OA_Source OUTPUT, @OA_Descr OUTPUT, @HelpFile OUTPUT, @HelpID OUTPUT
Note that you're not limited to working with spreadsheets – you can work with charts also.

One last note: Excel's help file gives us most of this information. Just look under "Programming Information", and then under "Microsoft Excel Visual Basic Reference" for all of the objects, methods and properties that can be used. Occasionally I would have to look up the constant values on the Internet – just do a search on the constant name.

