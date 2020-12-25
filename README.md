# MBA_Rep
... So, about the program:

1. Purpose:
   - Execution of daily reports (so far - statistics on laboratory analyzers,
     but it is possible for all users who periodically need to receive reports
     from MS SQL databases and maybe store them somewhere). Report result
     can be displayed on the screen or directly output to Excel or Word.
   - Adding information to the MIS (records in MS SQL tables: for example, add results
     manual analysis, add a new doctor / laboratory assistant to the MIS).

2. Features for the user:
   - Convenient selection of reports for the user. He can (and already knows how to :) arrange
     and edit the list of names of your reports yourself, the way he wants,
     without the help of a programmer;
   - saving the results of reports of each user in its own directory (will not be lost :)

3. Features for the programmer:
   - ease of adding and changing reports: no need to recompile the program;
   - one program for different users,
     but the functionality, the name of the program and its icons are different for each user;
   - easy installation of new functionality: the file with a new report is simply copied
     to the directory to the desired user;
   - no need to bother with the appearance of the report - all registration takes place
     not in a program, but in Excel or Word;
   - the initial design of the report can even be entrusted to a competent user
     (when macro recording is on).

More details:

Convenient selection of reports - when selected, a standard Windows Explorer window opens
with a list of files. These files are the names of user reports.
He can rename them as he pleases so
so that the list of its reports is conveniently located when selected.
(These are his reports, his list, and even if he himself "brings beauty",
setting the order, indents and whatever comes to mind,
but within the Windows file naming standard :).

For each user, a directory with the names of all his reports is specified in the settings.
Also, the settings specify the directory where the results are saved by default.
execution of reports in Excel or Word format.
By default, user-selected ones are added to the name of the resulting reports.
report parameters and date-time of its execution
(so that someday later he could still find him :).
If desired, when saving the report results, the user
can change the filename as it sees fit.
(Was it worth giving the user such an opportunity?
He will call "bad" - he will not find it later! Himself to blame! :)

All design of reports is performed visually in the user's natural environment
- in Excel'e or in Word'e with enabled macro recording.
Then this macro is saved in a special "macro library" with the name,
linked to this report.
A minor correction of a recorded macro usually consists of replacing
absolute address constants to standard predefined Excel / Word constants.
The name of this macro is given as a parameter
when generating a report in a * .sql file - NOT IN THE PROGRAM!
That'a all!
Programming the design of reports has been reduced to almost zero!
And it is separated from the main program,
THE PROGRAM ITSELF WHEN ADDING / CHANGING REPORTS DOESN'T CHANGE!

User report names are text files with the .sql extension,
(it would be nice to have a name / term for this here).
For the user, these are just the names of his reports,
in fact - the names of files located in a directory accessible (only to him).
The user can only be given rights to rename these files.
.Sql extension can be changed to any other,
so as not to confuse overly curious advanced users,
setting your extension for each user in the program settings,
but now is it necessary?
For the programmer, these are text files with the .sql extension - the text of the SQL script,
which will be transferred for execution to MS SQL Server.
In the same text file, after the script itself, there may also be parameters,
which are processed by the program, but not transferred to MS SQL Server.
The programmer can edit these text files "on the go"
in any text editor and immediately execute and see the result.

A simple example of the file "Sapphire400 - 01 Analyzes for today, DECREASING BY TIME.sql":
SET DATEFORMAT dmy;
DECLARE @dat0  date = getdate();
DECLARE @dat   date = cast(@dat0 as date);
DECLARE @anId int =11;
SELECT [id],[Analyzer_Id],[HistoryNumber],[ResultDate], [CntParam]
    ,[ParamName1], [ParamValue1], [ParamMsr1], [ParamRef1], [ParamName2] ,[ParamValue2], [ParamMsr2], [ParamRef2]
    ,[ParamName3], [ParamValue3], [ParamMsr3], [ParamRef3], [ParamName4], [ParamValue4], [ParamMsr4], [ParamRef4]
    ,[ParamName5], [ParamValue5], [ParamMsr5], [ParamRef5], [ParamName6], [ParamValue6], [ParamMsr6], [ParamRef6]
    ,[ParamName7], [ParamValue7], [ParamMsr7], [ParamRef7], [ParamName8], [ParamValue8], [ParamMsr8], [ParamRef8]
    ,[ParamName9], [ParamValue9], [ParamMsr9], [ParamRef9], [ParamName10],[ParamValue10],[ParamMsr10], [ParamRef10]
    ,[ParamName11],[ParamValue11],[ParamMsr11],[ParamRef11],[ParamName12],[ParamValue12],[ParamMsr12], [ParamRef12]
	,[ResultText]
  FROM [LabAutoResult].[dbo].[AnalyzerResults] where analyzer_Id = @anId
  AND CAST(ResultDate as DATE) = cast( @dat as date)
  order by ResultDate DESC;
  
  An example of the file "Analytics 03 for Sapphire 500 DELAYS for the selected period .sql" with the parameters:
  SET DATEFORMAT dmy; 
DECLARE @AnId int =5;
SELECT [id],[Analyzer_Id]
 ,Cast(substring(ResultText,1,18) as DATEtime) as DatAn
 ,convert(varchar(5), ResultDate - Cast(substring(ResultText,1,18) as DATEtime), 108) as Late
 ,[ResultDate]
 ,ResultText
 ,[HistoryNumber],[CntParam]
 ,[HostName],[ProtocolList_Id]
 ,[ParamName1]+' '+[ParamName2]+' '+[ParamName3]+' '+[ParamName4]+' '+[ParamName5]+' '+[ParamName6]+' '+[ParamName7]+' '+[ParamName8]+' '+[ParamName9]+' '+[ParamName10] as List_An
 ,[ParamName1]+', '+[ParamName2] as List_An2
 ,ParamName1+'; '+ParamName2+'; '+ParamName3+'; '+ParamName4+'; '+ParamName5+'; '+ParamName6+'; '+ParamName7+'; '+ParamName8+'; '+ParamName9+'; '+ParamName10 as List_An3
    ,[ParamName1], [ParamName2] 
    ,[ParamName3], [ParamName4]  
    ,[ParamName5], [ParamName6]  
    ,[ParamName7], [ParamName8]  
    ,[ParamName9], [ParamName10]
    ,[ParamName11],[ParamName12]
    ,[ParamName13],[ParamName14]
    ,[ParamName15],[ParamName16]
    ,[ParamName17],[ParamName18]
    ,[ParamName19],[ParamName20]
    ,[ParamName21],[ParamName22]
 FROM [LabAutoResult].[dbo].[AnalyzerResults]  
 WHERE [Analyzer_Id] = @AnId   
 AND ResultDate between @dat1 and @dat2+0.999
 order by ResultDate DESC;
Параметры
 Период;
 run Excel Rep_Late01
--

