using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
#region --- LICENSE GNU
//********************************************************************************************
//                           LICENSE INFORMATION
//********************************************************************************************
// MBA_Pep. Copyright (C) 2020 Vladimir A. Maltapar
//   Email: maltapar@gmail.com
// Created: 07 January 2020
//
// This program is free software: you can redistribute it and/or modify  it under the terms of
// the GNU General Public License as published by the Free Software Foundation, 
// either version 3 of the License, or (at your option) any later version.
//
// This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; 
// without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.
// See the GNU General Public License for more details.
//
// You should have received a copy of the GNU General Public License along with this program.
// If not, see http://www.gnu.org/licenses/.
//********************************************************************************************
#endregion --- LICENSE GNU
#region --- Small Description
/*
1. Appointment:
   - Implementation of daily reports (for now - statistics on laboratory analyzers,
     and it is possible for all users who periodically need to receive reports
     from MS SQL databases and possibly store them somewhere). Report Result
     can be displayed on screen or immediately displayed in Excel or Word.
   - Adding information to MIS (records in MS SQL tables: for example, add results
     manual analysis, add a new doctor / laboratory assistant to the IIA).

2. Features for the user:
   - the convenience of selecting reports for the user. He can (and already knows :)
     and edit the list of names of your reports yourself, as he wants,
     without the help of a programmer;
   - saving the results of each user's reports in their directory (they will not be lost :)

3. Features for the programmer:
   - ease of adding and changing reports: no need to recompile the program;
   - one program for different users,
     but the functionality, the name of the program and its icons - each user has their own;
   - ease of installing new functionality: a file with a new report is simply copied
     to the directory of the desired user;
   - no need to bother with the appearance of the report - all design takes place
     1 not in the program, but in Excel'e or in Word'e;
   - The initial execution of the report can even be entrusted to a competent user
     (with macro recording enabled).

More details:

Ease of choosing reports - when you select, a standard Windows Explorer window opens
with a list of files. These files are the names of user reports.
He can rename them as he sees fit,
so that the list of his reports is conveniently located when choosing.
(These are his reports, his list, and let him "bring beauty" himself,
setting order, indentation and everything that comes to his mind,
but as part of the standard for naming files in Windows :).

For each user, a directory with the names of all his reports is indicated in the settings.
Also in the settings indicate the directory where the results are saved by default
execution of reports in Excel or Word format.
By default, the selected reports are added to the name of the resulting reports.
report parameters and date-time of its execution
(so that someday later he could still find him :).
If you wish, when saving the report results, the user
can change the file name as you wish.
(Was it worth giving the user such an opportunity?
He will call it “bad” - he himself will not find it later! Himself to blame! :)

All reporting is done in the natural environment of the user.
- in Excel'e or in Word'e when the macro recording is enabled.
Then this macro is saved in a special "macro library" with the name,
tied to this report.
A mini-correction of a recorded macro usually consists of replacing
absolute address constants to standard Excel/Word predefined constants.
The name of this macro is specified as a parameter.
when generating a report in a * .sql file - NOT IN THE MAIN PROGRAM!
THAT'S ALL!
Reporting programming has been reduced to almost zero!
And it’s separated from the main program,
THE PROGRAM BY ADDING / MODIFYING REPORTS DOES NOT CHANGE!

The names of user reports are text files with the extension .sql,
(it would be nice to come up with a name / term for this).
For the user, these are just the names of his reports,
in fact, the names of files located in an accessible (only to him) directory.
The user can only be given rights to rename these files.
The .sql extension can be changed to any other,
so as not to embarrass the curiously advanced users,
setting in the program settings its extension for each user,
but is it necessary now?
For the programmer, these are text files with the extension .sql - the text of the SQL script,
which will be submitted for execution in MS SQL Server.
After the script itself, there may also be parameters in the same text file,
which are processed by the program, but not transferred to MS SQL Server.
The programmer can edit these text files "on the go"
in any text editor and immediately execute and see the result.


A simple example of the file "Sapphire400 - 01 Анализы за сегодня, ПО ВРЕМЕНИ, ПО УБЫВАНИЮ.sql":
------------------------------------------------------------------------------------------------
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
------------------------------------------------------------------------------------------------

Example of file "Аналтика 03 по Sapphire 500 ЗАДЕРЖКИ за выбранный период .sql" with parameters:
------------------------------------------------------------------------------------------------
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
--Конец папаметров - далее можно писать всё, что угодно!
 ,substring(ResultText,1,18) as DatAnal
 ,Cast(substring(ResultText,1,18) as DATEtime) as DatAn
 ,ResultDate - Cast(substring(ResultText,1,18) as DATEtime) as Lat
 ,convert(varchar(5), ResultDate - Cast(substring(ResultText,1,18) as DATEtime), 108) as Late
 ID анализатора
 where analyzer_Id = @anId   
 AND ResultDate between @dat1 and @dat2+0.999
-- set @dat1='25.01.2020'; 
-- set @dat2='28.01.2020'; 
Номер истории
название анализа
название анализатора, где выполнялся анализ
------------------------------------------------------------------------------------------------

2. The programmer can perform various "temporary" reports, edit requests.
3. ...
4. etc. :))

   Name - the name for each user, that is set in the parameters.
   The icon is also unique for each user and is set in the parameters.
 
   For example :
    - show which analyzes are already loaded in the LIS for the current date.
      for Beckman, Super-Z, (at Sapphire 500 this is in my program).

*/
#endregion --- Small Description
namespace MBA_Rep
{
    public partial class FormMBA_Rep : Form
    {
        #region --- Общие параметры
        // параметры в строчках ini-файла 
        private static string connStr;      // строка коннекта к SQL
        private static string sModes;       // строка списка режимов работы 
        private static string PathLog;      // путь к лог-файлу с вычисленными каталогом и датой ...\GGGG-MM\BeckLog2019-08-06.txt 
        private static string PathLogDir;   // путь к лог-файлу, заданный в параметрах ini-файла  
        private static string PathErrLog;   // путь к лог-файлу ошибок 
        private string PathRep = "";        // путь к файлам с названиями отчётов.sql
        private string PathDoc = "";        // путь к файлам с результирующими отчётами (документы .xls .doc и т.д.)
        private string PathIni;             // Путь к ini-файлу (там, откуда запуск)
        private string pathIniFile;         // Путь и имя ini-файла
        private static string[] str_ini;    // строки ini-файла
        // глобальные
        private static DateTime dt0 = DateTime.Now; // время старта 
        private static DateTime dtm=dt0;            // для измерения времени выполнения запроса 
        private string UserName = System.Environment.UserName;
        private string ComputerName = System.Environment.MachineName;
        private string myIP = "";
        private string strVer = "первоначальная версия - 1.0.1. ( потом изменю :)"; // :))
        private static string AppName;      // static т.к. исп. в background-процессе
        private static string qqreq = "";   // запрос о работе ("кукареку")
        private static string sHeader1 = "";    // строка заголовка - назначение, для чего.
        private static string strV1, strV2, strV3, strV4, strV5; // на форме - напоминалки :)
        //private static Int32 nHistNo = -1;
        //private static string dateDone = "2019-12-31 23:59";  // дата-время выполнения анализа по часам на анализаторе
        //private static string dateDone999 = "31-12-2019 23:59:00.000";  // дата-время выполнения анализа для SQL
        // для работы с Excel
        private Excel.Application exApp, exApp1, exApp2;
        private Excel.Window exWind;
        private int F_Remind = 0; // флаг - напоминалки
        private Random rnd = new Random(321);
        #endregion --- Общие параметры
        #region --- для отчётов по Select
        // для результата Select
        private string scriptSql = "";     // текст Select из файла  - file.OpenText().ReadToEnd();
        private string scriptSqlPar = "";  // текст с параметрами из файла со скриптом (Select...)
        private string fnScript;
        private string pathFnCSV = ""; // инициализируется в ReadParmsIni();  // @"D:\TempData\Last_csv.csv";
        private string fnCSV = ""; // инициализируется в ReadParmsIni(); 
        private string fnSh = "Шаблон_ОтчётЛАБ.xls"; // имя файла с шаблонами всех отчётов
        private string sres;    // result of SQL select
        private int kRow = 0;   // кол-во строк в выборке по Select
        private int kCol = 0;   // кол-во колонок/полей в Select
        private string[,] aRes = null; // результат Select'а
        //private string[] aNamSel = null; // названия колонок в результате Select'а
        //private Boolean Fl_RunExel = false, Fl_RunWord=false;   // флаги: после Select вызывать Excel или Word 
        //private string sPeriod = "Период";                      // эта строка-признак того, чно надо ввести период с dat1 по dat2
        //private string sParSeparator = "Параметры";      // эта строка-признак начала параметров
        //private string sRunExcel = "Run Excel", sRunWord = "Run Word";  // строки-признаки вызова Excel или Word
        //private string sCond1 = "Условие1", sCond2 = "Условие2", sCond3 = "Условие3";  // строки-принаки наличия условий
        //private DateTime dat1, dat2;            // период выборки
        //private string cond1 = "", cond2 = "", cond3 = "";  // условия выборки - и в классе ParmRep
        //private string LastRep = "нет ничего :(";            // название последнего выбранного отчёта
        #endregion --- для отчётов по Select
        public FormMBA_Rep()
        {
            InitializeComponent();
            ReadParmsIni();         // берём параметры из ini-файла 
            FormIni();              // установки на форме, которые не делает Visual Studio - IP-адреса
            ParmRep.ParmRepIni();   // инициализация параметров запроса
        }
        private void FormIni()          // мой начальный вывод на форму, что прочитали из ini-файла
        {
            string Host = Dns.GetHostName();
            foreach (IPAddress adr in Dns.GetHostEntry(Host).AddressList)
            {
                myIP += adr.ToString() + "  ";
            }
            /*
            IPHostEntry host = Dns.GetHostByName(Host);
            foreach (IPAddress ip in host.AddressList) myIP += ip.ToString() + "  ";
            //Lbl_IP_List.Text = $"все IP: {myIP}";
            string IP = Dns.GetHostByName(Host).AddressList[0].ToString();
            //Lbl_IP.Text = "IP: " + IP;
            */
            CmbTest.SelectedIndex = 0;  // первый (нулевой) элемент - текущий, видимый.
            this.Pic1.Image = new Bitmap($"{PathIni}\\Pic.png");  // на форме картинка - для различных приложений должна быть другая!
        }
        private void ReadParmsIni()     // читать настройки из ini-файла 
        {
            PathIni = Application.StartupPath;
            AppName = AppDomain.CurrentDomain.FriendlyName;
            AppName = AppName.Substring(0, AppName.IndexOf(".exe"));
            pathIniFile = PathIni + @"\" + $"\\{AppName}" + ".ini";
            if (!File.Exists(pathIniFile))
            {
                string errmsg = "Не найден файл " + pathIniFile + "\n Работа завершается!";
                MessageBox.Show(errmsg, " Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                Environment.Exit(1);
            }

            INIManager iniFile = new INIManager(pathIniFile);   // cоздание объекта для работы с ini-файлом

            // секция [Description]
            sHeader1 = iniFile.GetPrivateString("Description", "Header1"); // получить значение  из секции Description по ключу Header1
            RtbHeader1.Text = sHeader1;
            this.Text = "  " + sHeader1;    // заголовок в основном окне  

            // секция [Connection]
            string cAnalyzer_Id = iniFile.GetPrivateString("Connection", "Analyzer_Id").Trim();
            //Analyzer_Id = Convert.ToInt32(cAnalyzer_Id);
            connStr = iniFile.GetPrivateString("Connection", "DbSQL").Trim(); // получить значение  из секции 1. по ключу 2.
            string nameSQLsrev = connStr.Substring(0, connStr.IndexOf(';'));

            // секция [ReportFiles]
            PathRep = iniFile.GetPrivateString("ReportFiles", "PathRep");
            PathDoc = iniFile.GetPrivateString("ReportFiles", "PathDoc");

            // секция [LogFiles]
            PathLogDir = iniFile.GetPrivateString("LogFiles", "PathLogDir");
            SetPathLog();
            PathErrLog = iniFile.GetPrivateString("LogFiles", "PathErrLog");

            // секция [Modes]  Режимы работы 
            sModes = iniFile.GetPrivateString("Modes", "sModes");
            qqreq  = iniFile.GetPrivateString("Modes", "qqreq");
            string sRemind = iniFile.GetPrivateString("Modes", "remind");
            int.TryParse(sRemind, out F_Remind); // F_Remind == 0 or 1
            /* Пример содержимого:
            sModes=(без лога приёма),(лог квиточка),(Log_Excel),(лог SQL),(лог пациентов без номера истории) 
            qqreq =Qq
            remind=0
            */

            // секция [Comments]
            iniFile.WritePrivateString("Comments", "TimeStart", "Начало " + dtm.ToString());// записать значение в секции Connection по ключу age

            var version = Assembly.GetExecutingAssembly().GetName().Version;
            var builtDate = new DateTime(2000, 1, 1).AddDays(version.Build).AddSeconds(version.Revision * 2);
            var versionString = String.Format("версия {0}, изм. от {1} {2}"
                , version.ToString(2), builtDate.ToShortDateString(), builtDate.ToShortTimeString());
            strVer = versionString;
            
            fnCSV =  "LastCSV.csv";
            pathFnCSV = PathDoc + "\\" + fnCSV; // имя файла будет изменено на имя файла отчёта при выполнении конкретного отчёта

            Lbl_v1.Text = "";
            Lbl_v2.Text = "";
            Lbl_v3.Text = "";
            Lbl_v4.Text = "";
            Lbl_v5.Text = "";
            if (F_Remind == 1) RemindText(); // обновить напоминалки в основном окне
            WLog($"--- Запуск {AppName} {ComputerName} {UserName} {strVer}");
        }
        /* -- ToSQL пока не используется
        private void ToSQL(string st)   // запись в MS-SQL сформированной строки  
        {
           //using (SqlConnection sqlConn = new SqlConnection(Properties.Settings.Default.connStr))
           using (SqlConnection sqlConn = new SqlConnection(connStr))
           {
               //SqlCommand sqlCmd = new SqlCommand("INSERT INTO [AnalyzerResults]([Analyzer_Id],[ResultText],[ResultDate]
               // ,[Hostname],[HistoryNumber])VALUES(@AnalyzerId,@ResultText,GETDATE(),@PCname,@HistoryNumber)", sqlConn);
               SqlCommand sqlCmd = new SqlCommand(st, sqlConn);
               try
               {
                   //Lbl_State.Text = "запись в SQL...";
                   sqlCmd.CommandType = System.Data.CommandType.Text;
                   sqlConn.Open();
                   sqlCmd.ExecuteNonQuery();
               }
               catch (Exception ex)
               {
                   string mes = $"Ошибка при записи в SQL. Номер истории: {nHistNo}."; //2019-11-13
                   WErrLog(mes + "\n" + ex.ToString());    // в файл ошибок...
                   WLog(mes + "\n" + ex.ToString());       // в лог файл тоже!
               }
           }
        }
        */
        private void SqlSel(string ps_select, ref string ps_res) // Sele * from ...
        {
            dataGridView1.DataSource = null;    // очистить предыдущую таблицу
            DateTime dt1 = DateTime.Now, dt2;
            using (SqlConnection sqlConn = new SqlConnection(connStr))
            {
                DataTable Res = new DataTable();
                try
                {   
                    SqlCommand com = new SqlCommand(ps_select, sqlConn);
                    // здесь выполняется?! где com.ExecuteNonQuery() или ExecuteReader или ExecuteScalar?
                    // https://metanit.com/sharp/adonet/2.5.php

                    // есть период...
                    if (ParmRep.IsPeriod)
                    {
                        //dat1 = ParmRep.Dat1;
                        //dat2 = ParmRep.Dat2;
                        SqlParameter namParDat1 = new SqlParameter("@dat1", ParmRep.Dat1);  // создаем параметр для dat1
                        SqlParameter namParDat2 = new SqlParameter("@dat2", ParmRep.Dat2);  
                        com.Parameters.Add(namParDat1);                             // добавляем параметр к команде
                        com.Parameters.Add(namParDat2);
                        // SqlParameter - это Выходные параметры запросов: 
                        // https://metanit.com/sharp/adonet/2.10.php
                    }
                    // есть Условие1
                    if (ParmRep.IsCond[1])
                    {
                        SqlParameter ParCond1 = new SqlParameter("@par1", ParmRep.Cond[1]); // создаем параметр для Cond1
                        com.Parameters.Add(ParCond1);                                       // добавляем параметр к команде
                    }
                    // есть Условие2
                    if (ParmRep.IsCond[2])
                    {
                        SqlParameter ParCond2 = new SqlParameter("@par2", ParmRep.Cond[2]); // создаем параметр для Cond1
                        com.Parameters.Add(ParCond2);                                       // добавляем параметр к команде
                    }
                    // есть Условие3
                    if (ParmRep.IsCond[3])
                    {
                        SqlParameter ParCond3 = new SqlParameter("@par3", ParmRep.Cond[3]); // создаем параметр для Cond1
                        com.Parameters.Add(ParCond3);                                       // добавляем параметр к команде
                    }

                    using (SqlDataAdapter adapter = new SqlDataAdapter(com))     // здесь выполняется?!
                    {
                        // https://docs.microsoft.com/ru-ru/dotnet/api/system.data.sqlclient.sqldataadapter?view=xamarinios-10.8
                        // Ответ в ссылке. Выполняется и заполняется здесь! 
                        adapter.Fill(Res);
                        dataGridView1.DataSource = Res;
                        // кол-во строк - ...dataGridView1.RowCount
                        #region --- comments SqlSel
                        /*
                        DataGridViewColumn col = new DataGridViewTextBoxColumn();
                        //DataGridViewRow row = new DataGridViewTextBoxRow();
                        foreach (DataGridViewRow item in dataGridView1.Rows)
                        {
                            item.Cells["foo"].Value = item.Cells[0].Value.ToString() + " edited text";
                            //aRes[i, j] = item.Cells[i,j].Value;
                        }
                        */

                        /*
                        kRow = dataGridView1.RowCount;
                        kCol = dataGridView1.ColumnCount;
                        aRes = new string[kRow, kCol] ;
                        for (int i=0; i < kRow; i++)
                        {
                            for (int j=0; j < kCol; j++)
                            {
                                aRes[i,j]=
                            }
                        }
                        */
                        #endregion --- comments SqlSel
                    }
                    dt2 = DateTime.Now;
                    int kRow = dataGridView1.RowCount;
                    //dt3 = dt2 - dt1;
                    Stat1.ForeColor = System.Drawing.Color.Black;
                    Stat2.ForeColor = System.Drawing.Color.Black;
                    Stat1.Text = $"Количество строк: {kRow}.  "; //      0123456789 123
                    string sVr = $"{dt2 - dt1}".Substring(3, 8); // dt1 "00:00:01.123456"
                    //Stat2.Text = $"Выполнено за {dt2 - dt1} сек. Время: {dt2}."; // формат
                    Stat2.Text = $"Выполнено за {sVr} сек. Время: {dt2}."; // формат
                    if (kRow == 0)
                    {
                        ps_res = "нет данных";
                        MessageBox.Show("Нет данных по заданным условиям выборки.", " Внимание!"
                            , MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        this.Pic1.Visible = true;
                    }
                    else ps_res = "ok";
                }
                catch (Exception ex)
                {
                    // 
                    string mes = $"Ошибка при выполнении SQL SELECT!"; //2019-12-30
                    WErrLog(mes + "\n" + ex.ToString() + "\n --- Текст SQL:\n" + ps_select);    // в файл ошибок...
                    Stat1.ForeColor = System.Drawing.Color.Red;
                    Stat2.ForeColor = System.Drawing.Color.Red;
                    Stat1.Text = $"{mes} Подробости в лог-файле. Время: {dt1}.";
                    Stat2.Text = $"{ex.Message}";
                    ps_res = mes;
                }
                finally
                {
                    if (sqlConn.State == ConnectionState.Open)
                        sqlConn.Close();
                }
            }
        }
        private void WriteToCsv(string pathFnCSV) // запись из dataGridView1 в файл .csv
        {
            FileStream fn = new FileStream(pathFnCSV, FileMode.Create);
            StreamWriter sw = new StreamWriter(fn, Encoding.GetEncoding(1251));
            string ss = "";
            for (int i = 0; i < dataGridView1.Rows[0].Cells.Count; i++)
            {   
                ss += dataGridView1.Columns[i].HeaderText + ";";
            }
            sw.WriteLine(ss); // названия колонок

            for (int j = 0; j < dataGridView1.Rows.Count; j++)
            {
                ss = "";
                for (int i = 0; i < dataGridView1.Rows[j].Cells.Count; i++)
                {  
                    ss += " "+dataGridView1.Rows[j].Cells[i].Value + ";";
                }
                sw.WriteLine(ss);
            }
            sw.Close();
            fn.Close(); 
        }
        /* ---  Вместо SqlSelect работает SqlSel !!!
        private void SqlSelect(string ps_select, ref string ps_res) // Вместо SqlSelect работает SqlSel !!!   Sele * from ...
        {
            // Вместо SqlSelect работает SqlSel
            using (SqlConnection sqlConn = new SqlConnection(connStr))
            {
                //SqlCommand sqlCmd = new SqlCommand("INSERT INTO [AnalyzerResults]([Analyzer_Id],[ResultText],[ResultDate]
                // ,[Hostname],[HistoryNumber])VALUES(@AnalyzerId,@ResultText,GETDATE(),@PCname,@HistoryNumber)", sqlConn);
                SqlCommand sqlCmd = new SqlCommand(ps_select, sqlConn);
                string si="";
                string[] astr = null; // одна строка результата Select
                // для результата Select
                int nrow = 0; // кол-во строк в выборке по Select
                int ncol = 0; // кол-во колонок/полей в Select
                string[,] aSel = null; // результат Select'а
                string[] aNameSel = null; // названия колонок в результате Select'а
                // '\n' - перевод строки; '\t' - табуляция; код Unicode: '\u0411' - кириллический символ 'Б'; '\x5A' - "Z"
                //string sa1 = '\x0D'.ToString();   // это \n 
                //string sa2 = '\x0A'.ToString();   // это \r 
                DataTable sRes = new DataTable();
                try
                {
                    
                    //Lbl_State.Text = "выполнение запроса SQL...";
                    sqlCmd.CommandType = System.Data.CommandType.Text;
                    sqlConn.Open();
                    //ps_res=sqlCmd.ExecuteReader().ToString();
                    SqlDataReader reader = sqlCmd.ExecuteReader();
                    if (!reader.HasRows) // если нет данных
                    {
                        ps_res = "нет данных в Select!";
                        this.Pic1.Visible = true;
                        return; 
                    }

                    ncol = reader.FieldCount;
                    aNameSel = new string[ncol];
                    //aNameSel = reader.GetName(i).ToString();
                    for (int i = 0; i < ncol; i++) aNameSel[i] = reader.GetName(i).ToString();

                    // reader.GetSchemaTable reader.GetString reader.GetType

                    //nrow = reader.j;

                    astr = new string[ncol];
                    aSel = new string[ncol, nrow]; //  *********************** nrow ОПРЕДЕЛИТЬ !!!!!
                    int js = 0; // по выбранным строчкам
                    while (reader.Read())
                    {
                        // элементы массива astr[] - это значения столбцов из запроса SELECT
                        for (int i = 0; i<ncol; i++)
                        {
                            si += reader[i].ToString() + ";";
                            aSel[js, i] = reader[i].ToString();
                        }
                        astr[js]=si;
                        js++;
                        si = "";
                    }
                    reader.Close(); // закрываем reader
                    sqlConn.Close();// закрываем соединение с БД
                    string fn_sel = PathIni + @"\" + $"{AppName}" + "_LastSel.txt";
                    File.WriteAllLines(fn_sel, astr, Encoding.GetEncoding(1251));
                    dataGridView1.DataSource = GetDataTable(aSel);
 
                }
                catch (Exception ex)
                {
                    string mes = $"Ошибка при чтении SQL. Номер истории (уже не нужен:) {nHistNo}."; //2019-12-30
                    WErrLog(mes + "\n" + ex.ToString());    // в файл ошибок...
                    //WLog(mes + "\n" + ex.ToString());       // в лог файл тоже!
                    sqlConn.Close();
                }
            }
            ps_res = "выполнено!";
        }
        */
        private DataTable GetDataTable(string[,] array)
        {
            DataTable table = new DataTable();
            for (int i = 0; i < array.GetLength(1); i++)
                table.Columns.Add();
            for (int i = 0; i < array.GetLength(0); i++)
            {
                table.Rows.Add(table.NewRow());
                for (int j = 0; j < array.GetLength(1); j++)
                    table.Rows[i][j] = array[i, j];
            }
            return table;
        }
        private void Btn_Excel_Click(object sender, EventArgs e) //вывод в Excel
        {
            if (dataGridView1.Rows.Count == 0)
            {
                MessageBox.Show("Нет данных - нечего выводить в Excel :((", " Внимание!"
                    , MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            string tmpFn = "Это временное название - придумайте своё.csv";
            if (ParmRep.LastRep.Length>0)
            {
                tmpFn = ParmRep.LastRep + " " + ParmRep.ListParms;
            }            

            dtm = DateTime.Now;
            ParmRep.DateDone = dtm.ToString("dd-MM-yyyy HH:mm:ss").Replace(":", "-");   // .ToString("yyyy-MM-dd")
            //fnCSV = $"{ParmRep.LastRep} {ParmRep.ListParms} (сформирован {ParmRep.DateDone})" + ".csv";
            fnCSV = $"{tmpFn} (сформирован {ParmRep.DateDone}).csv";

            pathFnCSV = PathDoc + "\\" + fnCSV;
            string pathFnSh = PathRep + "\\" + fnSh;    // путь к файлам с названиями отчётов.sql - находится там же.
            WriteToCsv(pathFnCSV);
            exApp = new Excel.Application
            {
                Visible = true
            };
            exApp.Workbooks.Open(pathFnCSV);
            exApp.DisplayAlerts = false;
            #region testExcel02 // пример заполнения
            /*
            Excel.Application excel_app = new Excel.ApplicationClass
            {
                Visible = true
            };
            // Откройте книгу.
            //string fileName = @"D:\TempData\XL00.xlsx"; //имя Excel файла 
            Excel.Workbook workbook = excel_app.Workbooks.Add();

            // Посмотрим, существует ли рабочий лист.
            string sheet_name = DateTime.Now.ToString("MM-dd-yyyy");

            Excel.Worksheet sheet; //= FindSheet(workbook, sheet_name);
                                   //if (sheet == null)
                                   //{
                                   // Добавить лист в конце.
            sheet = (Excel.Worksheet)workbook.Sheets.Add(
                Type.Missing, workbook.Sheets[workbook.Sheets.Count],
                1, Excel.XlSheetType.xlWorksheet);
            sheet.Name = DateTime.Now.ToString("MM-dd-yy");
            //}

            // Добавить некоторые данные в отдельные ячейки.
            sheet.Cells[1, 1] = "A";
            sheet.Cells[1, 2] = "B";
            sheet.Cells[1, 3] = "C";
            // Делаем этот диапазон ячеек жирным и красным.
            Excel.Range header_range = sheet.get_Range("A1", "C1");
            header_range.Font.Bold = true;
            header_range.Font.Color =
                System.Drawing.ColorTranslator.ToOle(
                    System.Drawing.Color.Red);
            header_range.Interior.Color =
                System.Drawing.ColorTranslator.ToOle(
                    System.Drawing.Color.Pink);
            // Добавьте некоторые данные в диапазон ячеек.
            int[,] sheet_values =
            {   { 2,  4,  6},
                { 3,  6,  9},
                { 4,  8, 12},
                { 5, 10, 15},
            };
            Excel.Range value_range = sheet.get_Range("A2", "C5");
            value_range.Value2 = sheet_values;
            */
            #endregion testExcel02 // пример заполнения
            #region testExcel01 // пример заполнения
            // Сохраните изменения и закройте книгу.
            //workbook.Close(true, Type.Missing, Type.Missing);
            // Закройте сервер Excel.
            //excel_app.Quit();
            //MessageBox.Show("Done");
            // 44444444444444444444444444444444444444444444444444444444444444444444444


            // 33333333333333333333333333333333333333333333333
            /* // Получить объект приложения Excel.
            Excel.Application excel_app = new Excel.ApplicationClass();
            excel_app.Visible = true;
            // Откройте книгу.
            string fileName = @"D:\TempData\XL00.xlsx"; //имя Excel файла 
            Excel.Workbook workbook = excel_app.Workbooks.Open(
                fileName,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing);

            // Посмотрим, существует ли рабочий лист.
            string sheet_name = DateTime.Now.ToString("MM-dd-yy");

            Excel.Worksheet sheet; //= FindSheet(workbook, sheet_name);
            //if (sheet == null)
            //{
                // Добавить лист в конце.
                sheet = (Excel.Worksheet)workbook.Sheets.Add(
                    Type.Missing, workbook.Sheets[workbook.Sheets.Count],
                    1, Excel.XlSheetType.xlWorksheet);
                sheet.Name = DateTime.Now.ToString("MM-dd-yy");
            //}

            // Добавить некоторые данные в отдельные ячейки.
            sheet.Cells[1, 1] = "A";
            sheet.Cells[1, 2] = "B";
            sheet.Cells[1, 3] = "C";
            // Делаем этот диапазон ячеек жирным и красным.
            Excel.Range header_range = sheet.get_Range("A1", "C1");
            header_range.Font.Bold = true;
            header_range.Font.Color =
                System.Drawing.ColorTranslator.ToOle(
                    System.Drawing.Color.Red);
            header_range.Interior.Color =
                System.Drawing.ColorTranslator.ToOle(
                    System.Drawing.Color.Pink);
            // Добавьте некоторые данные в диапазон ячеек.
            int[,] values =
            {   { 2,  4,  6},
                { 3,  6,  9},
                { 4,  8, 12},
                { 5, 10, 15},
            };
            Excel.Range value_range = sheet.get_Range("A2", "C5");
            value_range.Value2 = values;

            // Сохраните изменения и закройте книгу.
            workbook.Close(true, Type.Missing, Type.Missing);
            // Закройте сервер Excel.
            //excel_app.Quit();
            MessageBox.Show("Done");
            */ // 33333333333333333333333333333333333333333333333

            /*// 22222222222222222222222222222222222
            Excel.Application app = new Excel.Application();
            app.Visible = true;
     
            string fileName = @"D:\TempData\XL01.xlsx"; //имя Excel файла 
            app.Workbooks.Open(fileName);
            Excel.Workbook book = app.ActiveWorkbook;
            Excel.Worksheet sh = (Excel.Worksheet)book.Worksheets[1];

            for (int i = 1; i <= 10; i++)
                sh.Cells[i, i] = $"Текст {i}.";

            for (int i = 1; i <= book.Worksheets.Count; i++)
            {
                Excel.Worksheet sheet = (Excel.Worksheet)book.Worksheets[i];
                sheet.Cells[1, 1] = $"Текст {i}.";
            }
            book.Save();
            //app.Quit();
            //MessageBox.Show("Файл сохранён!");
            // 2222222222222222222222222222222222222
            */


            /* // 1111111111111111111111111111111111111
            string fileName = @"D:\TempData\XL01.xlsx"; //имя Excel файла  
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWb = xlApp.Workbooks.Open(fileName); //открываем Excel файл
            xlWb.Sheet[1]
            Excel.Worksheet xlSht = xlWb.Sheets[1]; //первый лист в файле
            int iLastRow = 2; // xlSht.Cells[xlSht.Rows.Count, "A"].End[Excel.XlDirection.xlUp].Row;  //последняя заполненная строка в столбце А
            for (int i = 1; i < 51; i++)
            {
                iLastRow++;
                xlSht.Cells[iLastRow, "A"].Value = i.ToString();
                //xlSht.Cells[iLastRow, "A"] = i.ToString();
                xlSht.Cells[iLastRow, "A"] = i.ToString();
            }
            //xlApp.Visible = true;
            xlWb.Close(true); //закрыть и сохранить книгу
            xlApp.Quit();
            MessageBox.Show("Файл успешно сохранён!");
            */ // 111111111111111111111111111111111111

            /* // 00000000000000000000000000000000000
            Excel.Application xlApp = new Excel.Application
            {
                Visible = true  // Сделать приложение Excel видимым
            };
            xlApp.Workbooks.Add();
            //Excel._Worksheet workSheet = exApp.ActiveSheet();
            Excel._Worksheet workSheet = xlApp.ActiveSheet;
            // Установить заголовки столбцов в ячейках
            workSheet.Cells[1, "A"] = "NameCompany";
            workSheet.Cells[1, "B"] = "Site";
            workSheet.Cells[1, "C"] = "Cost";

            string parser = File.ReadAllText(@"parser.txt", Encoding.Default);

            int parsers = Convert.ToInt32(parser);
            int row = 1;
            foreach (Price c in vPices)
            {
                row++;
                workSheet.Cells[parsers, "A"] = c.Name;
                workSheet.Cells[parsers, "B"] = c.Site;
                workSheet.Cells[parsers, "C"] = c.Cost;
            }


            xlApp.DisplayAlerts = false;
            workSheet.SaveAs(string.Format(@"{0}\Price.xlsx", Environment.CurrentDirectory));

            xlApp.Quit();
            */ // 000000000000000000000000000000000000000
            #endregion testExcel01 // пример заполнения
        }
        // ---
        #region --- ( Easter eggs :))
        private void PictureBox2_Click(object sender, EventArgs e)
        {
            //pictureBox2.Visible = !pictureBox2.Visible;
        }
        private void PictureBox1_Click(object sender, EventArgs e)
        {
            //pictureBox2.Visible = !pictureBox2.Visible;
        }
        private void RemindText() // Текст напоминалок - отображается в основном окне
        {
            str_ini = File.ReadAllLines(pathIniFile, Encoding.GetEncoding(1251));
            // для случайного
            int k1 = 0, k2 = 0, k3 = 0, k4 = 0;
            for (int i = 0; i < str_ini.Length; i++)
            {
                if (str_ini[i].IndexOf("*** start test ***") != -1) k1 = i; // ограничитель начала напоминалок
                if (str_ini[i].IndexOf("*** end test ***") != -1) k2 = i; // ограничитель конца  напоминалок
            }
            if ((k1 < k2) & (k1 != 0) & (k2 != 0))
            {
                k1++; k2--;  // текст по 5 строк с k1 - начало, по k2 - конец  (весь текст в ini-файле :)
                //Random rnd = new Random();       // сразу в глобальных c инициализацией
                int irand = rnd.Next(0, k2 - k1);  //очередное  случайное число в диапазоне k1 - k2.
                k3 = irand / 5;
                k4 = k1 + k3 * 5; // по 5 строк на "напоминалку" - strV1-strV5 
                strV1 = str_ini[k4];
                strV2 = str_ini[k4 + 1];
                strV3 = str_ini[k4 + 2];
                strV4 = str_ini[k4 + 3];
                strV5 = str_ini[k4 + 4];
                Lbl_v1.Text = strV1;
                Lbl_v2.Text = strV2;
                Lbl_v3.Text = strV3;
                Lbl_v4.Text = strV4;
                Lbl_v5.Text = strV5;
                Lbl_v1.Show();
                Lbl_v2.Show();
                Lbl_v3.Show();
                Lbl_v4.Show();
                Lbl_v5.Show();
                //Stat1.Text = $"ir={irand}, k1={k1}, k2={k2}, k3={k3}, k4={k4}.";
            }
        }
        // ---
        private void Lbl_v1_Click(object sender, EventArgs e)
        {
            F_Remind = 0; // флаг - напоминалки
            Lbl_v1.Text = "";
            Lbl_v2.Text = "";
            Lbl_v3.Text = "";
            Lbl_v4.Text = "";
            Lbl_v5.Text = "";
            Lbl_v1.Show();
            Lbl_v2.Show();
            Lbl_v3.Show();
            Lbl_v4.Show();
            Lbl_v5.Show();
        }
        #endregion --- ( Easter eggs :))
        // ---
        #region --- методы Wlog, WErrLog, WTest;  SetPathLog, Add_RTB, ExitApp...
        private static void SetPathLog()     // нужна, если программа работает много дней и меняется текущая дата //2019-08-06
        {
            dtm = DateTime.Now;
            PathLogDir = Path.GetFullPath(PathLogDir + @"\.");
            string PathLogGodMes = PathLogDir + @"\" + dtm.ToString("yyyy-MM");
            if (!Directory.Exists(PathLogGodMes))
                Directory.CreateDirectory(PathLogGodMes);
            //PathLog += @"\BeckLog" + $"{dtm.Year}-" +
            //    $"{dtm.Month.ToString().PadLeft(2, '0')}-" +
            //    $"{dtm.Day.ToString().PadLeft(2,'0')}.txt";
            PathLog = PathLogGodMes + @"\" + $"{AppName}_" + dtm.ToString("yyyy-MM-dd") + ".txt";
        }
        private static void WLog(string st) // записать в лог FLog
        {
            FileStream fn = new FileStream(PathLog, FileMode.Append);
            StreamWriter sw = new StreamWriter(fn, Encoding.GetEncoding(1251));
            dtm = DateTime.Now;
            string ss = dtm.ToString("yyyy-MM-dd HH:mm:ss").Replace("-", ".");
            sw.WriteLine($"{ss} {st}");
            sw.Close();
        }
        private static void WErrLog(string st) // записать ErrLog
        {
            string fnPathErrLog = PathErrLog + @"\Log_ERR.txt";
            fnPathErrLog = Path.GetFullPath(fnPathErrLog);
            FileStream fn = new FileStream(fnPathErrLog, FileMode.Append);
            StreamWriter sw = new StreamWriter(fn, Encoding.GetEncoding(1251));
            dtm = DateTime.Now;
            string ss = dtm.ToString("yyyy-MM-dd HH:mm:ss").Replace("-", ".");
            sw.WriteLine($"\n{ss} {st}");
            sw.Close();
        }
        private static void WTest(string FileNam, string st) // записать в FileNam.txt
        {
            FileStream fn = new FileStream(PathErrLog + "\\" + FileNam + ".txt", FileMode.Append);
            StreamWriter sw = new StreamWriter(fn, Encoding.GetEncoding(1251));
            dtm = DateTime.Now;
            string ss = dtm.ToString("yyyy-MM-dd HH:mm:ss").Replace("-", ".");
            sw.WriteLine($"\n{ss} {st}");
            sw.Close();
        }
        private static void Add_RTB(RichTextBox rtbOut, string addText)
        {
            Add_RTB(rtbOut, addText, Color.Black);
        }
        private static void Add_RTB(RichTextBox rtbOut, string addText, Color myColor)
        {
            Int32 p1, p2;
            p1 = rtbOut.TextLength;
            p2 = addText.Length;
            rtbOut.AppendText(addText);
            rtbOut.Select(p1, p2);
            rtbOut.SelectionColor = myColor;
            // 1 rtbOut.Select(0, 0);
            // 2 rtbOut.Select(p1 + p2, 0);
            // 2 rtbOut.AppendText("");
            rtbOut.SelectionStart = rtbOut.Text.Length;
            rtbOut.ScrollToCaret();
            // или: rtbOut.Select(p1, p2);
            //      SendKeys.Send("^{END}");  // это прокрутка в конец :)
        }
        private static void ExitApp(string mess, int ErrCode = 1001) // Завершение работы по ошибке.
        {
            string title = "Аварийное завершение работы.";
            WErrLog($"{title}\n{mess}");
            MessageBox.Show(mess, title
                , MessageBoxButtons.OK, MessageBoxIcon.Stop);
            Environment.Exit(ErrCode);
        }
        private static void ExitApp(string mess) // Нормальное завершение работы (по кнопке Х)
        {
            //    WLog("--- передумал выходить :) ");
            //    return; // просто передумал :)
            WLog("--- " + mess);
            Environment.Exit(0);
        }
        private void НастройкаToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // Меню/Сервис/Параметры в .ini-файле:
            string s = "";
            string nameSQLsrv = connStr.Substring(0, connStr.IndexOf(';'));
            string sep = $"---------------------------------------------------------------------------\n";
            s += $"              {sHeader1}\n";
            s += $"AppName: {AppName}, {strVer}\n";
            s += $"Время старта: {dt0}\n";
            s += $"ComputerName: {ComputerName}, UserName: {UserName}\n";
            s += $"IP: {myIP}\n";
            s += sep;
            s += $"путь к логам:\n";
            s += $"PathLog: {PathLog}\n";
            s += $"PathLogDir: {PathLogDir}\n";
            s += $"PathIni: {PathIni}\n";
            s += $"путь к файлам с названиями отчётов.sql (pathRep): {PathRep}\n";
            s += $"путь к файлам с результирующими отчётами (документы .xls .doc) (pathDoc): {PathDoc}\n";
            s += sep;
            s += $"Режимы работы: {sModes}\n";
            s += $"SQL: {nameSQLsrv}\n";
            s += sep;
            s += $"\n\n\n\n";
            DialogResult result = MessageBox.Show(s, "  Параметры в .ini-файле:"
                , MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        // ---+++
        private void ВыходToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ExitApp("Выход по меню <Выход>");
        }

        private void Pic1_click(object sender, EventArgs e)
        {
            this.Pic1.Visible = false;
            Lbl_v1.Visible = false;
            Lbl_v2.Visible = false;
            Lbl_v3.Visible = false;
            Lbl_v4.Visible = false;
            Lbl_v5.Visible = false;
            RtbHeader1.Location = new Point(Pic1.Location.X, Pic1.Location.Y);
            RtbHeader1.Width = this.Width - 10;
        }

        private void ВыходToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            ExitApp("Выход из меню <Выход>");
        }
        private void Form1_FormClosing(object sender, FormClosingEventArgs e) // закрыть App
        {
            // Закрытие формы - FormClosing
            string mess = "Завершение работы.";
            DialogResult result = MessageBox.Show("Вы действительно хотите завершить работу\n c программой?"
                , mess, MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (result == DialogResult.No)
            {
                WLog("--- передумал выходить :) " + e.CloseReason.ToString()
                    + " " + result.ToString()); // UserClosing No
                e.Cancel=true;
                return;
            }
            WLog("--- " + mess);
            Environment.Exit(0);
            //ExitApp(mess);
        }
        #endregion --- методы Wlog, WErrLog;  SetPathLog, Add_RTB, ExitApp...

        #region --- Menu methods --- Основное меню: Отчёты/ ...
        //private void Rep01ToolStripMenuItem_Click(object sender, EventArgs e)       // Выбрать файл со скриптом отчёта и выполнить его.
        //private void Rep01ToolStripMenuItem_Click(object sender, EventArgs e)       // Выбрать файл со скриптом отчёта и выполнить его.
        private void ОтчётыToolStripMenuItem_Click(object sender, EventArgs e)       // Выбрать файл со скриптом отчёта и выполнить его.
        {
            if (F_Remind==1) RemindText(); // обновить напоминалки в основном окне
            // Выбрать файл со скриптом отчёта и выполнить его. ОтчётыToolStripMenuItem_Click
            fnScript = "";
            ParmRep.ParmRepIni(); // инициализация параметров запроса. Период по умолчанию - текущая дата
            using (OpenFileDialog openFileDialog1 = new OpenFileDialog())
            {
                openFileDialog1.InitialDirectory = PathRep;
                openFileDialog1.Filter = "Выберите отчёт (*.sql)|*.sql|Все файлы (*.*)|*.*";
                openFileDialog1.FilterIndex = 1;
                openFileDialog1.RestoreDirectory = true;
                if (openFileDialog1.ShowDialog() == DialogResult.OK) 
                {
                    fnScript = openFileDialog1.FileName;
                    ParmRep.LastRep = fnScript.Substring(fnScript.LastIndexOf("\\")+1 );   // название последнего выбранного отчёта (без пути)
                    ParmRep.LastRep = ParmRep.LastRep.Substring(0, ParmRep.LastRep.Length - 4); // без последних 4-х знаков ".sql"
                    Stat3.Text = "Отчёт: " + ParmRep.LastRep; // записать название выполняемого отчёта
                    RtbHeader1.Text = Stat3.Text;
                }
                else
                {
                    MessageBox.Show(" Ничего не выбрано,\n сейчас никакой отчёт не выполнен!" +
                        "\n\n Отображаются старые данные!"
                       ," Обратите внимание!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
            }
            // начинается обработка параметров...
            //string[] aScript = File.ReadAllLines(fnScript);
            string[] aScript = File.ReadAllLines(fnScript, Encoding.GetEncoding(1251));
            scriptSql = String.Concat(aScript);
            Stat2.Text = $" Параметров нет."; // первоначально!

            // проверка на наличие параметров в теле скрипта запроса
            ParmRep.IndParms  = scriptSql.ToUpper().IndexOf(ParmRep.SParam.ToUpper());  // есть признак-строка начала параметров
            ParmRep.IsParms = (ParmRep.IndParms != -1);
            ParmRep.IndPeriod = scriptSql.ToUpper().IndexOf(ParmRep.SPeriod.ToUpper());        // есть признак-строка наличия периода
            ParmRep.IsPeriod = (ParmRep.IndPeriod != -1);

            //ParmRep.IsParms = ( ParmRep.IsPeriod | ParmRep.IsCond1 | ParmRep.IsCond2 | ParmRep.IsCond3 );
            //ParmRep.IsParms = (ParmRep.IsPeriod); // первоначально!
            /* // будем определять условия перебором по строкам в aScript[i]
            ParmRep.IndCondx   = scriptSql.ToUpper().IndexOf(ParmRep.SCondx); // есть признак-строка наличия хотя бы одного условия
            for (int i = 0; i < ParmRep.MaxNumCond; i++)
            {
                ParmRep.IndCond[i] = scriptSql.ToUpper().IndexOf(ParmRep.SCond[i].ToUpper()); // есть признак-строка наличия условия i
                ParmRep.IsCond[i] = (ParmRep.IndCond[i] != -1 ); // есть признак-строка наличия условия i
                ParmRep.IsParms = (ParmRep.IsPeriod | ParmRep.IsCond[i]);   // есть ли параметры
                ParmRep.CntCond = ParmRep.IsCond[i] ? (ParmRep.CntCond++) : (0); // считаем количество условий
            }
            */
            if (ParmRep.IsParms)
            {
                // обработка параметров...
                /*  //scriptSqlPar = scriptSql.Substring(iParms + sParSeparator.Length); // если в скрипте есть строка-признак sParSeparator
                // а если решим, чно он не нужен, тогда без него так:
                // определить, где начинаются параметры
                int[] iPar = { scriptSql.Length }; // первоначально в неё какое-то максимальное число
                // теперь к iPar добавим все параметры(их IndexOf), которые присутствуют в текущием скрипте
                if (ParmRep.IsPeriod) iPar = iPar.Concat(new int[] { iPeriod }).ToArray();
                if (ParmRep.IsCond1) iPar = iPar.Concat(new int[] { iCond1 }).ToArray();
                if (ParmRep.IsCond2) iPar = iPar.Concat(new int[] { iCond2 }).ToArray();
                if (ParmRep.IsCond3) iPar = iPar.Concat(new int[] { iCond3 }).ToArray();
                // ... и выберем из ним минимальный, т.е. с которого и начинаются все параметры
                // LINQ --- LINQ to Objects --- Операции Min и Max, ссылка: https://professorweb.ru/my/LINQ/base/level3/3_9.php
                int iMinPar = iPar.Min();
                */
                // или Вариант 2:
                // ведь можно и так получить iMinPar:  см. пример по ссылке: https://professorweb.ru/my/LINQ/base/level1/1_5.php
                //int[] Allindex = { iPeriod, iCond1, iCond2, iCond3 };
                //IEnumerable<int> ExistIndex = from n in Allindex where n != -1 select n;
                //int iMinPar = ExistIndex.Min();
                /*  или Вариант 3:
                IEnumerable<int> ExistIndCond = from n in ParmRep.IndCond where n != -1 select n;
                int iMinCond = ExistIndCond.Min();  // минимальный индекс из условий
                int iMinPar = ParmRep.IsPeriod ? ( Math.Min(ParmRep.IndPeriod, iMinCond) ) : (iMinCond); // индекс периода или условия

                scriptSqlPar = scriptSql.Substring(iMinPar);   // параметры начинаются от первого существующего параметра и до конца скрипта
                scriptSql = scriptSql.Substring(0, iMinPar - 1);
                */
                //int[] array = new int[] { 3, 4 }; // пример увеличения размера массива :)
                //array = array.Concat(new int[] { 2 }).ToArray();

                // Вариант 4: // параметры есть, если есть признак-строка начала параметров
                scriptSqlPar = scriptSql.Substring(ParmRep.IndParms+ ParmRep.SParam.Length);
                scriptSql = scriptSql.Substring(0, ParmRep.IndParms);

                // будем определять наличие условий перебором по строкам в aScript[i]
                for (int i = 0; i < aScript.GetLength(0); i++)
                {
                    if (aScript[i].ToUpper().IndexOf(ParmRep.SCondx.ToUpper()) != -1)
                    {
                        ParmRep.CntCond++;
                        ParmRep.NamCond[ParmRep.CntCond] = aScript[i + 1];  // в следующей строке скрипта - название условия
                        ParmRep.IndCond[ParmRep.CntCond] = ParmRep.CntCond;
                        ParmRep.IsCond[ParmRep.CntCond] = true; // есть признак-строка наличия условия k ParmRep.CntCond
                        ParmRep.ListParms += " " + ParmRep.NamCond[ParmRep.CntCond] + " " + ParmRep.Cond[ParmRep.CntCond];
                    }
                    // проверка: надо ли вызывать Word или Ехсеl. Имя макроса - это следующая строка после "RUN EXCEL"
                    if (aScript[i].ToUpper().IndexOf(ParmRep.SRunExcel.ToUpper()) != -1)
                    {
                        ParmRep.IsRunExel = true;
                        ParmRep.ExcelMacro= aScript[i + 1].Trim();  // в следующей строке скрипта - название макроса
                    }
                    if (aScript[i].ToUpper().IndexOf(ParmRep.SRunWord.ToUpper()) != -1)
                    {
                        ParmRep.IsRunWord = true;
                        ParmRep.WordMacro = aScript[i + 1].Trim();  // в следующей строке скрипта - название макроса
                    }
                }
                ParmRep.CntCond =Math.Min(3, ParmRep.CntCond); // органичим тремя, оcтальные игнор. т.к. только три условия на форме
                //Stat2.Text = $" Параметры: {ParmRep.LispParms}";
                Stat2.Text = $"{ParmRep.ListParms}";
                // если в теле скрипта есть параметры, 
                // то надо вызвать форму для выбранного отчёта для заполнения парамеров отчёта
                // Как передать данные из одной формы в другую - см. ссылку: (п.2.3)
                // http://www.cyberforum.ru/windows-forms/thread110436.html#a_Q2

                // ---------------------------------------------------------------------------------
                FormParmRep formParmRep = new FormParmRep(ParmRep.LastRep);
                formParmRep.ShowDialog();   // здесь заполнили параметры выбранного отчёта  .Show() - немодальное окно
                // ---------------------------------------------------------------------------------

                if (ParmRep.IsPeriod)
                {
                    ParmRep.ListParms += " с " +ParmRep.Dat1.ToString("dd-MM-yyyy") 
                                       + " по "+ParmRep.Dat2.ToString("dd-MM-yyyy"); // было: .ToString("yyyy-MM-dd")
                }
                Stat3.Text += "\n Параметры: " + ParmRep.ListParms;
                RtbHeader1.Text += "\n" + ParmRep.ListParms;
                //RtbHeader1.Text = Stat3.Text;

                // получить введённые значения из формы - в SqlSel !
            }

            sres = "-";
            SqlSel(scriptSql, ref sres);    // done 2020-01-13 16:40

            // нет данных в выборке
            if (sres != "ok") // была ошибка (this.dataGridView1.RowCount=0)
            {
                //MessageBox.Show("??? ничего не выбрано - была ошибка???\n Нет данных по заданным условиям выборки.", " Внимание!"
                //           , MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return; 
            }
            if (ParmRep.IsRunExel)     // вызывать Excel  2020-01-15 было: if (Chk_Excel.Checked )
            {
                dtm = DateTime.Now;
                ParmRep.DateDone = dtm.ToString("dd-MM-yyyy HH:mm:ss").Replace(":", "-");   // .ToString("yyyy-MM-dd")
                fnCSV = $"{ParmRep.LastRep} {ParmRep.ListParms} (сформирован {ParmRep.DateDone})" + ".csv";
                pathFnCSV = PathDoc + "\\" + fnCSV;
                string pathFnSh = PathRep + "\\" + fnSh;    // путь к файлам с названиями отчётов.sql - находится там же.
                WriteToCsv(pathFnCSV);
                exApp = new Excel.Application
                {
                    Visible = true
                };
                exApp.Workbooks.Open(pathFnCSV);
                exApp.Workbooks.Open(pathFnSh);
                exApp.DisplayAlerts = false;
                exApp.Run("ot", fnCSV, "LIS_Delay");
                //exApp.Windows(fnSh).Activate();
                exApp.ActiveWindow.ActivateNext();
                exApp.ActiveWindow.Close();
                int gg = 0;
                #region --- Run Excel Macro       
                /*
                fn PathFnCSV,
                fn0 fnCSV
                ole2 exApp

                m.fnCSV="Ot_Id0.xls" 
                m.PathFnCSV=pathOt+m.fnCSV
                COPY TO &fn TYPE XL5 as 1251
                m.oXL=CreateObject("Excel.Application")
                m.oXL.application.visible=.T.
                m.oXL.application.WorkBooks.Open(m.PathFnCSV)
                m.ole2=m.oXL.application

                typ_ot='hist0'
                m.ole2.DisplayAlerts=.F.
                m.ole2.WorkBooks.Open(m.fnsh)
                m.ole2.Run("ot",fnCSV,m.typ_ot) && "ALL_1" - имя макроса, m.par1 - параметр

                ole2.Windows(fnsh_n).Activate
                ole2.ActiveWindow.Close

                */
                #endregion --- Run Excel Macro       

            }
            WLog(ParmRep.LastRep);  // записать название выполняемого отчёта
        }
        private void ОпрограммеToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            //  Справка - О программе
            //System.Diagnostics.Process.Start(Environment.GetEnvironmentVariable("systemroot") + "\\System32\\calc.exe");
            //System.Diagnostics.Process.Start("calc");
            //System.Diagnostics.Process.Start("NotePad");
            //System.Diagnostics.Process.Start("Word.exe"); // так не вызывается :(
            //string sRun = Environment.GetEnvironmentVariable("systemroot") + "\\System32\\Notepad.exe" +$" {PathIni}\\{AppName}.txt";
            //System.Diagnostics.Process.Start($"{PathIni}\\{AppName}.cmd"); // и так запускается
            System.Diagnostics.Process.Start("NotePad", $"{PathIni}\\{AppName}.txt"); // так запускается
        }

        private void НастройкаToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Вместо пункта меню Настройка выбероите Параметры.\n      Пока так! :))", " Подсказка:"
                       , MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
        }
        private void Rep02ToolStripMenuItem1_Click(object sender, EventArgs e)      // Добавить/Изменить отчёт
        {
            // Добавить/Изменить отчёт
            //
            fnScript = "";
            using (OpenFileDialog openFileDialog1 = new OpenFileDialog())
            {
                openFileDialog1.InitialDirectory = PathRep;
                openFileDialog1.Filter = "Выберите отчёт (*.sql)|*.sql|Все файлы (*.*)|*.*";
                openFileDialog1.FilterIndex = 1;
                openFileDialog1.RestoreDirectory = true;
                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    fnScript = openFileDialog1.FileName;
                }
            }
            WLog("Добавить/Изменить отчёт");
        }
        #endregion --- Menu methods
        // ---
        #region --- Действия по кнопкам на форме
        private void Btn_tst_Click(object sender, EventArgs e)
        {
            FormTest frmTest = new FormTest();
            frmTest.ShowDialog();   // здесь заполнили параметры выбранного отчёта  .Show() - немодальное окно
            return;
        }
        // ---
        private void Btn_ini_Click(object sender, EventArgs e)  // Кнопка: прочитать ещё раз ini-файл
        {
            ReRead_ini_File();
        }
        private void ReRead_ini_File() // прочитать ещё раз ini-файл
        {
            ReadParmsIni();   // читать настройки из ini-файла 
            //string[] lines = File.ReadAllLines(fnPathIni, Encoding.GetEncoding(1251));  // - переобъявил в глобальных 2020-01-16
            string st = "";
            for (int i = 0; i < str_ini.Length; i++)
                st += $"{i + 1} :  " + str_ini[i] + "\n";
            MessageBox.Show(st, " Параметры в .ini-файле:");
        }
        private void BtnClear_Click(object sender, EventArgs e) // Кнопка: очистить форму
        {
            //RTBout.Clear();
            dataGridView1.DataSource = null;
        }
        private void BtnRun_Click(object sender, EventArgs e)   // кнопка Выполнить отчёт с параметрами
        {
            // кнопка Выполнить отчёт с параметрами
            FormParmRep formParmRep = new FormParmRep(ParmRep.LastRep);
            formParmRep.ShowDialog();   // здесь заполнили параметры выбранного отчёта

            MessageBox.Show($"Кто послал: {sender},\n аргументы: {e}.\n", "Выполнить отчёт с параметрами (BtnRun_Click)."
                , MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        #endregion --- Действия по кнопкам на форме
        // ---
        #region --- Тесты по кнопке Выполнить
        private void BtnRunTest_Click(object sender, EventArgs e)
        {
            //SqlCommand sqlc = new SqlCommand;
            string sel = "", sel1="", sel2="", sres = "";
            string fn_sel = @"D:\TempData\Last_Sel.txt";
            string fn0 = "Last_csv.csv";
            //string fn_csv = @"D:\TempData\Last_csv.csv";
            string fnShNam = "Шаблон_ОтчётЛАБтест.xls";
            string pathUsr = @"D:\TempData";
            //string fnSh = pathUsr + @"\..\Шаблоны\" + fnShNam;
            string fnSh = pathUsr + @"\" + fnShNam;
            // ... проверить на существование файла шаблона по пути fnsh !!!
            //
            string typ_ot = "ot_TestNew";
            //
            string[] astr = null;
            int k_col = 4;

            string testSelect = CmbTest.SelectedItem.ToString();
            int nSelected = CmbTest.SelectedIndex;
            string sMes = "";
            DateTime dt1 = DateTime.Now, dt2;

            switch (testSelect)
            {
                case "тест 00":    // 2020-03-23
                    sMes = "тест 00";
                    // ... 

                    MessageBox.Show($"{sMes}", $"Выполнен {CmbTest.SelectedItem}");
                    break;
                case "Next Remind :)":    // 2020-05-22 // Next Remind :)
                    //sMes = "тест 00";
                    RemindText();
                    //MessageBox.Show($"{sMes}", $"Выполнен {CmbTest.SelectedItem}");
                    break;
                
                case "тест 03":    // 2020-02-05                    
                    string sCheck = @"\w*"+ ParmRep.SParam+@"\w*";        // шаблон для поиска sCheck = @"\b[M]\w+";
                    Regex rg = new Regex(sCheck, RegexOptions.IgnoreCase);  // Создаем экземпляр Regex 
                    scriptSql = @"SET DATEFORMAT dmy; " +
                               "DECLARE @anId int =5;" +
                               "SELECT top 999 * FROM [LabAutoResult].[dbo].[AnalyzerResults] " +
                               "  where analyzer_Id = @anId  " +
                               "  AND ResultDate between @dat1 and @dat2 " +
                               " order by ResultDate DESC;\n  " +
                               " параметры: " +
                               "  период " +
                               " ";
                    // Получаем все совпадения  
                    MatchCollection matched = rg.Matches(scriptSql);

                    string st = "f";
                    // Выводим всех подходящих авторов  
                    for (int j = 0; j < matched.Count; j++)
                        sMes += matched[j].Value+"\n";
                    MessageBox.Show( $"{sMes}", $"Выполнен {CmbTest.SelectedItem}");
                    break;
                case "Show Last SQL text":    // 2020-002-11
                    MessageBox.Show($"{scriptSql}\n *** Параметры:\n{scriptSqlPar}\n ***", "Тест последнего SQL:"
                        , MessageBoxButtons.OK, MessageBoxIcon.Stop);
                    break;
                case "test SQL with parms":    // 2020-01-27
                    #region -- case "test SQL with parms":    // 2020-01-27
                    Stat1.Text = "test SQL with parms"; // проверка  SqlSel c параметрами
                    Stat2.Text = "";
                    Stat3.Text = "test SQL with parms"; // проверка  SqlSel c параметрами
                    string ps_select, ps_res;
                    
                    using (SqlConnection sqlConn = new SqlConnection(connStr))
                    {
                        DataTable Res = new DataTable();
                        try
                        {
                            sqlConn.Open();
                            ps_select = "SET DATEFORMAT dmy;DECLARE @dat1 date;DECLARE @dat2  date;DECLARE @anId int =5;SELECT [id],[Analyzer_Id],[HistoryNumber],[ResultDate], [CntParam],[ResultText]    ,[ParamName1], [ParamValue1], [ParamName2] ,[ParamValue2]     ,[ParamName3], [ParamValue3], [ParamName4], [ParamValue4]     ,[ParamName5], [ParamValue5], [ParamName6], [ParamValue6]     ,[ParamName7], [ParamValue7], [ParamName8], [ParamValue8]     ,[ParamName9], [ParamValue9], [ParamName10],[ParamValue10]    ,[ParamName11],[ParamValue11],[ParamName12],[ParamValue12]    ,[ParamName13],[ParamValue13],[ParamName14],[ParamValue14]    ,[ParamName15],[ParamValue15],[ParamName16],[ParamValue16]    ,[ParamName17],[ParamValue17],[ParamName18],[ParamValue18]    ,[ParamName19],[ParamValue19],[ParamName20],[ParamValue20]    ,[ParamName21],[ParamValue21],[ParamName22],[ParamValue22]  FROM [LabAutoResult].[dbo].[AnalyzerResults] where analyzer_Id = @anId  AND cast( @dat1 as date) <= CAST(ResultDate as DATE)   AND CAST(ResultDate as DATE) = cast( @dat2 as date)  order by ResultDate DESC;-- Параметры --;-- Период;";
                            ps_select = "SET DATEFORMAT dmy; " +
                                "DECLARE @anId int =5;" +
                                "SELECT top 999 * FROM [LabAutoResult].[dbo].[AnalyzerResults] " +
                                "  where analyzer_Id = @anId  " +
                                "  AND ResultDate between @dat1 and @dat2 " +
                                " order by ResultDate DESC;";
                            //ps_select ="SELECT top 99 * FROM [LabAutoResult].[dbo].[AnalyzerResults] where analyzer_Id = 5 order by ResultDate DESC;";
                            // cast( @dat1 as date) = CAST(ResultDate as DATE)
                            //    "set @dat1='25.01.2020'; " +
                            //    "set @dat2='28.01.2020'; " 
                            //    "DECLARE @dat1 date; " +
                            //    "DECLARE @dat2 date; " +
                            SqlCommand com = new SqlCommand(ps_select, sqlConn);
                            // здесь выполняется?! где com.ExecuteNonQuery() или ExecuteReader или ExecuteScalar?
                            //  https://metanit.com/sharp/adonet/2.5.php

                            // есть период...
                            ParmRep.IsPeriod = true;
                            if (ParmRep.IsPeriod)
                            {
                                /* // первоначальные значения ...
                                ParmRep.Dat1 = Convert.ToDateTime("10.01.2020");
                                ParmRep.Dat2 = Convert.ToDateTime("15.01.2020");
                                dat1 = ParmRep.Dat1;
                                dat2 = ParmRep.Dat2;
                                */

                                // вызвать форму с парамерами для выбранного отчёта
                                FormParmRep formParmRep = new FormParmRep(ParmRep.LastRep);
                                formParmRep.ShowDialog();   // здесь заполнили параметры выбранного отчёт
                                // получить новые введённые значения ...
                                //dat1 = ParmRep.Dat1;
                                //dat2 = ParmRep.Dat2;
                                //
                                SqlParameter namParDat1 = new SqlParameter("@dat1", ParmRep.Dat1);  // создаем параметр для dat1
                                com.Parameters.Add(namParDat1);                             // добавляем параметр к команде
                                
                                SqlParameter namParDat2 = new SqlParameter("@dat2", ParmRep.Dat2);
                                com.Parameters.Add(namParDat2);
                                // */
                                com.ExecuteNonQuery(); // для SELECT ExecuteNonQuery работает!!! 
                                //com.ExecuteReader();
                            }
                            // /*
                            using (SqlDataAdapter adapter = new SqlDataAdapter(com))     // здесь выполняется?!
                            {
                                // https://docs.microsoft.com/ru-ru/dotnet/api/system.data.sqlclient.sqldataadapter?view=xamarinios-10.8
                                // Ответ в ссылке. Выполняется и заполняется здесь! 
                                adapter.Fill(Res);
                                dataGridView1.DataSource = Res;
                            }
                            // */

                            dt2 = DateTime.Now;
                            int kRow = dataGridView1.RowCount;
                            Stat1.Text = $"Выполнено: {dt2}, за {dt2 - dt1} сек. Кол-во строк: {kRow}.  ";
                        }
                        catch (Exception ex)
                        {
                            // 
                            string mes = $"Ошибка при выполнении SQL SELECT!"; //2019-12-30
                            WErrLog(mes + "\n" + ex.ToString());    // в файл ошибок...
                            //Stat1.Text = $"{mes} Подробности см. в лог-файле. Время: {dt1}.";
                            Stat1.Text = $"{mes} {ex.Message} Время: {dt1}.";
                            ps_res = mes;
                        }
                        finally
                        {
                            if (sqlConn.State == ConnectionState.Open)
                                sqlConn.Close();
                        }
                    }

                    break;
                #endregion -- case "test SQL with parms":    // 2020-01-27
                case "тест - вызов Excel с макросом":
                    #region
                    exApp = new Excel.Application
                    {
                        Visible = true
                    };
                    /*
                    exApp.Workbooks.Open(fn_csv, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                        Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                        Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                    exApp2 = new Excel.Application();
                    exApp2.Workbooks.Open(fnSh, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                        Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                        Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                    */
                    exApp.Workbooks.Open(pathFnCSV);
                    exApp2 = new Excel.Application();
                    exApp2.Workbooks.Open(fnSh); 
                    exApp2.Visible = true;
                    exApp2.Run("ot",fn0,typ_ot);
                    exApp2.Run("ot2");
                    //exApp2.w
                    //exApp2.Windows(fnShNam).Activate;
                    exApp2.ActiveWindow.Close();
                    MessageBox.Show("Выполнен.", $"{CmbTest.SelectedItem}");
                    break;
                #endregion
                case "тест - вызов Excel c Last_csv.csv":
                    #region
                    exApp = new Excel.Application
                    {
                        Visible = true
                    };
                    /*
                    exApp.Workbooks.Open(fn_csv, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                        Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                        Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                    exApp2 = new Excel.Application();
                    exApp2.Workbooks.Open(fnSh, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                        Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                        Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                    */
                    exApp.Workbooks.Open(pathFnCSV);
                    exApp2 = new Excel.Application();
                    exApp2.Workbooks.Open(fnSh);
                    exApp2.Visible = true;
                    exApp2.Run("ot", fn0, typ_ot);
                    exApp2.Run("ot2");
                    //exApp2.w
                    //exApp2.Windows(fnShNam).Activate;
                    exApp2.ActiveWindow.Close();
                    MessageBox.Show("Выполнен.", $"{CmbTest.SelectedItem}");
                    break;

                    exApp = new Excel.Application();
                    exApp.Visible = true;
                    //exApp.SheetsInNewWorkbook = 3;
                    //exApp.Workbooks.Add(Type.Missing);
                    exApp.Workbooks.Open(pathFnCSV, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                        Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, 
                        Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                    MessageBox.Show("Выполнен.", $"{CmbTest.SelectedItem}");
                    break;
                    #endregion

                case "тест - запись в .csv":
                    WriteToCsv(pathFnCSV);
                    #region
                    /*
                    FileStream fn = new FileStream(fn_csv, FileMode.Append);
                    StreamWriter sw = new StreamWriter(fn, Encoding.GetEncoding(1251));
                    string ss = "";
                    for (int i = 0; i < dataGridView1.Rows[0].Cells.Count; i++)
                    {   //sw.Write(dataGridView1.Columns[k].HeaderText + ";");
                        ss += dataGridView1.Columns[i].HeaderText + ";";
                    }
                    sw.WriteLine(ss); // названия колонок

                    for (int j = 0; j < dataGridView1.Rows.Count; j++)
                    {
                        ss = "";
                        for (int i = 0; i < dataGridView1.Rows[j].Cells.Count; i++)
                        {   //sw.Write(dataGridView1.Rows[j].Cells[i].Value + ";");
                            ss += " "+dataGridView1.Rows[j].Cells[i].Value + ";";
                        }
                        //sw.WriteLine();
                        sw.WriteLine(ss);
                    }
                    sw.Close();
                    fn.Close();
                    */
                    MessageBox.Show("Выполнен.", $"{CmbTest.SelectedItem}");
                    break;
                    #endregion
                case "Выбрать файл со скриптом отчёта и выполнить его":
                    #region
                    // Выбрать файл со скриптом отчёта и выполнить его
                    fnScript = "";
                    using (OpenFileDialog openFileDialog1 = new OpenFileDialog())
                    {
                        openFileDialog1.InitialDirectory = PathRep;
                        openFileDialog1.Filter = "Выберите отчёт (*.sql)|*.sql | Все файлы (*.*)|*.*";
                        openFileDialog1.FilterIndex = 2;
                        openFileDialog1.RestoreDirectory = true;
                        if (openFileDialog1.ShowDialog() == DialogResult.OK)
                        {
                            fnScript = openFileDialog1.FileName;
                        }
                        else
                        {
                            MessageBox.Show(" Ничего не выбрано,\n сейчас никокой отчёт не выполнен!" +
                                "\n\n Отображаются старые данные!"
                                , " Обратите внимание!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            return;
                        }
                    }
                    string[] aScript = File.ReadAllLines(fnScript, Encoding.GetEncoding(1251));
                    scriptSql = String.Concat(aScript);
                    string Fl_SqlParSeparator = "-- Parameters ---";
                    int iPar = scriptSql.IndexOf(Fl_SqlParSeparator); // признак начала параметров
                    if (iPar != -1) // есть парамеры
                    {
                        scriptSqlPar = scriptSql.Substring(iPar + Fl_SqlParSeparator.Length);
                        scriptSql = scriptSql.Substring(0, iPar - 1);
                    }
                    Stat2.Text = $" Параметры: {scriptSqlPar}";
                    //
                    // ToDo 2020-01-17
                    // удалить конечные строки  - там должны быть парамеры ...
                    //
                    sres = "";
                    SqlSel(scriptSql, ref sres);    // done 2020-01-13 16:40
                    MessageBox.Show("Выполнен.", $"{CmbTest.SelectedItem}");
                    break;
                    #endregion
                case "тест - Select":
                    #region
                    //sel = "Select top 9 * FROM [LabAutoResult].[dbo].[AnalyzerResults] where analyzer_Id<=10 order by id desc";
                    /*
                    sel = "SET DATEFORMAT ymd; "+
                        " DECLARE @dat0  date = '2019.07.31'; "+
                        " DECLARE @dat   date = cast(@dat0 as date);"+
                        " declare @anId int = 6; "+
                        " SELECT[id],[Analyzer_Id],[HistoryNumber],[ResultDate], [CntParam],[ResultText] "+
                        "	,[ParamName1], [ParamValue1], [ParamName2] ,[ParamValue2] "+
                        "   ,[ParamName3], [ParamValue3], [ParamName4], [ParamValue4] "+
                        " FROM[LabAutoResult].[dbo].[AnalyzerResults] "+
                        " where analyzer_Id = 6 "+
                        " AND CAST(ResultDate as DATE) = cast(@dat as date) "+
                        " order by id desc; ";
                    */
                    //sel = "Select * FROM [LabAutoResult].[dbo].[AnalyzerResults] where analyzer_Id=9 order by ResultDate desc";
                    //sel1 = "Select top 3 * FROM [LabAutoResult].[dbo].[AnalyzerResults] where analyzer_Id=9 order by ResultDate desc;";
                    //sel2 = "Select top 5 * FROM [LabAutoResult].[dbo].[AnalyzerResults] where analyzer_Id=4 order by id desc;";
                    //sel = sel1 + sel2;
                    FileInfo fn6 = new FileInfo(@"D:\TempData\test05.sql");
                    scriptSql = fn6.OpenText().ReadToEnd();

                    /*SqlConnection sqlConn = new SqlConnection(connStr);
                    Server server = new Server(new ServerConnection(sqlConn));
                    server.ConnectionContext.ExecuteNonQuery(script);
                    */

                    sres = "";
                    SqlSel(scriptSql, ref sres);    // done 2020-01-13 16:40
                    /*
                    using (SqlConnection sqlConn = new SqlConnection(connStr))
                    {
                        DataTable Res = new DataTable();
                        try
                        {
                            SqlCommand com = new SqlCommand(scriptSql, sqlConn);
                            // здесь выполняется?! где com.ExecuteNonQuery() или ExecuteReader или ExecuteScalar?
                            //  https://metanit.com/sharp/adonet/2.5.php
                            using (SqlDataAdapter adapter = new SqlDataAdapter(com))
                            {
                                adapter.Fill(Res);
                                dataGridView1.DataSource = Res;
                            }
                        }
                        catch (Exception ex)
                        {
                            string mes = $"Ошибка при чтении SQL. Номер истории (уже не нужен:) {nHistNo}."; //2019-12-30
                            WErrLog(mes + "\n" + ex.ToString());    // в файл ошибок...
                        }
                        finally
                        {
                            if (sqlConn.State == ConnectionState.Open)
                                sqlConn.Close();
                        }
                    }
                    */
                    // вместо SqlSel(scriptSql, ref sres);    // done 2020-01-13 16:40
                    //SqlSel(sel, ref sres);

                    MessageBox.Show("Выбрано: "+sres+".", $"{CmbTest.SelectedItem}");
                    break;
                    #endregion
                case "перечитать ini-файл":
                    ReRead_ini_File();
                    break;
                case "тест 01":
                    #region
                    MessageBox.Show("выполняется ...", $"{CmbTest.SelectedItem}");
                    sel = "Select top 9 * FROM [LabAutoResult].[dbo].[AnalyzerResults] where analyzer_Id=9 order by id desc";
                    sres = "";
                    //SqlSelect(sel, ref sres);
                    MessageBox.Show("Выполнен.", $"{CmbTest.SelectedItem}");
                    break;
                    #endregion
                case "тест 02": // 2019-12-27
                    #region
                    fn_sel = @"D:\TempData\aa1.txt";
                    //string[] astr =null;
                    k_col = 4;
                    astr = new string[k_col];
                    for (int i = 0; i < k_col; i++)
                    {
                        astr[i] = $"строка номер {i}, ...";
                    }
                    File.WriteAllLines(fn_sel, astr, Encoding.GetEncoding(1251));
                    MessageBox.Show("Выполнен.", $"{CmbTest.SelectedItem}");
                    break;
                    #endregion
                case "тест - запонить dataGridView  тестовой строкой":
                    #region
                    string[,] arstr = null;
                    int ncol = 13, nrow = 27;
                    arstr = new string[nrow, ncol];
                    for (int i = 0; i < nrow; i++)
                    {
                        for (int j = 0; j < ncol; j++)
                        {
                            arstr[i, j] = $" Cтрока [ {i + 1}, {j + 1} ] ...";
                        }
                    }
                    dataGridView1.DataSource = GetDataTable(arstr);
                    MessageBox.Show("Выполнен.", $"{CmbTest.SelectedItem}");
                    break;
                    #endregion
                case "тест 04":
                    dt0 = DateTime.Now;
                    string s = " 321 456 987 987 Normal 3.4umol/L";
                    MessageBox.Show("Выполнен.", $"{CmbTest.SelectedItem}");
                    break;

                case "тест - 00x :)":
                    MessageBox.Show("Выполнен.", $"{CmbTest.SelectedItem}");
                    break;

                default:
                    MessageBox.Show($"Нет теста для:  {CmbTest.SelectedItem}\n");
                    break;
            }
        }
        #endregion --- Тесты по кнопке Выполнить
    }
}
