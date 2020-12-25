using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace MBA_Rep
{
    class ParmRep // параметры отчёта - для ввода и передачи между формами
    {
        // парамертры на форме для передачи в скрипт отчёта
        public static DateTime Dat1    { get; set; }            // период с Dat1 по Dat2
        public static DateTime Dat2    { get; set; }
        public static int MaxNumCond   { get; } = 9;           // максимальное количество условий в скрипте
        public static int CntCond      { get; set; }           // текущее количество условий в скрипте - нач. с 1 !
        public static string[] Cond    { get; set; } = new string[MaxNumCond];
        public static string[] NamCond { get; set; } = new string[MaxNumCond];  // название условия 1, 2, 3
 
        public static Boolean IsParms  { get; set; }       // есть ли строка-признак -- Параметры: --
        public static Boolean IsPeriod { get; set; }       // есть ли строка-признак Период
        public static int IndParms     { get; set; } = -1; // номер индекса строки-признака начала параметров
        public static int IndPeriod    { get; set; } = -1; // номер индекса строки-признака наличия периода

        public static Boolean IsRunExel  { get; set; } = false; // флаг: после Select вызывать Excel  
        public static Boolean IsRunWord  { get; set; } = false; // флаг: после Select вызывать Word 

        public static int       IndCondx { get; set; }          // номер индекса строки-признака наличия условий
        public static Boolean   IsCondx  { get; set; }
        public static int[]     IndCond  { get; set; } = new int[MaxNumCond];     // номер индекса строки условия в скрипте
        public static Boolean[] IsCond   { get; set; } = new Boolean[MaxNumCond]; // есть ли строка-признак Условие

        // public static string[] SCond { get; set;} = new string[MaxNumCond];    // строки-принаки наличия условий
        public static string SCondx         { get; } = "Условие";   // строка-принак наличия условий: м.б. достаточно одного признака?
        public static string SParam         { get; } = "Параметры"; // эта строка-признак начала параметров
        public static string SPeriod        { get; } = "Период";    // эта строка-признак того, чно надо ввести период с dat1 по dat2
        public static string SRunExcel      { get; } = "Run Excel"; // строка-признак вызова Excel 
        public static string SRunWord       { get; } = "Run Word";  // строка-признак вызова Word
        public static string ExcelMacro{ get; set; } = "";          // имя макроса при вызове Excel 
        public static string WordMacro { get; set; } = "";          // имя макроса при вызове Word 
        public static string LastRep   { get; set; } = "";          // название последнего выбранного отчёта
        public static string ListParms { get; set; } = "";          // список параметов - для печати
        public static string DateDone  { get; set; } = "";          // список параметов - для печати

        public static void ParmRepIni()   // инициализация параметров запроса
        {
            // инициализация параметров запроса
            ParmRep.Dat1 = DateTime.Now;
            ParmRep.Dat2 = ParmRep.Dat1;
            ParmRep.IsParms   = false;
            ParmRep.IsPeriod  = false;
            ParmRep.IndParms  = -1;
            ParmRep.IndPeriod = -1;
            ParmRep.CntCond   = 0;          // текущее количество условий в скрипте - нач. с 1 !

            ParmRep.IndCondx   = -1;
            ParmRep.IsCondx    = false;

            ParmRep.IsRunExel  = false;
            ParmRep.IsRunWord  = false;
            ParmRep.ExcelMacro = "";
            ParmRep.WordMacro  = "";
            ParmRep.LastRep    = "";
            ParmRep.ListParms  = "";
            ParmRep.DateDone   = "";
            // что-то еще ...
            for (int i = 0; i< MaxNumCond; i++)
            {
                ParmRep.IndCond[i] = -1;
                ParmRep.IsCond[i]  = false;
                //ParmRep.SCond[i] = $"Условие{i}";
                ParmRep.Cond[i]    = "";       //здесь будет результат из поля, введённый на форме ввода параметров
                ParmRep.NamCond[i] = "";
            }

        }
        // ...
    }
}
