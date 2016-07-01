using System;
using System.IO;
using System.Reflection;

namespace TC1.config
{
    class Constants
    {
        //System Variables
        public static String URL = "http://1.54.249.84/User/Login";
        public static string path = Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "..\\..\\"));
        public static String Path_TestData = path + @"/dataEngine/DataEngine.xlsx";
        //public static String Path_TestData = @"E:/WorkNew/SeleniumC#/TC1/TC1/dataEngine/DataEngine.xlsx";
        public static String Path_OR = path + @"/config/ObjectRepository.ini";
	    public static String File_TestData = @"DataEngine.xlsx";
        public static String Path_FileReport = path + @"/report/Report.xlsx";
        public static String Path_FileLog = path + @"/LogFile/log4Net.config";
	    public static String KEYWORD_FAIL = "FAIL";
	    public static String KEYWORD_PASS = "PASS";

	    //Data Sheet Column Numbers in Sheet TestSteps
	    public static int Col_TestCaseID = 1;
        public static int Col_TestScenarioID = 2;
        public static int Col_PageObject = 5;
        public static int Col_ActionKeyword = 6;
        public static int Col_RunMode = 3;
        public static int Col_Result = 4;
        public static int Col_DataSet = 7;
        public static int Col_TestStepResult = 8;

        //Data Sheet Column Numbers in Sheet DataTest
        public static int Col_ProductName = 5;
        public static int Col_Quantity = 6;
        public static int Col_UoM = 7;

        // Data Engine Excel sheets
        public static  String Sheet_TestSteps = "Test Steps";
	    public static  String Sheet_TestCases = "Test Cases";
	    public static  String Sheet_TestData = "Test Data";
    }
}
