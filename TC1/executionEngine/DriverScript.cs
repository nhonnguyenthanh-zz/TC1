
using System;
using System.Reflection;
using TC1.config;
using TC1.utility;

namespace TC1.executionEngine
{
    class DriverScript
    {
        public static ActionKeywords actionKeywords;
        public static MethodInfo mt;
        public static Type type;
        public static String sActionKeyword;
        public static String sPageObject;
        public static int iTestStep;
        public static int iTestLastStep;
        public static String sTestCaseID;
        public static String sRunMode;
        public static String sData;
        public static bool bResult;
        public static bool bSkip;

        public DriverScript() {

            actionKeywords = new ActionKeywords();
            //method = typeof(ActionKeywords).GetMethods();
            type = actionKeywords.GetType();
            log4net.Config.BasicConfigurator.Configure();
        }


        public static void Main(String[] args)
        {
            
            ExcelUtils.setExcelFile(Constants.Path_TestData);
            new RespositoryParser(Constants.Path_OR);
            DriverScript startEngine = new DriverScript();
            startEngine.execute_TestCase();
            ExcelUtils.ExcelWBook.Save();
            ExcelUtils.excel.Quit();

        }

        private void execute_TestCase()
        {
            int iTotalTestCases = ExcelUtils.getRowCount(Constants.Sheet_TestCases);
            for (int iTestcase = 2; iTestcase <= iTotalTestCases; iTestcase++)
            {
                bResult = true;
                bSkip = false;
                sTestCaseID = ExcelUtils.getCellData(iTestcase, Constants.Col_TestCaseID, Constants.Sheet_TestCases);
                sRunMode = ExcelUtils.getCellData(iTestcase, Constants.Col_RunMode, Constants.Sheet_TestCases);
                if (sRunMode.Equals("Yes"))
                {
                    Log.startTestCase(sTestCaseID);
                    iTestStep = ExcelUtils.getRowContains(sTestCaseID, Constants.Col_TestCaseID, Constants.Sheet_TestSteps);
                    iTestLastStep = ExcelUtils.getTestStepsCount(Constants.Sheet_TestSteps, sTestCaseID, iTestStep);
                    bResult = true;
                    for (; iTestStep <= iTestLastStep; iTestStep++)
                    {
                        sActionKeyword = ExcelUtils.getCellData(iTestStep, Constants.Col_ActionKeyword, Constants.Sheet_TestSteps);
                        sPageObject = ExcelUtils.getCellData(iTestStep, Constants.Col_PageObject, Constants.Sheet_TestSteps);
                        sData = ExcelUtils.getCellData(iTestStep, Constants.Col_DataSet, Constants.Sheet_TestSteps);
                        execute_Actions();

                        if (bResult == false)
                        {


                            ExcelUtils.setCellData(Constants.KEYWORD_FAIL, iTestcase, Constants.Col_Result, Constants.Sheet_TestCases);
                            break;
                        }


                    }
                    if (bResult == true)
                    {
                        ExcelUtils.setCellData(Constants.KEYWORD_PASS, iTestcase, Constants.Col_Result, Constants.Sheet_TestCases);
                        Log.endTestCase(sTestCaseID);
                    }
                }
            }
           // throw new NotImplementedException();
        }
            
        

        private void execute_Actions()
        {
            mt = type.GetMethod(sActionKeyword);
            mt.Invoke(null, new object[] { sPageObject, sData });
            if (bResult == true)
            {
                ExcelUtils.setCellData(Constants.KEYWORD_PASS, iTestStep, Constants.Col_TestStepResult, Constants.Sheet_TestSteps);
            }
            else
            {
                ExcelUtils.setCellData(Constants.KEYWORD_FAIL, iTestStep, Constants.Col_TestStepResult, Constants.Sheet_TestSteps);
                ActionKeywords.closeBrowser("", "");
                
            }

            //throw new NotImplementedException();
        }
    }
}

