

using System;
using Microsoft.Office.Interop.Excel;
using System.Reflection;
using TC1.executionEngine;

namespace TC1.utility
{
    class ExcelUtils
    {
        public static Application excel;
        public static Workbook ExcelWBook;
        private static Worksheet ExcelWSheet;

        public static void setExcelFile(string FilePath)
        {
            //string path = FilePath;
            try {

                excel = new Application();
                //excel.Visible = true;
                ExcelWBook = excel.Workbooks.Open(FilePath, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);

            } catch (Exception ex){
                Log.error("Class Utils | Method setExcelFile | Exception desc : " + ex.Message);
                DriverScript.bResult = false;
            }
        }
        
        public static string getCellData(int RowNum, int ColNum, string SheetName)
        {
            
            try
            {
                ExcelWSheet = ExcelWBook.Sheets[SheetName];
                string cell = ExcelWSheet.Cells[RowNum, ColNum].Value == null ? string.Empty : ExcelWSheet.Cells[RowNum, ColNum].Value.ToString();
                return cell;
                
            }
            catch (Exception ex){
                Log.error("Class Utils | Method setExcelFile | Exception desc : " + ex.Message);
                DriverScript.bResult = false;
                return null;
            }
            
        }

        public static int getRowCount(string SheetName)
        {
            int iNumber = 1;
            
            try
            {
                ExcelWSheet = ExcelWBook.Sheets[SheetName];     
                Range lastRow = ExcelWSheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing);
                iNumber = lastRow.Row+1;
            }
            catch (Exception ex)
            {
                Log.error("Class Utils | Method setExcelFile | Exception desc : " + ex.Message);
                DriverScript.bResult = false;             
            }
            return iNumber;
        }
        public static int getRowContains(string sTestCaseName, int colNum, string SheetName)
        {

            int iRowNum = 1;	
            try 
            {
               
        		int rowCount = ExcelUtils.getRowCount(SheetName);
        		for(; iRowNum<rowCount; iRowNum++)
                {
        			if(ExcelUtils.getCellData(iRowNum,colNum,SheetName).Equals(sTestCaseName))
                    {
        				break;
        			}
        		}
            } 
            catch (Exception ex){
                Log.error("Class Utils | Method setExcelFile | Exception desc : " + ex.Message);
                DriverScript.bResult = false;
            }
        	return iRowNum;
        }
        public static int getTestStepsCount(string SheetName, string sTestCaseID, int iTestCaseStart)
        {
            int number = 0;  
            try
            {
               for (int i = iTestCaseStart; i <= ExcelUtils.getRowCount(SheetName); i++)
               {
                   if (!sTestCaseID.Equals(ExcelUtils.getCellData(i, config.Constants.Col_TestCaseID, SheetName)))
                   {
                       number = i;
                       return number;
                   }
               }
               ExcelWSheet = ExcelWBook.Sheets[SheetName];
               Range lastRow = ExcelWSheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing);
               number = lastRow.Row;
               return number;
            }
            catch (Exception ex)
            {
                Log.error("Class Utils | Method setExcelFile | Exception desc : " + ex.Message);
                DriverScript.bResult = false;
                return 0;
            }
        }

        public static void setCellData(string Result, int RowNum, int ColNum, string SheetName)
        {
            try
            {

                ExcelWSheet = ExcelWBook.Sheets[SheetName];
                var Cell = ExcelWSheet.Cells[RowNum, ColNum];
                Cell.Value = Result;
                ExcelUtils.ExcelWBook.Save();

            }
            catch (Exception ex){
                Log.error("Class Utils | Method setCellData | Exception desc : " + ex.Message);
                DriverScript.bResult = false;
                
            }
        }
        public static void reportBug(string filename, int totalBug)
        {
            try
            {
                setExcelFile(filename);
                ExcelWSheet = ExcelWBook.Sheets["Sheet1"];
                var Cell = ExcelWSheet.Cells[1, 2];
                Cell.Value = totalBug.ToString();
                ExcelUtils.ExcelWBook.Save();

            }
            catch (Exception ex)
            {
                Log.error("Class Utils | Method reportBug | Exception desc : " + ex.Message);
            }
        }
        public static int countBugFail(int colResult,string SheetName)
        {
            int count = 0;
            try
            {
                for (int i=2 ; i <= ExcelUtils.getRowCount(SheetName); i++)
                {
                    string result = ExcelUtils.getCellData(i, colResult, SheetName).ToLower().ToString();
                    if (result.Equals("fail"))
                        count = +count;
                }
                return count;
            }
            catch (Exception ex)
            {
                Log.error("Class Utils | Method countBugFail | Exception desc : " + ex.Message);
                return 0;
            }
        }
    }
}
