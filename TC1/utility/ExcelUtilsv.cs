

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
        
        public static String getCellData(int RowNum, int ColNum, string SheetName)
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

        public static int getRowCount(String SheetName)
        {
            int iNumber = 1;
            
            try
            {
                ExcelWSheet = ExcelWBook.Sheets[SheetName];     
                Range lastRow = ExcelWSheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing);
                iNumber = lastRow.Row;
            }
            catch (Exception ex)
            {
                Log.error("Class Utils | Method setExcelFile | Exception desc : " + ex.Message);
                DriverScript.bResult = false;             
            }
            return iNumber;
        }
        public static int getRowContains(String sTestCaseName, int colNum, String SheetName)
        {

            int iRowNum = 1;	
            try 
            {
               
        		int rowCount = ExcelUtils.getRowCount(SheetName)+1;
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
        public static int getTestStepsCount(String SheetName, String sTestCaseID, int iTestCaseStart)
        {
            
            int number = 1;
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
                return number;
            }
        }

        public static void setCellData(String Result, int RowNum, int ColNum, String SheetName)
        {
            try
            {

                ExcelWSheet = ExcelWBook.Sheets[SheetName];
                var Cell = ExcelWSheet.Cells[RowNum, ColNum];
                Cell.Value = Result;
                
            }
            catch (Exception ex){
                Log.error("Class Utils | Method setCellData | Exception desc : " + ex.Message);
                DriverScript.bResult = false;
                
            }
        }
    }
}
