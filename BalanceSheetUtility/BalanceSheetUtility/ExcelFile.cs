using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;
using System.Data;
using System.IO;
using System.Runtime.InteropServices;

namespace BalanceSheetUtility
{
    public class ExcelFile 
    {
        public string excelFilePath = string.Empty;
        private int rowNumber = 1;

        public int HeaderRowNumber = 0;

        BalanceSheetUtility t_MainClass = new BalanceSheetUtility();

        Excel.Application myExcelApplication;
        Excel.Workbook myExcelWorkbook;
        Excel.Worksheet myExcelWorkSheet;

        public string ExcelFilePath
        {
            get { return excelFilePath; }
            set { excelFilePath = value; }
        }

        public int Rownumber
        {
            get { return rowNumber; }
            set { rowNumber = value; }
        }

        public int openExcel()
        {
            t_MainClass.LogProgress("Opening Excel File For Writing");
            int _ret = 0;
            try
            {
                myExcelApplication = null;

                myExcelApplication = new Excel.Application();
                myExcelApplication.DisplayAlerts = false;


                myExcelWorkbook = (Excel.Workbook)(myExcelApplication.Workbooks._Open(excelFilePath, System.Reflection.Missing.Value,
                   System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                   System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                   System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                   System.Reflection.Missing.Value, System.Reflection.Missing.Value));

                myExcelWorkSheet = (Excel.Worksheet)myExcelWorkbook.Worksheets[1];

                ////int numberOfWorkbooks = myExcelApplication.Workbooks.Count; // get number of workbooks (optional)        
                /// myExcelWorkSheet = (Excel.Worksheet)myExcelWorkbook.Worksheets[1]; // define in which worksheet, do you want to add data            
                ////myExcelWorkSheet.Name = "Sheet1"; // define a name for the worksheet (optinal)
                ////int numberOfSheets = myExcelWorkbook.Worksheets.Count; // get number of worksheets (optional)
                _ret = 0;
                t_MainClass.LogProgress("Excel File Opened Successfully");
            }
            catch (Exception EX)
            {
                t_MainClass.LogProgress("Error in Opening the File " + EX.Message );
                _ret = 1;
            }
            return _ret;
        }

        public void addDataToExcel(string firstname, string lastname, string language, string email, string company)
        {

            myExcelWorkSheet.Cells[rowNumber, "H"] = firstname;
            myExcelWorkSheet.Cells[rowNumber, "J"] = lastname;
            myExcelWorkSheet.Cells[rowNumber, "Q"] = language;
            myExcelWorkSheet.Cells[rowNumber, "BH"] = email;
            myExcelWorkSheet.Cells[rowNumber, "CH"] = company;
            rowNumber++;  // if you put this method inside a loop, you should increase rownumber by one or wat ever is your logic

        }

        public void closeExcel()
        {
            t_MainClass.LogProgress("Closing the Excel File");
            try
            {
                if (myExcelApplication != null)
                {
                    myExcelWorkbook.SaveAs(excelFilePath, System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                                                   System.Reflection.Missing.Value, System.Reflection.Missing.Value, Excel.XlSaveAsAccessMode.xlNoChange,
                                                   System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                                                   System.Reflection.Missing.Value, System.Reflection.Missing.Value); // Save data in excel


                    myExcelWorkbook.Close(true, excelFilePath, System.Reflection.Missing.Value); // close the worksheet
                }


            }
            finally
            {
                if (myExcelApplication != null)
                {
                    myExcelApplication.Quit(); // close the excel application
                    while (Marshal.ReleaseComObject(myExcelApplication) != 0) { }
                    while (Marshal.ReleaseComObject(myExcelWorkbook) != 0) { }
                    while (Marshal.ReleaseComObject(myExcelWorkSheet) != 0) { }
                    myExcelApplication = null;
                    myExcelWorkbook = null;
                    myExcelWorkSheet = null;
                   
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                    t_MainClass.LogProgress("File Closed");
                }
            }

        }

        public List<string>[] RetrieveColumnByHeader(string FindWhat)
        {
            Excel.Worksheet sheet = new Excel.Worksheet();
            sheet = myExcelWorkSheet;

            Excel.Range rngHeader = null;

            rngHeader = sheet.Rows[1] as Excel.Range;


            int rowCount = sheet.UsedRange.Rows.Count;
            int columnCount = sheet.UsedRange.Columns.Count;
            int index = 0;

            Excel.Range rngResult = null;
            string FirstAddress = null;

            List<string>[] columnValue = new List<string>[columnCount];

            rngResult = rngHeader.Find(What: FindWhat, LookIn: Excel.XlFindLookIn.xlValues,
            LookAt: Excel.XlLookAt.xlPart, SearchOrder: Excel.XlSearchOrder.xlByColumns);

            if (rngResult != null)
            {
                FirstAddress = rngResult.Address;
                Excel.Range cRng = null;

                do
                {
                    columnValue[index] = new List<string>();
                    for (int i = 1; i <= rowCount; i++)
                    {
                        cRng = sheet.Cells[i, rngResult.Column] as Excel.Range;
                        if (cRng.Value != null)
                        {
                            columnValue[index].Add(cRng.Value.ToString());
                        }
                    }

                    index++;
                    rngResult = rngHeader.FindNext(rngResult);
                } while (rngResult != null && rngResult.Address != FirstAddress);

            }
            Array.Resize(ref columnValue, index);
            return columnValue;
        }

        public List<string>[] RetrieveColumnGeneral(/*Excel.Worksheet sheet,*/ string FindWhat)
        {
            Excel.Worksheet sheet = new Excel.Worksheet();
            sheet = myExcelWorkSheet;

            int columnCount = sheet.UsedRange.Columns.Count;
            List<string>[] columnValue = new List<string>[columnCount];
            Excel.Range rngResult = null;
            Excel.Range rng = null;

            int index = 0;
            int rowCount = sheet.UsedRange.Rows.Count;
            Excel.Range FindRange = null;
            for (int columnIndex = 1; columnIndex <= sheet.UsedRange.Columns.Count; columnIndex++)
            {
                FindRange = sheet.UsedRange.Columns[columnIndex] as Excel.Range;
                FindRange.Select();
                rngResult = FindRange.Find(What: FindWhat, LookIn: Excel.XlFindLookIn.xlValues,
                    LookAt: Excel.XlLookAt.xlPart, SearchOrder: Excel.XlSearchOrder.xlByRows);
                if (rngResult != null)
                {
                    columnValue[index] = new List<string>();

                    for (int rowIndex = 1; rowIndex <= sheet.UsedRange.Rows.Count; rowIndex++)
                    {
                        rng = sheet.UsedRange[rowIndex, columnIndex] as Excel.Range;
                        if (rng.Value != null)
                        {
                            columnValue[index].Add(rng.Value.ToString());
                        }
                    }
                    index++;
                }
            }
            Array.Resize(ref columnValue, index);
            return columnValue;
        }

        public List<string> RetrieveColumnByColumnIndex(Excel.Worksheet p_Sheet, string p_ColumnIndex)
        {
            t_MainClass.LogProgress("Getting Target Excel Column List");
            List<string> _retList = new List<string>();
            string _data = string.Empty;

            try
            {
                foreach (Excel.Range row in p_Sheet.UsedRange.Rows)
                {
                    int _row = row.Row + HeaderRowNumber;
                    int _column = Convert.ToInt32(p_ColumnIndex);

                    _data = (string)(p_Sheet.Cells[_row, _column]).Value;
                    if (!string.IsNullOrEmpty(_data))
                    {
                        _retList.Add(_data);
                    }
                }
            }
            catch (Exception EX)
            {
                t_MainClass.LogProgress("Error in Getting Excel Column Data : " + EX.Message);
            }
            return _retList;
        }

        public string ManuplateExcel(DataSet p_DataSource, List<BalanceSheetUtility.CreateMappingString> p_sheetinfo, string p_SourceColumnToRead, string p_ExcelColumnToRead)
        {
            t_MainClass.LogProgress("Manuplating Excel For Writing");
            string _ret = "";
            var _workSheetAddress = string.Empty;
            var _excelColumntoRead = string.Empty;

            Excel.Worksheet _activeWorkSheet = new Excel.Worksheet();

            try
            {
                if (p_DataSource != null && p_sheetinfo.Count > 0 && p_SourceColumnToRead != string.Empty && p_ExcelColumnToRead != string.Empty)
                {
                    for (int sheetnum = 0; sheetnum < p_sheetinfo.Count; sheetnum++)
                    {
                        _workSheetAddress = string.Empty;
                        _activeWorkSheet = null;

                        if (myExcelWorkSheet != null)
                        {
                            int _sheetNumber = 0;
                            bool _isint;

                            _workSheetAddress = p_sheetinfo[sheetnum].ExcelSheetName.ToString();

                            _isint = int.TryParse(_workSheetAddress, out _sheetNumber);
                            if (_isint)
                            {
                                myExcelWorkSheet = (Excel.Worksheet)myExcelWorkbook.Worksheets[Convert.ToInt32(_workSheetAddress)];
                                _activeWorkSheet = myExcelWorkSheet;
                            }
                            else
                            {
                                for (int i = 1; i <= myExcelWorkbook.Worksheets.Count; i++)
                                {
                                    var _getWorksheet = (Excel.Worksheet)myExcelWorkbook.Worksheets[i];
                                    string _getWorkSheetName = string.Empty;
                                    _getWorkSheetName = _getWorksheet.Name.ToString();

                                    if (_getWorkSheetName.ToString().ToLower() == _workSheetAddress.ToString().ToLower())
                                    {
                                        myExcelWorkSheet = (Excel.Worksheet)myExcelWorkbook.Worksheets[i];
                                        _activeWorkSheet = myExcelWorkSheet;
                                        break;
                                    }
                                }
                            }

                            if (_activeWorkSheet != null)
                            {
                                if (p_ExcelColumnToRead.ToString().Contains(":"))
                                {
                                    _excelColumntoRead = p_ExcelColumnToRead;
                                }
                                else
                                {
                                    //_excelColumntoRead = p_ExcelColumnToRead + ":" + p_ExcelColumnToRead;
                                    _excelColumntoRead = p_ExcelColumnToRead;
                                }

                                List<string> _ExcelTargetColumnValues = RetrieveColumnByColumnIndex(_activeWorkSheet, _excelColumntoRead);

                                List<BalanceSheetUtility.CreateMappingString> _sheetInfo = new List<BalanceSheetUtility.CreateMappingString>();
                                _sheetInfo.Add(new BalanceSheetUtility.CreateMappingString { ExcelSheetName = p_sheetinfo[sheetnum].ExcelSheetName, MappingString = p_sheetinfo[sheetnum].MappingString });

                                if (_ExcelTargetColumnValues.Count > 0)
                                {
                                    _ret = WritingDatatoExcel(p_DataSource, _sheetInfo, p_SourceColumnToRead, _ExcelTargetColumnValues, _activeWorkSheet);
                                }
                                else
                                {
                                    _ret = "No Excel Data Found to Match";
                                    t_MainClass.LogProgress(_ret);
                                }
                            }
                            else
                            {
                                _ret = "No Excel Sheet opened";
                                t_MainClass.LogProgress(_ret);
                            }
                        }
                        else
                        {
                            _ret = "No Excel Sheet opened";
                            t_MainClass.LogProgress(_ret);
                        }
                    }
                }
                else
                {
                    _ret = "Please Provide input for Writing Data in Excel";
                    t_MainClass.LogProgress(_ret);
                }

            }
            catch (Exception EX)
            {
                _ret = EX.Message.ToString();
                t_MainClass.LogProgress(_ret);
            }

            return _ret;
        }

        public string WritingDatatoExcel(DataSet p_DataSource, List<BalanceSheetUtility.CreateMappingString> p_sheetinfo, string p_SourceColumnToRead, List<string> p_ExcelColumnData, Excel.Worksheet p_Sheet)
        {
            t_MainClass.LogProgress("Start Writing into Excel File");
            string _ret = "OK";
            try
            {
                if (p_ExcelColumnData.Count > 0)
                {
                    var _excelData = string.Empty;
                    for (int i = 0; i <= p_ExcelColumnData.Count - 1; i++)
                    {
                        _excelData = p_ExcelColumnData[i].ToString().Trim().ToLower();

                        if (!string.IsNullOrEmpty(_excelData))
                        {
                            for (int j = 0; j <= p_DataSource.Tables[0].Rows.Count - 1; j++)
                            {
                                var _dataSouceValue = string.Empty;
                                _dataSouceValue = p_DataSource.Tables[0].Rows[j][p_SourceColumnToRead].ToString().Trim().ToLower();

                                if (!string.IsNullOrEmpty(_excelData) && !string.IsNullOrEmpty(_dataSouceValue))
                                {
                                    if (_excelData == _dataSouceValue)
                                    {
                                        var _excelSheetName = string.Empty;
                                        var _columnMappingString = string.Empty;
                                        _excelSheetName = p_sheetinfo[0].ExcelSheetName.ToString();
                                        _columnMappingString = p_sheetinfo[0].MappingString.ToString().Trim();

                                        if (!string.IsNullOrEmpty(_columnMappingString))
                                        {
                                            List<BalanceSheetUtility.ExcelColumnMapping> _ExcelColumnMapping = new List<BalanceSheetUtility.ExcelColumnMapping>();
                                            _ExcelColumnMapping = t_MainClass.GetExcelColumnMappingList(_excelSheetName, _columnMappingString);

                                            if (_ExcelColumnMapping.Count > 0)
                                            {
                                                for (int k = 0; k <= _ExcelColumnMapping.Count - 1; k++)
                                                {
                                                    int _columnIndex = Convert.ToInt32(_ExcelColumnMapping[k].ExcelColumnNumber);
                                                    var _dataSourceColumnName = _ExcelColumnMapping[k].DataBaseColumnName.ToString().Trim();
                                                   
                                                    p_Sheet.Cells[(i + 1 + HeaderRowNumber), _columnIndex] = p_DataSource.Tables[0].Rows[j][_dataSourceColumnName].ToString();
                                                }
                                            }

                                        }
                                        else
                                        {
                                            t_MainClass.LogProgress("Excel-DataSource Column Mapping Missing");
                                            _ret = "Excel-DataSource Column Mapping Missing";
                                        }
                                        break;
                                    }
                                }

                            }
                        }
                    }
                }
            }
            catch (Exception EX)
            {
                t_MainClass.LogProgress("Error While Writing into Excel File " + EX.Message);
                _ret = EX.Message.ToString();
                return _ret;
            }
            return _ret;
        }
    }
}
