using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Input;
using Excel = Microsoft.Office.Interop.Excel;

namespace WhatsDependentAndHow.Commands
{
    public class ButtonGenerateTreeCommand : ICommand
    {
        private MainWindowViewModel _mainWindowViewModel;

        public ButtonGenerateTreeCommand(MainWindowViewModel mainWindowViewModel)
        {
            _mainWindowViewModel = mainWindowViewModel;
        }

        public event EventHandler CanExecuteChanged;

        public bool CanExecute(object parameter)
        {
            if (_mainWindowViewModel.IsCellAddressValid)
                return true;
            else
                return false;
        }

        public void Execute(object parameter)
        {
            Excel.Application xlApp = null;
            Excel.Workbook xlWorkBook = null;

            try
            {
                OpenExcelWorkbook(out xlApp, out xlWorkBook);
                CreateCellObjectsFromExcel(xlWorkBook, _mainWindowViewModel.WorkSheetNames, _mainWindowViewModel.WorkBookCellObjects);
            }
            finally
            {
                if (xlWorkBook != null)
                    xlWorkBook.Close();

                if (xlApp != null)
                    xlApp.Quit();

                Marshal.ReleaseComObject(xlWorkBook);
                Marshal.ReleaseComObject(xlApp);
            }
            
        }

        private void CreateCellObjectsFromExcel(Excel.Workbook xlWorkBook, List<string> workSheetNames, CellsKeyedCollection workBookCellObjects)
        {
            foreach (Excel.Worksheet worksheet in xlWorkBook.Worksheets)
            {
                workSheetNames.Add(worksheet.Name);

                Excel.Range usedRange = worksheet.UsedRange;

                foreach (Excel.Range cell in usedRange)
                {
                    CellObject cellObject = GetCellObjectFromExcelCell(worksheet, cell);
                    workBookCellObjects.Add(cellObject);
                }
            }
        }

        private CellObject GetCellObjectFromExcelCell(Excel.Worksheet worksheet, Excel.Range cell)
        {
            CellObject cellObject = PopulateCellObjectMetadata(worksheet, cell);

            AddLocalDependents(worksheet, cell, cellObject);

            AddLocalPrecedents(worksheet, cell, cellObject);

            return cellObject;
        }

        private void AddLocalPrecedents(Excel.Worksheet worksheet, Excel.Range cell, CellObject cellObject)
        {
            try
            {
                Excel.Range precedents = cell.Precedents;

                if(precedents != null)
                {
                    foreach (Excel.Range precedentExcelCell in precedents)
                    {
                        CellObject precedentCellObject = PopulateCellObjectMetadata(worksheet, precedentExcelCell);
                        cellObject.Precedents.Add(precedentCellObject);
                    }
                }
            }
            catch (COMException ce)
            {
                // ignore
            }
        }

        private void AddLocalDependents(Excel.Worksheet worksheet, Excel.Range cell, CellObject cellObject)
        {
            try
            {
                Excel.Range dependents = cell.Dependents;

                if (dependents != null)
                {
                    foreach (Excel.Range dependentExcelCell in dependents)
                    {
                        CellObject dependentCellObject = PopulateCellObjectMetadata(worksheet, dependentExcelCell);
                        cellObject.Dependents.Add(dependentCellObject);
                    }
                }
            }
            catch (COMException ce)
            {
                // ignore
            }
        }

        private CellObject PopulateCellObjectMetadata(Excel.Worksheet worksheet, Excel.Range cell)
        {
            CellObject cellObject = new CellObject();
            cellObject.Name = worksheet.Name + "!" + cell.Address[false, false];
            cellObject.Formula = cell.Formula ?? "";
            cellObject.Value = cell.Value == null ? "" : cell.Value.ToString();
            return cellObject;
        }

        private void OpenExcelWorkbook(out Excel.Application xlApp, out Excel.Workbook xlWorkBook)
        {
            try
            {
                xlApp = new Excel.Application();
                xlWorkBook = xlApp.Workbooks.Open(_mainWindowViewModel.ExcelFileDetails);
            }
            catch (Exception e)
            {
                xlApp = null;
                xlWorkBook = null;

                // TODO: Error Message - Error in opening excel file

                throw;
            }
        }
    }
}
