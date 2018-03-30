using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Input;
using Excel = Microsoft.Office.Interop.Excel;
using System.Configuration;
using Microsoft.Practices.EnterpriseLibrary.Logging;
using System.Diagnostics;

namespace WhatsDependentAndHow.Commands
{
    public class ButtonGenerateTreeCommand : ICommand
    {
        private MainWindowViewModel _mainWindowViewModel;
        private CellObject _workBookCellObjects = new CellObject() { Name = "ROOT" };
        private Dictionary<string, CellObject> _allWorkBookCellObjects = new Dictionary<string, CellObject>();
        private string _mode = string.Empty;
        private bool _incompleteProcessing = false;
        private int _totalNodesInTree = 0;

        private int _maxChildren = int.Parse(ConfigurationManager.AppSettings["MaxChildren"].ToString());
        private int _maxTreeNodes = int.Parse(ConfigurationManager.AppSettings["MaxTreeNodes"].ToString());
        private int _timeOutInMinutes = int.Parse(ConfigurationManager.AppSettings["TimeoutInMinutes"].ToString());
        private int _maxDepth = int.Parse(ConfigurationManager.AppSettings["MaxDepth"].ToString());

        public ButtonGenerateTreeCommand(MainWindowViewModel mainWindowViewModel)
        {
            _mainWindowViewModel = mainWindowViewModel;
            Logger.SetLogWriter(new LogWriterFactory().Create());
        }

        public event EventHandler CanExecuteChanged;

        public bool CanExecute(object parameter)
        {
            return true;
            //CanExecuteChanged(parameter, new EventArgs());

            //if (_mainWindowViewModel.IsCellAddressValid)
            //    return true;
            //else
            //    return false;
        }

        public void Execute(object parameter)
        {
            _mainWindowViewModel.IsBusy = true;

            _mode = parameter.ToString();
            _mainWindowViewModel.StatusMessage = string.Empty;
            _workBookCellObjects.Children.Clear();
            _allWorkBookCellObjects.Clear();
            _incompleteProcessing = false;
            _totalNodesInTree = 0;

            DoWorkAsync();
        }

        private async Task DoWorkAsync()
        {
            await Task.Run(() => LoadTree());

            _mainWindowViewModel.IsBusy = false;
        }

        private void LoadTree()
        {
            Excel.Application xlApp = _mainWindowViewModel.XlApp;
            Excel.Workbook xlWorkbook = _mainWindowViewModel.XlWorkBook;

            if(xlApp == null || xlWorkbook == null)
            {
                UpdateStatus("xlApp or xlWorkbook is null. Can't proceed.");
                return;
            }

            Excel.Worksheet selectedExcelWorkSheet = null;
            Excel.Range selectedExcelCell = null;

            try
            {
                selectedExcelWorkSheet = xlWorkbook.Sheets[_mainWindowViewModel.SelectedSheetName];
                selectedExcelCell = selectedExcelWorkSheet.Range[_mainWindowViewModel.CellAddress];

                if (ValidateInput(selectedExcelWorkSheet, selectedExcelCell) == false)
                    return;

                UpdateStatus(string.Format("Generating {0} Tree for {1}!{2}...", _mode, selectedExcelWorkSheet.Name, selectedExcelCell.Address[false, false]));

                Stopwatch stopWatch = Stopwatch.StartNew();

                Stack<Excel.Range> stackExcelCells = new Stack<Excel.Range>();
                stackExcelCells.Push(selectedExcelCell);

                while(stackExcelCells.Count > 0)
                {
                    if (CheckTimeoutAndTotalNodeLimit(stopWatch.Elapsed.Minutes) == false)
                        break;

                    Excel.Range poppedExcelCell = stackExcelCells.Pop();

                    System.Windows.Application.Current.Dispatcher.Invoke(() =>
                    {
                        AddChildren(poppedExcelCell, ref stackExcelCells);
                    });
                }

                if (_incompleteProcessing)
                    UpdateStatus(string.Format("At least one path is not processed completely. See log for details. Total nodes processed: {0}. Total time taken: {1}", _totalNodesInTree, stopWatch.Elapsed));
                else
                    UpdateStatus(string.Format("All processing complete. Total nodes processed: {0}. Total time taken: {1}.", _totalNodesInTree, stopWatch.Elapsed));

                System.Windows.Application.Current.Dispatcher.Invoke(() =>
                {
                    _mainWindowViewModel.WorkBookRootCellObject = _workBookCellObjects;
                });

            }
            catch(Exception e)
            {
                UpdateStatus(string.Format("Error: {0} \n Stack Trace: {1}", e.Message, e.StackTrace));
            }
        }

        private void AddChildren(Excel.Range poppedExcelCell, ref Stack<Excel.Range> stackExcelCells)
        {
            CellObject poppedCell = PopulateCellObjectMetadata(poppedExcelCell);

            GetCellObjectReference(ref poppedCell);

            if (CheckDepthLimit(poppedCell) == false)
                return;

            UpdateStatus(string.Format("Now processing: '{0}'. Getting {1}...", poppedCell.Name, _mode));

            List<Excel.Range> excelChildrenOfPoppedCell = GetChildren(ref poppedExcelCell, ref stackExcelCells);

            if(excelChildrenOfPoppedCell != null)
            {
                UpdateStatus(string.Format("'{0}' has {1} {2}", poppedCell.Name, excelChildrenOfPoppedCell.Count, _mode));

                if(excelChildrenOfPoppedCell.Count > 0)
                {
                    AddChildrenToTree(poppedCell, excelChildrenOfPoppedCell);
                }
            }
        }

        private void AddChildrenToTree(CellObject poppedCell, List<Excel.Range> excelChildrenOfPoppedCell)
        {
            CellObject cellObjOfGlobalCollection;

            foreach (Excel.Range excelChild in excelChildrenOfPoppedCell)
            {
                CellObject child = PopulateCellObjectMetadata(excelChild);

                if(_allWorkBookCellObjects.TryGetValue(child.Name, out cellObjOfGlobalCollection))
                {
                    child = cellObjOfGlobalCollection;
                    
                    if (!poppedCell.Children.Contains(child))
                        poppedCell.Children.Add(child);
                }
                else
                {
                    _allWorkBookCellObjects.Add(child.Name, child);
                    poppedCell.Children.Add(child);
                }

                child.Parent = poppedCell;
                _totalNodesInTree++;
            }
        }

        private List<Excel.Range> GetChildren(ref Excel.Range poppedExcelCell, ref Stack<Excel.Range> stackExcelCells)
        {
            string sourceAddress = string.Format("{0}!{1}", poppedExcelCell.Worksheet.Name, poppedExcelCell.Address);
            int arrowNumber = 1;

            bool towardPrecedent = (_mode == "Precedents") ? true : false;

            if (towardPrecedent)
                poppedExcelCell.ShowPrecedents(false);
            else
                poppedExcelCell.ShowDependents(false);

            List<Excel.Range> children = new List<Excel.Range>();

            do
            {
                string targetAddress = null;
                int linkNumber = 1;

                do
                {
                    try
                    {
                        Excel.Range target = poppedExcelCell.NavigateArrow(towardPrecedent, arrowNumber, linkNumber++);

                        targetAddress = string.Format("{0}!{1}", target.Worksheet.Name, target.Address);

                        if (sourceAddress == targetAddress)
                            break;

                        if(target.Count > _maxChildren)
                        {
                            children.Add(target);
                            stackExcelCells.Push(target);
                        }
                        else
                        {
                            var listTarget = target.Cast<Excel.Range>().ToList().Where(x => x.Value != null);

                            foreach(Excel.Range cell in listTarget)
                            {
                                if(cell.Value != null)
                                {
                                    children.Add(cell);
                                    stackExcelCells.Push(cell);
                                }
                            }
                        }
                    }
                    catch(COMException cex)
                    {
                        if (cex.Message == "NavigateArrow method of Range class failed")
                            break;
                        throw;
                    }
                } while (true);

                if (sourceAddress == targetAddress)
                    break;

                arrowNumber++;

            } while (true);

            poppedExcelCell.Worksheet.ClearArrows();
            return children;
        }

        private void GetCellObjectReference(ref CellObject poppedCell)
        {
            CellObject cellObjectOfGlobalCollection;
            if(_allWorkBookCellObjects.TryGetValue(poppedCell.Name, out cellObjectOfGlobalCollection))
            {
                poppedCell = cellObjectOfGlobalCollection;
            }
            else
            {
                _allWorkBookCellObjects.Add(poppedCell.Name, poppedCell);
                _workBookCellObjects.Children.Add(poppedCell);
                poppedCell.Parent = _workBookCellObjects;
                _totalNodesInTree++;
            }
        }

        private bool CheckDepthLimit(CellObject cellObject)
        {
            int currentDepth = 0;
            CellObject clonedObject = CloneExtensions.CloneFactory.GetClone(cellObject);

            while(true)
            {
                CellObject parent = clonedObject.Parent;

                if (parent == null)
                    break;

                currentDepth++;
                clonedObject = parent;
            }

            if (currentDepth >= _maxDepth)
            {
                UpdateStatus(string.Format("Max Allowed Depth {0} Reached for {1}. Detected depth: {2} NOTE: Won't process it's {3}", _maxDepth, cellObject.Name, currentDepth, _mode));
                _incompleteProcessing = true;
                return false;
            }

            return true;
        }

        private bool CheckTimeoutAndTotalNodeLimit(int minutes)
        {
            if(minutes > _timeOutInMinutes)
            {
                UpdateStatus("Timeout of " + _timeOutInMinutes + " reached. NOTE: Exiting with incomplete processing.");
                _incompleteProcessing = true;
                return false;
            }

            if(_totalNodesInTree >= _maxTreeNodes)
            {
                UpdateStatus("Processed maximum allowed node: " + _allWorkBookCellObjects.Count + ". NOTE: Exiting with incomplete processing.");
                _incompleteProcessing = true;
                return false;
            }

            return true;
        }

        private bool ValidateInput(Excel.Worksheet selectedWorkSheet, Excel.Range selectedCell)
        {
            if(string.IsNullOrEmpty(_mainWindowViewModel.CellAddress) || string.IsNullOrEmpty(_mainWindowViewModel.SelectedSheetName))
            {
                UpdateStatus("Please select Sheet Name from Combo and Enter Cell Address in Text Box.");
                return false;
            }

            if(selectedWorkSheet == null)
            {
                UpdateStatus("Selected Sheet not found in Excel Workbook. Can't proceed.");
                return false;
            }

            if(selectedCell == null || selectedCell.Value2 == null || selectedCell.Value2.ToString().Trim().Length == 0)
            {
                UpdateStatus("Entered Cell Address doesn't exist or it is blank. Can't proceed.");
                return false;
            }

            if(string.IsNullOrEmpty(_mode))
            {
                UpdateStatus("Command Parameter is blank - code needs to be checked. Can't Proceed.");
                return false;
            }

            return true;
        }

        private void UpdateStatus(string message)
        {
            Logger.Write(message);
            System.Windows.Application.Current.Dispatcher.Invoke(() =>
            {
                _mainWindowViewModel.StatusMessage += string.Format("[{0} {1}]: {2}\n", DateTime.Now.ToShortDateString(), DateTime.Now.ToShortTimeString(), message);
            });
        }

        private CellObject PopulateCellObjectMetadata(Excel.Range cell)
        {
            CellObject cellObject = new CellObject();

            cellObject.Name = string.Format("{0}!{1}", cell.Worksheet.Name, cell.Address[false, false]);
            cellObject.Formula = cell.Formula ?? "";
            cellObject.Value = cell.Value == null ? "" : cell.Value.ToString();
            cellObject.RowHeading = ((cell.Worksheet.Cells[1][cell.Row]).Value == null ? "" : (cell.Worksheet.Cells[1][cell.Row]).Value.ToString());

            return cellObject;
        }
    }
}
