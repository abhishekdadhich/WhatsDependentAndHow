using System;
using Microsoft.Practices.EnterpriseLibrary.Logging;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Win32;
using System.Configuration;
using Excel = Microsoft.Office.Interop.Excel;

namespace WhatsDependentAndHow.Helpers
{
    public static class Helpers
    {
        static Helpers()
        {
            try { Logger.SetLogWriter(new LogWriterFactory().Create()); } catch { }
        }

        public static OpenFileDialog GetExcelOpenFileDialog(string title)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();

            openFileDialog.Filter = ConfigurationManager.AppSettings["OpenDialogFilter"].ToString();
            openFileDialog.InitialDirectory = @ConfigurationManager.AppSettings["DefaultDirectory"].ToString();
            openFileDialog.Title = title;

            return openFileDialog;
        }

        public static void OpenExcelFile(out Excel.Application xlApp, out Excel.Workbook xlWorkBook, string filePath)
        {
            try
            {
                xlApp = new Excel.Application();
                xlApp.DisplayAlerts = false;
                xlApp.AskToUpdateLinks = false;
                xlWorkBook = xlApp.Workbooks.Open(filePath, UpdateLinks: false, ReadOnly: true);
            }
            catch (Exception)
            {
                xlApp = null;
                xlWorkBook = null;
                throw;
            }
        }
    }
}
