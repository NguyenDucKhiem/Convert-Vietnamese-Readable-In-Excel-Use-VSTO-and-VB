using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelImportData
{
    public partial class Helllo
    {
        private void Helllo_Load(object sender, RibbonUIEventArgs e)
        {

        }
        /// <summary>
        /// action click btnHellWord
        /// write my personal information in cell "B2"
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnHelloWord_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Worksheet activeWorksheet = Globals.ThisAddIn.Application.ActiveSheet as Excel.Worksheet;

            if (activeWorksheet != null)
            {
                //Assign cell B2 to range
                Excel.Range range1 = activeWorksheet.get_Range("B2", System.Type.Missing);
                //set Column Width 70
                range1.ColumnWidth = 70;

                //write my personal information
                range1.Value2 = "Hello! My name is Khiem.\n" +
                    "I'm study in Ha Noi University of Science and Technology.\n" +
                    "This is my Project 2 on convert numbers to Vietnamese readable in excel.";
            }
        }
    }
}
