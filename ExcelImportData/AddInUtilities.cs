using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelImportData
{
    [ComVisible(true)]
    public interface IAddInUtilities
    {
        string ImportData(Excel.Range cell);
        void Translate(Excel.Range cell);
    }

    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    public class AddInUtilities : IAddInUtilities
    {
        /// <summary>
        /// convert number in cell to Vietnamese readable and return it
        /// </summary>
        /// <param name="cell">number cell convert</param>
        /// <returns>string is Vietnamese readable</returns>
        public string ImportData(Excel.Range cell)
        {
            //read value in cell 
            dynamic value = cell.Value;
            //return "None" if cell don't have value
            if (value == null) return "None";

            try
            {
                //convert value
                return ConvertNumber.Convert(decimal.Parse(value.ToString()));
            }
            catch (Exception)
            {
                //return error if exception
                return "Error";
            }

        }

        /// <summary>
        /// Open web translate string in cell
        /// </summary>
        /// <param name="cell">cell is being click</param>
        public void Translate(Excel.Range cell)
        {
            //read value in cell
            dynamic value = cell.Value;
            //return "None" if cell don't have value
            if (value != null)
            {
                try
                {
                    System.Diagnostics.Process.Start("https://translate.google.com/#auto/en/" + value.ToString());
                }
                catch (System.ComponentModel.Win32Exception e)
                {
                    //Win32Exception
                    System.Windows.Forms.MessageBox.Show(e.ToString(), "Win32Exception");
                }
                catch (ObjectDisposedException e)
                {
                    //ObjectDisposedException
                    System.Windows.Forms.MessageBox.Show(e.ToString(), "ObjectDisposedException");
                }
                catch (System.IO.FileNotFoundException e)
                {
                    //FileNotFoundException
                    System.Windows.Forms.MessageBox.Show(e.ToString(), "FileNotFoundException");
                }
                catch (Exception e)
                {
                    // Exception
                    System.Windows.Forms.MessageBox.Show(e.ToString(), "Error");
                }
            }
        }
    }
}
