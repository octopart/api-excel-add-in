using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Xml.Linq;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using Microsoft.Office.Interop.Excel;

namespace OctopartXll
{
    [ComVisible(true)]
    public class OctopartRibbon : ExcelRibbon
    {
        /// <summary>
        /// Log (for debugging or otherwise)
        /// </summary>
        private static readonly log4net.ILog Log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        
        /// <summary>
        /// Constructor for the Octopart Ribbon toolbar
        /// </summary>
        public OctopartRibbon()
        { }

        /// <summary>
        /// Formats all octopart url queryies to have a hyperlink
        /// </summary>
        /// <param name="control"></param>
        public void HyperlinkUrlQueries(IRibbonControl control)
        {
            try
            {
                dynamic xlApp = ExcelDnaUtil.Application;
                dynamic cellsToCheck = xlApp.ActiveSheet.Cells.SpecialCells(XlCellType.xlCellTypeFormulas);

                if (cellsToCheck != null)
                {
                    Worksheet ws = xlApp.ActiveSheet;

                    foreach (Range cell in cellsToCheck.Cells)
                    {
                        if (cell.Formula.Contains("=OCTOPART_DETAIL_URL") || cell.Formula.Contains("=OCTOPART_DATASHEET_URL") || cell.Formula.Contains("=OCTOPART_DISTRIBUTOR_URL"))
                        {
                            if (cell.Value.Contains("http"))
                            {
                                ws.Hyperlinks.Add(cell, cell.Value, Type.Missing, "Click for more information", Type.Missing);
                            }
                            else if (cell.Hyperlinks.Count > 0)
                            {
                                cell.Hyperlinks.Delete();
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Log.Error(ex.Message);
            }
        }

        /// <summary>
        /// Forces a refresh of all 'OCTOPART' type formulas
        /// </summary>
        /// <param name="control">Ribbon control button</param>
        public void RefreshAllQueries(IRibbonControl control)
        {
            try
            {
                dynamic xlApp = ExcelDnaUtil.Application;
                dynamic cellsToCheck = xlApp.ActiveSheet.Cells.SpecialCells(XlCellType.xlCellTypeFormulas);

                // This will 're-formulate' the cell, changing nothing, but causing a refresh!
                // Neat, eh?
                if (cellsToCheck != null)
                {
                    cellsToCheck.Replace(
                        @"=OCTOPART_",
                        @"=OCTOPART_",
                        XlLookAt.xlPart,
                        XlSearchOrder.xlByRows,
                        true,
                        Type.Missing,
                        Type.Missing,
                        Type.Missing);
                }
            }
            catch (Exception ex)
            {
                Log.Error(ex.Message);
            }
        }

        /// <summary>
        /// This function will upload a copy of the spreadsheet to Octopart.com
        /// </summary>
        /// <param name="control"></param>
        public void UploadBom(IRibbonControl control)
        {
            dynamic xlApp = null;
            dynamic backupCalc = null;

            try
            {
                xlApp = ExcelDnaUtil.Application;
                string fileName = null;
                xlApp.ScreenUpdating = false;
                xlApp.EnableEvents = false;
                backupCalc = xlApp.Calculation;
                xlApp.Calculation = XlCalculation.xlCalculationManual;

                // Make a copy of the workbook
                var sourceWb = xlApp.ActiveWorkbook;
                Workbook destWb = null;
                try
                {
                    xlApp.ActiveSheet.Copy();
                    destWb = xlApp.ActiveWorkbook;

                    // Remove all references (paste values instead)
                    destWb.Sheets[1].UsedRange.Cells.Copy();
                    destWb.Sheets[1].UsedRange.Cells.PasteSpecial(XlPasteType.xlPasteValues);

                    string tempPath = Environment.GetEnvironmentVariable("temp") + "\\";
                    string tempFile = "Upload Copy of " + sourceWb.Name + " " +
                                      DateTime.Now.ToString("dd-mmm-yy h-mm-ss") + ".csv";
                    fileName = tempPath + tempFile;

                    destWb.SaveAs(fileName, XlFileFormat.xlCSV);
                }
                catch (Exception ex)
                {
                    Log.Error(ex.Message);
                }
                finally
                {
                    if (destWb != null)
                        destWb.Close(SaveOptions.None);
                }

                // Post the workbook to Octopart.com
                if (!string.IsNullOrEmpty(fileName))
                {
#warning "THE REDIRECT FOR THIS IS NOT WORKING"
                    string result = OctopartApi.ApiV4.UploadBom("test@gmail.com"/*OctopartAddIn.QueryManager.Email*/, fileName);
                    if (!string.IsNullOrEmpty(result))
                    {
                        MessageBox.Show(result);
                    }
                }
            }
            catch (Exception ex)
            {
                Log.Error(ex.Message);
            }
            finally
            {
                if (xlApp != null)
                {
                    xlApp.ScreenUpdating = true;
                    xlApp.EnableEvents = true;
                    if (backupCalc != null)
                        xlApp.Calculation = backupCalc;
                }
            }
        }
    }


}
