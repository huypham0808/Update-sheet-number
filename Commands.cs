using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using AcAp = Autodesk.AutoCAD.ApplicationServices.Application;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Exception = Autodesk.AutoCAD.Runtime.Exception;
using System.IO;
using System.Drawing;

namespace UpdateSheetNumInModel
{
    public class Commands
    {
        [CommandMethod("CSS_MTEXT2EXCEL")]
        public void ExportMtext2Excel()
        {
            var doc = AcAp.DocumentManager.MdiActiveDocument;
            var db = doc.Database;
            var ed = doc.Editor;

            PromptSelectionOptions promptGetMtext = new PromptSelectionOptions();
            promptGetMtext.MessageForAdding = "Select MText objects by crossing selection: ";
            promptGetMtext.SingleOnly = false;
            promptGetMtext.SinglePickInSpace = false;
            TypedValue[] filterList = new TypedValue[]
            {
                new TypedValue((int)DxfCode.Start, "MTEXT"),
                new TypedValue((int)DxfCode.LayerName, "Sheetnumber")
            };


            SelectionFilter selFilter = new SelectionFilter(filterList);
            
            PromptSelectionResult selectionResult = ed.GetSelection(promptGetMtext, selFilter);
            
            if (selectionResult.Status != PromptStatus.OK)
            {
                ed.WriteMessage("\nNo valid MText objects on 'Sheetnumber' layer selected.");
                return;
            }
            //string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            string excelFilePath = @"C:\temp\MTextExport.xlsx";
            Excel.Application excelApp = new Excel.Application();
            excelApp.Visible = true;
            Excel.Workbook workbook = excelApp.Workbooks.Add();
            Excel.Worksheet worksheet = workbook.Sheets[1];
            try
            {
                int row = 1;
                worksheet.Cells[row, 1] = "ObjectID";
                worksheet.Cells[row, 2] = "SHEET NUMBER";
                worksheet.Cells[row, 3] = "NEW SHEET NUMBER";

                Excel.Range headerRange = worksheet.Range[worksheet.Cells[row, 1], worksheet.Cells[row, 3]];
                headerRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Orange);
                row++;

                using (var tr = db.TransactionManager.StartTransaction())
                {
                    #region Old solution
                    foreach (SelectedObject selObjc in selectionResult.Value)
                    {
                        if (selObjc != null)
                        {
                            if (selObjc.ObjectId.ObjectClass.Name == "AcDbMText")
                            {
                                MText mtext = tr.GetObject(selObjc.ObjectId, OpenMode.ForRead) as MText;
                                if (mtext != null)
                                {
                                    worksheet.Cells[row, 1] = mtext.Id.Handle.ToString();
                                    worksheet.Cells[row, 2] = mtext.Contents;
                                    row++;
                                }
                            }
                        }
                    }
                    tr.Commit();
                    #endregion
                }
                worksheet.UsedRange.Columns.AutoFit();
                workbook.SaveAs(excelFilePath);
                MessageBox.Show("Exported all selected sheets number successfully!", "Information", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Information);

            }
            catch (Exception ex)
            {
                MessageBox.Show($"\nError occurred: {ex.Message}", "Error", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Error);
            }
            finally
            {
                Marshal.ReleaseComObject(worksheet);
                Marshal.ReleaseComObject(workbook);
                Marshal.ReleaseComObject(excelApp);
            }
        }
        [CommandMethod("CSS_EXCEL2MTEXT")]
        public void ImportExcel2Mtext()
        {
            Document doc = AcAp.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            PromptSelectionOptions promptGetMtext = new PromptSelectionOptions();
            promptGetMtext.MessageForAdding = "Select MText objects by crossing selection: ";
            promptGetMtext.SingleOnly = false;
            promptGetMtext.SinglePickInSpace = false;
            TypedValue[] filterList = new TypedValue[]
            {
                new TypedValue((int)DxfCode.Start, "MTEXT"),
                new TypedValue((int)DxfCode.LayerName, "Sheetnumber")
            };
            SelectionFilter selFilter = new SelectionFilter(filterList);

            PromptSelectionResult selectionResult = ed.GetSelection(promptGetMtext, selFilter);

            if (selectionResult.Status != PromptStatus.OK)
            {
                ed.WriteMessage("\nNo objects selected.");
                return;
            }
            Excel.Application excelApp = new Excel.Application();
            excelApp.Visible = true;
            Excel.Workbook workbook = excelApp.Workbooks.Open(@"C:\temp\MTextExport.xlsx");
            Excel.Worksheet worksheet = workbook.Sheets[1];

            try
            {
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    foreach (SelectedObject selObj in selectionResult.Value)
                    {
                        if (selObj != null)
                        {
                            if (selObj.ObjectId.ObjectClass.Name == "AcDbMText")
                            {
                                MText mtext = tr.GetObject(selObj.ObjectId, OpenMode.ForWrite) as MText;
                                if (mtext != null)
                                {
                                    string objectIdStr = mtext.Id.Handle.ToString();
                                    string newContent = GetExcelContentForObjectId(worksheet, objectIdStr);

                                    if (!string.IsNullOrEmpty(newContent))
                                    {
                                        mtext.Contents = newContent;
                                    }
                                }
                            }
                        }
                    }
                    tr.Commit();
                }
                workbook.Close();
                excelApp.Quit();
                MessageBox.Show("Import successfully", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"\nError: {ex.Message}", "Infor", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Information);
            }
            finally
            {
                
                Marshal.ReleaseComObject(worksheet);
                Marshal.ReleaseComObject(workbook);
                Marshal.ReleaseComObject(excelApp);
                worksheet = null;
                workbook = null;
                excelApp = null;

                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }
        private string GetExcelContentForObjectId(Excel.Worksheet worksheet, string objectId)
        {
            Excel.Range idRange = worksheet.Columns[1];
            Excel.Range contentRange = worksheet.Columns[3];

            for (int i = 1; i <= idRange.Rows.Count; i++)
            {
                string idValue = idRange.Cells[i].Value2.ToString();
                if (idValue == objectId)
                {
                    return contentRange.Cells[i].Value2.ToString();
                }
            }

            return null; // If content for the ObjectId is not found in Excel
        }
    }
}
