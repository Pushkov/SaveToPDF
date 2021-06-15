using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System.Windows.Forms;
using System.Diagnostics;
using System.IO;

namespace SaveToPDF
{
    static class Program
    {
        static SldWorks swApp;
        
        [STAThread]
        static void Main()
        {
            swApp = GetSolidWorks();
            
            if (isApp())
            {
                IModelDoc2 doc = swApp.ActiveDoc;
                if (isDrawing(doc))
                {
                    IDrawingDoc drawing = (IDrawingDoc)doc;
                    string fullDrawingPath = doc.GetPathName();
                    string drawingPath = getPath(fullDrawingPath);

                    string currentSheet = (drawing.GetCurrentSheet() as Sheet).GetName();

                    foreach(string name in drawing.GetSheetNames())
                    {
                        drawing.ActivateSheet(name);

                        Sheet sheet = drawing.GetCurrentSheet();

                        object sheetProps = sheet.GetProperties();
                        double[] sheetPropsAsArray = (double[])sheetProps;


                        string pdfFileName = getFileName(fullDrawingPath).Replace(".pdf", @"_" + sheet.GetName() + ".pdf");
                        string format = getFormat((int)sheetPropsAsArray[0], sheetPropsAsArray[5], sheetPropsAsArray[6]);
                        string fullPdfPath = drawingPath + @"PDF\" + format + @"\";

                        if (!Directory.Exists(fullPdfPath))
                        {
                            Directory.CreateDirectory(fullPdfPath);
                        }
                        ExportPdfData data = swApp.GetExportFileData((int)swExportDataFileType_e.swExportPdfData);
                        data.ExportAs3D = false;
                        data.ViewPdfAfterSaving = false;
                        data.SetSheets( (int) swExportDataSheetsToExport_e.swExportData_ExportSpecifiedSheets, sheet);

                        doc.Extension.SaveAs(fullPdfPath + pdfFileName, 0, 0, data, 0, 0);
                    }
                    drawing.ActivateSheet(currentSheet);
                    Frame frame = swApp.Frame();
                    frame.SetStatusBarText("Создание PDF файлов завершено.");
                }
            }
            Application.Exit();
        }

        static string getFormat(int type, double wigth, double heigth)
        {
            string format = "";

            if (type != 12)
            {
                switch (type)
                {
                    case 6: return "А4";
                    case 7: return "А4";
                    case 8: return "А3";
                    case 9: return "А2";
                    case 10: return "А1";
                    case 11: return "А0";
                }
            }
            else
            {
                switch (heigth * 1000)
                {
                    case 297:
                        format = "А4";
                        break;
                    case 420:
                        format = "А3";
                        break;
                    case 594:
                        format = "А2";
                        break;
                    case 841:
                        format =  "А1";
                        break;
                }
            }
            if (wigth * 1000 / getFormatWigth(format) > 1)
                return format + "x" + wigth * 1000 / getFormatWigth(format);
            else
                return format;
        }

        static int getFormatWigth(string key)
        {
            switch (key)
            {
                case "А4": return 210;
                case "А3": return 297;
                case "А2": return 420;
                case "А1": return 594;
            }
            return 0;
        }

        static string getFileName(string fullPath)
        {
            string[] words = fullPath.Split(new char[] { '\\' });
            return words[words.Length-1].Replace("SLDDRW", "pdf");
        }

        static string getPath(string fullPath)
        {
            string[] words = fullPath.Split(new char[] { '\\' });
            return fullPath.Replace(words[words.Length - 1], "");
        }

        static bool isApp()
        {
            return swApp != null;
        }

        static bool isDrawing(IModelDoc2 doc)
        {
            return doc.GetType() == (int)swDocumentTypes_e.swDocDRAWING;
        }


        private static SldWorks GetSolidWorks()
        {
            Process[] processes = Process.GetProcessesByName("SLDWORKS");
            Process SolidWorks = processes[0]; int ID = SolidWorks.Id;
            try
            {
                return (SldWorks)ROTHelper.GetActiveObjectList(ID.ToString())
                    .Where(keyvalue => (keyvalue.Key.ToLower().Contains("solidworks")))
                    .Select(keyvalue => keyvalue.Value)
                    .First();
            }
            catch { return null; }
        }

    }
}