// from <https://stackoverflow.com/questions/7298765/how-to-remove-macros-from-binary-ms-office-documents>

using System;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;
using Microsoft.Vbe.Interop;

namespace OfficeMacros {
    class Program {
        static void Main(string[] args) {

            removeMacrosWord("/Users/hgnehm/Desktop/Macros.docm", "/Users/hgnehm/Desktop/Macros.docm.doc");
        }

        static void removeMacrosWord(string fileName, string newFileName) {
            var word = new Word.Application();
            var document = word.Documents.Open(fileName,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing,Type.Missing, Type.Missing);
            foreach (VBComponent component in document.VBProject.VBComponents) {
                switch (component.Type) {
                    case (vbext_ComponentType.vbext_ct_StdModule):
                    case (vbext_ComponentType.vbext_ct_MSForm):
                    case (vbext_ComponentType.vbext_ct_ClassModule):
                        document.VBProject.VBComponents.Remove(component);
                        break;
                    default:
                        component.CodeModule.DeleteLines(1, component.CodeModule.CountOfLines);
                        break;
                }
            }
            document.Close(true, newFileName, Type.Missing);
            document = null;
            word = null;
            GC.Collect();
        }

        static void RemoveMacrosExcel(string fileName, string newFileName) {
            var excel = new Excel.Application();
            var workbook = excel.Workbooks.Open(fileName,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing);

            foreach (VBComponent component in workbook.VBProject.VBComponents) {
                switch (component.Type) {
                    case (vbext_ComponentType.vbext_ct_StdModule):
                    case (vbext_ComponentType.vbext_ct_MSForm):
                    case (vbext_ComponentType.vbext_ct_ClassModule):
                        workbook.VBProject.VBComponents.Remove(component);
                        break;
                    default:
                        component.CodeModule.DeleteLines(1, component.CodeModule.CountOfLines);
                        break;
                }
            }

            workbook.Close(true, newFileName, Type.Missing);

            // Release variables
            workbook = null;
            excel = null;

            // Collect garbage
            GC.Collect();
        }

        /*
            Convert a Microsoft Word document to a PDF/A document.
        */
        private static void convertWordToPDF(string input, string output) {
            try {
                Word._Application word = new Word.Application();
                var document = word.Documents.Open(input);
                document.ExportAsFixedFormat(output, Word.WdExportFormat.wdExportFormatPDF, UseISO19005_1: true);
                document.Close();
                word.Quit();
                Console.WriteLine("   Microsoft Word document converted successfully.");
            }
            catch (Exception ex) {
                Console.WriteLine("   Error " + ex.Message);
            }
        }

    }
}
