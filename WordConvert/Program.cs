using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Word2013 = Microsoft.Office.Interop.Word;
//using Spire.Doc;
namespace WordConvert
{
    class Program
    {
        static void Main(string[] args)
        {
            string templateFile = @"f:\template.docx";
            string templateOutputFile = @"f:\template_fill.docx";
            string sourceFile = @"f:\1 - 副本.doc";
            string sourceFile1 = @"f:\1.doc";
            string outputFile = @"f:\1.pdf";
            string outputFile1 = @"f:\11.pdf";
            string pdfWaterOutputFile = @"f:\2.pdf";
            string pdfImageOutputFile = @"f:\3.pdf";
            string qfzSourceFile = @"f:\9.doc";
            string qfzOutputFile = @"f:\9_N.doc";
            string imageFile = @"f:\9.png";
            string down = "https://pps.mingyuanyun.com/api/yf-download-file?fId=293002867493";
            //word2pdf(sourceFile1, outputFile);
            //ExportPdf(sourceFile1, outputFile, imageFile);
            //HttpDldFile.Download(down, sourceFile1);
            aspose.AsposeWord2pdf(sourceFile1, outputFile1);
        }

        public static bool ExportPdf(string fileName, string outputFileName, string imageFile)
        {
            if (string.IsNullOrEmpty(fileName) || string.IsNullOrEmpty(outputFileName))
            {
                return false;
            }

            if (!File.Exists(fileName) && !File.Exists(fileName + "x"))
            {
                return false;
            }

            string extension = Path.GetExtension(fileName);
            string formatExtension = Path.GetExtension(outputFileName);
            if (string.IsNullOrEmpty(extension) || string.IsNullOrEmpty(formatExtension))
            {
                return false;
            }

            if (formatExtension != ".pdf")
            {
                return false;
            }

            switch (extension)
            {
                case ".doc":
                    return WordExportAsPdf(fileName, outputFileName, imageFile);
                case ".docx":
                    return WordExportAsPdf(fileName, outputFileName, imageFile);
                case ".mht":
                    return WordExportAsPdf(fileName, outputFileName, imageFile);
                case ".htm":
                    return WordExportAsPdf(fileName, outputFileName, imageFile);
                case ".html":
                    return WordExportAsPdf(fileName, outputFileName, imageFile);
                //case ".xls":
                //    return ExcelExportAsPdf( fileName , outputFileName );
                //case ".xlsx":
                //    return ExcelExportAsPdf( fileName , outputFileName );
                //case ".ppt":
                //    return PowerPointExportAsPdf( fileName , outputFileName );
                //case ".pptx":
                //    return PowerPointExportAsPdf( fileName , outputFileName );
                default:
                    return false;
            }
        }

        /// <summary>
        /// 转换为pdf文件，适合（.doc、.docx、.mht、.htm文件类型）
        /// </summary>
        /// <param name="fileName"></param>
        /// <param name="outputFileName"></param>
        /// <param name="waterFileName"></param>
        /// <returns></returns>
        private static bool WordExportAsPdf(string fileName, string outputFileName, string waterFileName)
        {
            bool isSucceed = false;
            Word2013.WdExportFormat fileFormat = Microsoft.Office.Interop.Word.WdExportFormat.wdExportFormatPDF;
            Word2013._Application wordApp = null;
            if (wordApp == null)
            {
                wordApp = new Word2013.Application();
            }

            Word2013._Document wordDoc = null;

            try
            {
                wordDoc = wordApp.Documents.Open(fileName);

                wordDoc.ExportAsFixedFormat(outputFileName, fileFormat);
                isSucceed = true;
            }
            finally
            {
                if (wordDoc != null)
                {
                    wordDoc.Close();
                    wordDoc = null;
                }
                if (wordApp != null)
                {
                    wordApp.Quit();
                    wordApp = null;
                }
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            return isSucceed;
        }


        //public static void word2pdf(string sourceFile, string outputFile)
        //{
        //    Spire.Doc.Document document = new Spire.Doc.Document();
        //    document.LoadFromFile(sourceFile);
        //    //Convert Word to PDF
        //    document.SaveToFile(outputFile, FileFormat.PDF);
        //    //Launch Document

        //}
    }
}
