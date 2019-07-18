#region << 代 码 注 释 >>
/*----------------------------------------------------------------
* 项目名称 ：WordCompare
* 项目描述 ：
* 类 名 称 ：WordToPDF
* 类 描 述 ：
* 所在的域 ：DESKTOP-L6TBB0E
* 命名空间 ：WordCompare
* 机器名称 ：DESKTOP-L6TBB0E 
* CLR 版本 ：4.0.30319.42000
* 作    者 ：zhouds
* 创建时间 ：2019/3/6 11:17:19
* 更新时间 ：2019/3/6 11:17:19
* 
* Ver      负责人        变更内容            变更日期
* ──────────────────────────────────────────────────────────────
* V1.0     周冬生    	 初版                2019/3/6 11:17:19 
*
* Copyright (c) 2018 MySoft Corporation. All rights reserved. 
*┌─────────────────────────────────────────────────────────────┐
*│　此技术信息为本公司机密信息，未经本公司书面同意禁止向第三方披露．                                                        │
*│　版权所有：明源云          　　　　　　　　　　　　　　                                                                  │
*└─────────────────────────────────────────────────────────────┘
//----------------------------------------------------------------*/
#endregion

using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;
using Spire.Pdf.Annotations;
using Spire.Pdf.Annotations.Appearance;
using Spire.Pdf.Graphics;
using System;
using System.Drawing;
using System.IO;
using Word2013 = Microsoft.Office.Interop.Word;

namespace WordCompare
{  
    public class InteropOffice
    {

        #region Microsoft.Office.Interop.Word

        /// <summary>
        /// 动态插入图片  骑缝章
        /// </summary>
        /// <param name="sourceFile"></param>
        /// <param name="outputFile"></param>
        public static void AddPicture(string sourceFile, string outputFile)
        {
            Word2013._Application wordApp = new Word2013.Application();
            Word2013._Document wordDoc = null;

            try
            {
                wordDoc = wordApp.Documents.Open(sourceFile);

                var stat = Word2013.WdStatistic.wdStatisticPages;
                int num = wordDoc.ComputeStatistics(stat);
                string imageFile = @"f:\9.png";
                var img = Image.FromFile(imageFile);
                int wordPage = num;
                int mod = img.Width % wordPage;

                var img_list = ImageTool.SplitImage(img, new System.Drawing.Rectangle(0, 0, img.Width / wordPage, img.Height), wordPage, mod);

                int index1 = 10000;
                foreach (var d in img_list)
                {
                    d.Save(@"f:\\" + index1++.ToString() + ".png");
                }


                int pagewidth = Convert.ToInt32(wordDoc.PageSetup.PageWidth);
                int pageheight = Convert.ToInt32(wordDoc.PageSetup.PageHeight);
                int rightmargin = Convert.ToInt32(wordDoc.PageSetup.RightMargin);
                //wordDoc.PageSetup.RightMargin = 0;
                object What = Word2013.WdGoToItem.wdGoToPage;
                object Which = Word2013.WdGoToDirection.wdGoToNext;

                int index = 0;
                if (wordDoc.Paragraphs != null && wordDoc.Paragraphs.Count > 0)
                {
                    object first = Word2013.WdGoToDirection.wdGoToFirst;
                    wordDoc.ActiveWindow.Selection.GoTo(ref What, ref first, 0); // 第二个参数可以用Nothing
                    for (int i = 0; i < num; i++)  //遍历查找
                    {
                        object LinkToFile = false;
                        object SaveWithDocument = true;
                        Microsoft.Office.Interop.Word.InlineShape Inlineshape = wordDoc.Application.ActiveDocument.InlineShapes.AddPicture(
                          $@"f:\{10000+i}.png", ref LinkToFile, ref SaveWithDocument);
                        Inlineshape.Select(); //这一步很重要 不然后面转换会报错
                        var shape = Inlineshape.ConvertToShape();


                        //var shape = wordApp.ActiveDocument.Shapes.AddPicture( $@"f:\1000{index++}.png" , ref LinkToFile , ref SaveWithDocument );
                        shape.WrapFormat.Type = Word2013.WdWrapType.wdWrapFront;
                        shape.Top = (pageheight - 146) / 2;   //设置在本页插入图片的位置（不同的文本要多试几次）
                        shape.Left =  pagewidth - rightmargin- img.Width / wordPage-15;

                        if (i < num - 1)
                        {
                            wordDoc.ActiveWindow.Selection.GoTo(ref What, ref Which, 1); // 第二个参数可以用Nothing
                        }
                    }

                    wordDoc.SaveAs2(FileName: outputFile);
                    Console.WriteLine("结束");
                }
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
        }
        /// <summary>
        /// 填充word模板中的窗体控件
        /// </summary>
        /// <param name="sourceFile"></param>
        /// <param name="outputFile"></param>
        /// <param name="propValue"></param>
        public static void OfficeFillTemplate(string sourceFile, string outputFile, string propValue)
        {
            Word2013._Application wordApp = new Word2013.Application();
            Word2013._Document wordDoc = null;

            try
            {
                wordDoc = wordApp.Documents.Open(sourceFile);
             
                foreach (Word2013.Paragraph item in wordDoc.Content.Paragraphs)
                {
                    var we = item.Range.Text;
                    //item.Range.
                    if (we.Contains("电子印章专用"))
                    {
                        object start = 0;
                        object end = 4;
                        //item.Range(ref start, ref end);
                    }
                }
                foreach (Word2013.FormField field in wordDoc.FormFields)
                {
                    if (field.Type == Word2013.WdFieldType.wdFieldFormTextInput)
                    {
                        field.Result = propValue;
                    }
                }

                wordDoc.ActiveWindow.View.FieldShading = Word2013.WdFieldShading.wdFieldShadingNever;
                wordDoc.SaveAs2(FileName: outputFile);
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
        }

        /// <summary>
        /// 进行与Word2013.Document类创建时指定的文件比较，
        /// 然后将差异显示在targetFile，并保存退出
        /// </summary>
        /// <param name="sourceFile">源文件（修改前文件）</param>
        /// <param name="targetFile">目标文件（修改后文件）</param>
        public static void CompareWordFile(string sourceFile, string targetFile)
        {
            object missing = System.Reflection.Missing.Value;
            object sFileName = sourceFile;
            var tFileName = targetFile;
            var wordApp = new Word2013.Application();
            wordApp.Caption = "CompareWordApp";
            wordApp.Visible = false;

            var wordDoc = wordApp.Documents.Open(ref sFileName, ref missing, ref missing, ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                ref missing, ref missing);

            wordDoc.TrackRevisions = true;
            wordDoc.ShowRevisions = true;
            wordDoc.PrintRevisions = true;

            object comparetarget = Word2013.WdCompareTarget.wdCompareTargetSelected;
            object addToRecentFiles = false;
            wordDoc.Compare(tFileName, ref missing, ref comparetarget, ref missing, ref missing, ref addToRecentFiles,
                ref missing, ref missing);

            int changeCount = wordApp.ActiveDocument.Revisions.Count;
            Object saveChanges = Word2013.WdSaveOptions.wdSaveChanges;
            wordDoc.Close(ref saveChanges, ref missing, ref missing);
            wordApp.Quit(ref saveChanges, ref missing, ref missing);

            Console.Read();
        }

        /// <summary>
                /// 转换生成pdf格式的文件
                /// 支持的文件类型(.doc、.docx、.mht、.htm、.xls、.xlsx、.ppt、pptx)
                /// </summary>
                /// <param name="fileName">源文件</param>
                /// <param name="outputFileName">目标文件</param>
                /// <returns></returns>
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

        ///// <summary>
        ///// 转换为pdf文件，适合（.xls、.xlsx文件类型）
        ///// </summary>
        ///// <param name="fileName"></param>
        ///// <param name="outputFileName"></param>
        ///// <returns></returns>
        //private static bool ExcelExportAsPdf( string fileName , string outputFileName )
        //{
        //    bool isSucceed = false;
        //    Excel.XlFixedFormatType fileFormat = Microsoft.Office.Interop.Excel.XlFixedFormatType.xlTypePDF;
        //    Excel.Application excelApp = null;
        //    if ( excelApp == null ) excelApp = new Excel.Application( );
        //    Excel.Workbook workBook = null;


        //    try
        //    {
        //        workBook = excelApp.Workbooks.Open( fileName );
        //        workBook.ExportAsFixedFormat( fileFormat , outputFileName );
        //        isSucceed = true;
        //    }

        //    finally
        //    {
        //        if ( workBook != null )
        //        {
        //            workBook.Close( );
        //            workBook = null;
        //        }
        //        if ( excelApp != null )
        //        {
        //            excelApp.Quit( );
        //            excelApp = null;
        //        }
        //        GC.Collect( );
        //        GC.WaitForPendingFinalizers( );
        //    }


        //    return isSucceed;
        //}


        ///// <summary>
        ///// 转换为pdf文件，适合（.ppt、pptx文件类型）
        ///// </summary>
        ///// <param name="fileName"></param>
        ///// <param name="outputFileName"></param>
        ///// <returns></returns>
        //private static bool PowerPointExportAsPdf( string fileName , string outputFileName )
        //{
        //    bool isSucceed = false;
        //    PowerPoint.PpFixedFormatType fileFormat = PowerPoint.PpFixedFormatType.ppFixedFormatTypePDF;


        //    PowerPoint.Application pptxApp = null;
        //    if ( pptxApp == null ) pptxApp = new PowerPoint.Application( );
        //    PowerPoint.Presentation presentation = null;


        //    try
        //    {
        //        presentation = pptxApp.Presentations.Open( fileName , MsoTriState.msoTrue , MsoTriState.msoFalse , MsoTriState.msoFalse );
        //        presentation.ExportAsFixedFormat( outputFileName , fileFormat );
        //        isSucceed = true;
        //    }

        //    finally
        //    {
        //        if ( presentation != null )
        //        {
        //            presentation.Close( );
        //            presentation = null;
        //        }
        //        if ( pptxApp != null )
        //        {
        //            pptxApp.Quit( );
        //            pptxApp = null;
        //        }
        //        GC.Collect( );
        //        GC.WaitForPendingFinalizers( );
        //    }


        //    return isSucceed;
        //}    

        #endregion

    }

    public class SpireOffice
    {
        #region spire.doc

        /// <summary>
        /// 填充word模板中的窗体控件
        /// </summary>
        /// <param name="sourceFile"></param>
        /// <param name="outputFile"></param>
        /// <param name="propValue"></param>
        public static void SpireFillTemplate(string sourceFile, string outputFile, string propValue)
        {
            Document doc = new Document();
            doc.LoadFromFile(sourceFile);

            //清除表单域阴影
            doc.Properties.FormFieldShading = false;

            for (int i = 0; i < doc.Sections.Count; i++)
            {
                foreach (FormField field in doc.Sections[i].Body.FormFields)
                {
                    //注意：文本域名称 == 模型中属性的 Description 值 ！！！！！！
                    //也可以： 文本域名称 == 模型中属性的 Name 值 ！！！！！！
                    if (field.Name == "name" || true)
                    {
                        //FieldType.fieldtext
                        if (field.DocumentObjectType == DocumentObjectType.TextFormField)   //文本域
                        {
                            field.Text = propValue;   //向Word模板中插入值                            
                            break;
                        }
                        else if (field.DocumentObjectType == DocumentObjectType.CheckBox)   //复选框
                        {
                            //( field as CheckBoxFormField ).Checked = ( value as bool? ).HasValue ? ( value as bool? ).Value : false;
                        }
                    }
                }
            }

            doc.SaveToFile(outputFile, FileFormat.Docx);
            doc.Close();
        }

        /// <summary>
        /// 增加图片
        /// </summary>
        /// <param name="sourceFile"></param>
        /// <param name="outputFile"></param>
        /// <param name="imgFile"></param>
        public static void AddPdfImg(string sourceFile, string outputFile, string imgFile)
        {
            Spire.Pdf.PdfDocument pdf = new Spire.Pdf.PdfDocument();
            pdf.LoadFromFile(sourceFile);

            Spire.Pdf.PdfPageBase page = pdf.Pages[0];

            PdfRubberStampAnnotation loStamp = new PdfRubberStampAnnotation(new RectangleF(new PointF(-5, -5), new SizeF(180, 180)));

            PdfAppearance loApprearance = new PdfAppearance(loStamp);

            PdfImage image = PdfImage.FromFile(imgFile);

            PdfTemplate template = new PdfTemplate(300, 300);

            template.Graphics.DrawImage(image, 200, 10);

            loApprearance.Normal = template;

            loStamp.Appearance = loApprearance;

            page.AnnotationsWidget.Add(loStamp);

            pdf.SaveToFile(outputFile);
        }

        /// <summary>
        /// 增加pdf水印
        /// </summary>
        /// <param name="sourceFile"></param>
        /// <param name="outputFile"></param>
        /// <param name="text"></param>
        public static void AddPdfWater(string sourceFile, string outputFile, string text)
        {
            Spire.Pdf.PdfDocument pdf = new Spire.Pdf.PdfDocument();
            pdf.LoadFromFile(sourceFile);

            System.Drawing.Font font = new System.Drawing.Font("宋体", 24, FontStyle.Regular);

            foreach (Spire.Pdf.PdfPageBase page in pdf.Pages)
            {
                var brush = new PdfTilingBrush(new SizeF(page.Canvas.ClientSize.Width / 2, page.Canvas.ClientSize.Height / 3));
                PdfTrueTypeFont trueTypeFont = new PdfTrueTypeFont(font, true);
                brush.Graphics.SetTransparency(0.1f);
                brush.Graphics.Save();
                brush.Graphics.TranslateTransform(brush.Size.Width / 2, brush.Size.Height / 2);
                brush.Graphics.RotateTransform(-38);
                brush.Graphics.DrawString(text, trueTypeFont, PdfBrushes.Blue, 0, 0, new PdfStringFormat(PdfTextAlignment.Center));
                brush.Graphics.Restore();
                brush.Graphics.SetTransparency(1);
                page.Canvas.DrawRectangle(brush, new RectangleF(new PointF(0, 0), page.Canvas.ClientSize));
            }

            pdf.SaveToFile(outputFile);
        }

        /// <summary>
        /// 增加水印
        /// </summary>
        /// <param name="sourceFile">源文件</param>
        /// <param name="outputFile">目标文件</param>
        /// <param name="text">水印文本</param>
        public static void AddWater(string sourceFile, string outputFile, string text)
        {
            try
            {
                Spire.Doc.Document doc = new Spire.Doc.Document();
                doc.LoadFromFile(sourceFile);

                string filename = Path.GetFileNameWithoutExtension(sourceFile);
                string extension = Path.GetExtension(sourceFile);

                TextWatermark txtWatermark = new TextWatermark();
                txtWatermark.Text = text;
                txtWatermark.FontSize = 30;
                txtWatermark.Layout = WatermarkLayout.Diagonal;

                doc.Watermark = txtWatermark;
                doc.SaveToFile(outputFile, FileFormat.PDF);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        /// <summary>
        /// 增加水印
        /// </summary>
        /// <param name="sourceFile">源文件</param>
        /// <param name="outputFile">目标文件</param>
        /// <param name="img">水印图片</param>
        public static void AddWater(string sourceFile, string outputFile, System.Drawing.Image img)
        {
            Spire.Doc.Document doc = new Spire.Doc.Document();
            doc.LoadFromFile(sourceFile);

            PictureWatermark picture = new PictureWatermark();
            picture.Scaling = 100;
            picture.Picture = img;

            doc.Watermark = picture;
            doc.SaveToFile(outputFile, FileFormat.PDF);
        }
        public static void word2pdf(string sourceFile, string outputFile)
        {
            Document document = new Document();
            document.LoadFromFile(sourceFile);
            //Convert Word to PDF
            document.SaveToFile(outputFile, FileFormat.PDF);
            //Launch Document

        }
        #endregion
    }

    public class NPOIOffice
    {
        #region NPOI

        /// <summary>
        /// 填充word模板中的窗体控件
        /// </summary>
        /// <param name="sourceFile"></param>
        /// <param name="outputFile"></param>
        /// <param name="propValue"></param>
        public static void OfficeFillTemplate(string sourceFile, string outputFile, string propValue)
        {
            using (FileStream stream = File.OpenRead(sourceFile))
            {

                NPOI.XWPF.UserModel.XWPFDocument doc = new NPOI.XWPF.UserModel.XWPFDocument(stream);
                //处理doc，代码控制编辑文档。 



                FileStream file = new FileStream(outputFile, FileMode.Create, FileAccess.Write);
                doc.Write(file);
                file.Close();
            }
        }

        #endregion
    }
}
