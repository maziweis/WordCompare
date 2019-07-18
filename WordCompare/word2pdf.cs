using Aspose.Words;
using iTextSharp.text;
using iTextSharp.text.pdf;
using Newtonsoft.Json;
using Spire.Pdf.General.Find;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
//using System.Web.Services;

namespace WordCompare
{
    public class word2pdf
    {
        //public string HelloWorld()
        //{
        //    var result = new { statue = "fail", url = "" };
        //    try
        //    {
        //        Student bll = new Student();
        //        string tempPath = @"E:\TFS72\商机管理系统\WordCompare\WordCompare\bin\Debug\";// HttpContext.Current.Server.MapPath("~/");//系统部署地址

        //        #region 创建文件夹

        //        List<string> li = new List<string>();
        //        List<Student> list = GetStudentData();
        //        foreach (Student model in list)
        //        {
        //            Document doc = new Document(tempPath + "表单模板.docx");//加载模板
        //            string fName = DateTime.Now.ToString("yyyyMMddHHmmssfff") + "_" + model.name + ".docx";
        //            string savePath = tempPath + fName;//文件保存路径
        //            //根据word中的标签赋值
        //            doc.Range.Bookmarks["NAME"].Text = "张三";
        //            doc.Range.Bookmarks["SEX"].Text = "男";
        //            doc.Range.Bookmarks["AGE"].Text = "18";

        //            doc.Save(savePath, SaveFormat.Docx);
        //            DOCConvertToPDF(tempPath + "表单模板.docx", savePath.Replace(".docx", ".pdf"));
        //            li.Add(savePath.Replace(".doc", ".pdf"));//生成附件的list集合
        //        }
        //        #endregion 创建文件夹
        //        //保存文件路径
        //        string sPath = tempPath + "_导出基本信息.doc";
        //        string pdf_path = sPath.Replace(".doc", ".pdf");
        //        DOCConvertToPDF(tempPath + "空模板.doc", pdf_path);
        //        MergePDFFiles(li.ToArray(), pdf_path);//合并文件

        //        //删除转换后的文件
        //        //foreach (PdfModel model in mergeFile_temp)
        //        //{
        //        //    model.pdf.Dispose();
        //        //    File.Delete(model.url);//删除临时文件
        //        //    File.Delete(model.url.Replace(".pdf", ".word"));
        //        //}

        //        result = new { statue = "success", url = pdf_path };
        //    }
        //    catch (Exception ex)
        //    {
        //    }
        //    return JsonConvert.SerializeObject(result); ;
        //}

        ///// <summary> 合并PDF </summary>  
        ///// <param name="fileList">PDF文件集合</param>  
        ///// <param name="outMergeFile">合并文件名</param>  
        ///// 
        //public static void MergePDFFiles(string[] fileList, string outMergeFile)
        //{
        //    PdfReader reader;
        //    iTextSharp.text.Rectangle re;
        //    PdfDictionary pd;
        //    //List<PdfReader> readerList = new List<PdfReader>();//记录合并PDF集合  
        //    iTextSharp.text.Document document = new iTextSharp.text.Document();
        //    PdfWriter writer = PdfWriter.GetInstance(document, new FileStream(outMergeFile, FileMode.Create));
        //    document.Open();

        //    PdfContentByte cb = writer.DirectContent;
        //    PdfImportedPage newPage;
        //    for (int i = 0; i < fileList.Length; i++)
        //    {
        //        if (!string.IsNullOrEmpty(fileList[i]))
        //        {
        //            reader = new PdfReader(fileList[i]);
        //            PdfModel model = new PdfModel();
        //            model.pdf = reader;
        //            model.url = fileList[i];
        //            //mergeFile_temp.Add(model);

        //            int iPageNum = reader.NumberOfPages;
        //            for (int j = 1; j <= iPageNum; j++)
        //            {
        //                //获取Reader的pdf页的打印方向
        //                re = reader.GetPageSize(reader.GetPageN(j));
        //                //设置合并pdf的打印方向
        //                document.SetPageSize(re);
        //                document.NewPage();
        //                newPage = writer.GetImportedPage(reader, j);
        //                cb.AddTemplate(newPage, 0, 0);
        //            }
        //            //reader.Dispose();
        //            //readerList.Add(reader);
        //        }
        //    }
        //    document.Close();
        //    //Process.Start(outMergeFile);//预览
        //}


        //private static bool DOCConvertToPDF(string sourcePath, string targetPath)
        //{
        //    bool result;
        //    try
        //    {
        //        Aspose.Words.Document document = new Aspose.Words.Document(sourcePath);
        //        document.Save(targetPath, Aspose.Words.SaveFormat.Pdf);
        //        result = true;
        //    }
        //    catch (Exception ex)
        //    {
        //        result = false;
        //    }
        //    return result;
        //}

        //public static List<Student> GetStudentData()
        //{

        //    List<Student> list = new List<Student>();

        //    for (int index = 0; index < 3; index++)
        //    {
        //        Student model = new Student();
        //        model.name = "张三_" + index;
        //        model.sex = "男-" + index;
        //        model.age = index.ToString();
        //        list.Add(model);
        //    }
        //    return list;
        //}
        ////访问实体类
        //public class Student
        //{
        //    /// <summary>
        //    /// 姓名
        //    /// </summary>
        //    public string name
        //    {
        //        set;
        //        get;
        //    }
        //    /// <summary>
        //    /// 性别
        //    /// </summary>
        //    public string sex
        //    {
        //        set;
        //        get;
        //    }
        //    /// <summary>
        //    /// 年龄
        //    /// </summary>
        //    public string age
        //    {
        //        set;
        //        get;
        //    }

        //}
        //public class PdfModel
        //{
        //    public PdfReader pdf
        //    {
        //        set;
        //        get;
        //    }
        //    public string url
        //    {
        //        set;
        //        get;
        //    }
        //}


        /// <summary>
        /// 添加倾斜水印
        /// </summary>
        /// <param name="pdf">pdf文件流</param>
        /// <param name="waterMarkName">水印字符串</param>
        /// <param name="width">页面宽度</param>
        /// <param name="height">页面高度</param>
        public MemoryStream SetWaterMark(string pdf, string waterMarkName, float width, float height)
        {
            try
            {
                int fontSize = 50;//设置字体大小
                int span = 40;//设置垂直位移
                MemoryStream outStream = new MemoryStream();
                PdfReader pdfReader = new PdfReader(pdf);
                PdfStamper pdfStamper = new PdfStamper(pdfReader, outStream);
                pdfStamper.Writer.CloseStream = false;
                int total = pdfReader.NumberOfPages + 1;
                PdfContentByte content;
                //BaseFont font = BaseFont.CreateFont(@"C:\WINDOWS\Fonts\STCAIYUN.TTF", BaseFont.IDENTITY_H, BaseFont.EMBEDDED);//华文云彩字体
                BaseFont font = BaseFont.CreateFont();//华文云彩字体
                PdfGState gs = new PdfGState();
                gs.FillOpacity = 0.15f;//透明度
                int waterMarkNameLenth = waterMarkName.Length;
                char c;
                int rise = 0;
                string spanString = " ";//水平位移
                for (int i = 1; i < total; i++)
                {
                  var size=  pdfReader.GetPageSize(i);
                       rise = waterMarkNameLenth * span;
                    content = pdfStamper.GetOverContent(i);//在内容上方加水印
                                                           //content = pdfStamper.GetUnderContent(i);//在内容下方加水印
                    content.SetGState(gs);
                    content.BeginText();
                    content.SetColorFill(BaseColor.GREEN);
                    content.SetFontAndSize(font, fontSize);
                    int heightNumbert = (int)Math.Ceiling((decimal)height / (decimal)rise);//垂直重复的次数，进一发
                    int panleWith = (fontSize + span) * waterMarkNameLenth;
                    int widthNumber = (int)Math.Ceiling((decimal)width / (decimal)panleWith);//水平重复次数
                                                                                             // 设置水印文字字体倾斜 开始 
                    for (int w = 0; w < widthNumber; w++)
                    {
                        for (int h = 1; h <= heightNumbert; h++)
                        {
                            int yleng = rise * h;
                            content.SetTextMatrix(w * panleWith, yleng);//x,y设置水印开始的绝对左边，以左下角为x，y轴的起点
                            for (int k = 0; k < waterMarkNameLenth; k++)
                            {
                                content.SetTextRise(yleng);//指定的y轴值处添加
                                c = waterMarkName[k];
                                content.ShowText(c + spanString);
                                yleng -= span;
                            }
                        }
                    }
                    content.EndText();
                }
                if (pdfStamper != null)
                    pdfStamper.Close();

                if (pdfReader != null)
                    pdfReader.Close();
                return outStream;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }



        #region 添加普通偏转角度文字水印
        /// <summary>
        /// 添加普通偏转角度文字水印 https://www.cnblogs.com/loyung/p/6879917.html
        /// </summary>
        /// <param name="inputfilepath">需要添加水印的pdf文件</param>
        /// <param name="outputfilepath">添加水印后输出的pdf文件</param>
        /// <param name="waterMarkName">水印内容</param>
        public static void setWatermark(string inputfilepath, string outputfilepath, string waterMarkName)
        {
          
            
            PdfReader pdfReader = null;
            PdfStamper pdfStamper = null;
            try
            {
                if (File.Exists(outputfilepath))
                {
                    File.Delete(outputfilepath);
                }

                pdfReader = new PdfReader(inputfilepath);
                pdfStamper = new PdfStamper(pdfReader, new FileStream(outputfilepath, FileMode.OpenOrCreate));

                int total = pdfReader.NumberOfPages + 1;
                iTextSharp.text.Rectangle psize = pdfReader.GetPageSize(1);
                
                float width = psize.Width;
                float height = psize.Height;
                PdfContentByte content;
                BaseFont font = BaseFont.CreateFont(@"C:\WINDOWS\Fonts\SIMFANG.TTF", BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
                PdfGState gs = new PdfGState();
                for (int i = 1; i < total; i++)
                {
                    content = pdfStamper.GetOverContent(i);//在内容上方加水印
                    
                    //content = pdfStamper.GetUnderContent(i);//在内容下方加水印
                    //透明度
                    gs.FillOpacity = 0.3f;
                    content.SetGState(gs);
                    //content.SetGrayFill(0.3f);
                    //开始写入文本
                    content.BeginText();
                    content.SetColorFill(BaseColor.LIGHT_GRAY);
                    content.SetFontAndSize(font, 100);
                    content.SetTextMatrix(0, 0);
                    content.ShowTextAligned(Element.ALIGN_CENTER, waterMarkName, width / 2 - 50, height / 2 - 50, 55);
                    //content.SetColorFill(BaseColor.BLACK);
                    //content.SetFontAndSize(font, 8);
                    //content.ShowTextAligned(Element.ALIGN_CENTER, waterMarkName, 0, 0, 0);
                    content.EndText();
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {

                if (pdfStamper != null)
                    pdfStamper.Close();

                if (pdfReader != null)
                    pdfReader.Close();
            }
        }
        #endregion
    }
}
