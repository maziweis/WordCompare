using iTextSharp.text;
using iTextSharp.text.pdf;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace WordCompare
{
    public class itextsharp
    {
        #region 图片插入PDF方法


        /// <summary>
        /// pdf添加图片
        /// </summary>
        /// <param name="imglist">图片的list</param>
        public static void CreatePdf(string pdfpath,string pdfImageOutputFile, string imagepath)
        {
            using (Stream inputPdfStream = new FileStream(pdfpath, FileMode.Open, FileAccess.Read, FileShare.Read))
            using (Stream outputPdfStream = new FileStream(pdfImageOutputFile, FileMode.Create, FileAccess.Write, FileShare.None))
            using (Stream inputImageStream = new FileStream(imagepath, FileMode.Open, FileAccess.Read, FileShare.Read))
            {
                var reader = new PdfReader(inputPdfStream);//读取原有pdf
                var num = reader.NumberOfPages;
                var stamper = new PdfStamper(reader, outputPdfStream);
                var pdfContentByte = stamper.GetOverContent(num);//获取内容 
                iTextSharp.text.Image image = iTextSharp.text.Image.GetInstance(inputImageStream);//获取图片
                image.ScalePercent(20);//设置图片比例
                image.SetAbsolutePosition(30, 30);//设置图片的绝对位置
                pdfContentByte.AddImage(image);
                stamper.Close();
            }
            /////分割list
            //string[] imgs = imglist.Split(',');

            //string pdfpath = Server.MapPath("pdf");

            //string imagepath = Server.MapPath("Image");

            //string pdfpath = @"G:\MyWeb\Web学习\Windows\LiveProject\LiveProject\Images\"; //文件路径
            //string imagepath = @"G:\MyWeb\Web学习\Windows\LiveProject\LiveProject\Images\";

            ///实例化一个doc 对象
            //Document doc = new Document();
            //try
            //{
            //    ///创建一个pdf 对象
            //    PdfWriter.GetInstance(doc, new FileStream(pdfpath , FileMode.Open));
            //    PdfReader pdfReader = new PdfReader(pdfpath);
                
            //    //打开文件
            //    doc.Open();


            //    ///向文件中添加单个图片
            //    iTextSharp.text.Image image = iTextSharp.text.Image.GetInstance(imagepath);

            //    image.ScaleToFit(520, 800);
            //    doc.Add(image);

            //    ///向文件中循环添加图片
            //    //iTextSharp.text.Image image;
            //    //for (int i = 0; i < imgs.Length; i++)
            //    //{
            //    //    image = iTextSharp.text.Image.GetInstance(imagepath + imgs[i].ToString());

            //    //    image.ScaleToFit(520, 800);
            //    //    doc.NewPage();
            //    //    doc.Add(image);
            //    //}

            //}

          
            //catch (Exception ex)
            //{

              
            //}

            //finally
            //{

            //    doc.Close();

            //}

        }

        #endregion
    }
}
