using System;
using System.Drawing;
using System.IO;

namespace WordCompare
{
    static class Program
    {
        /// <summary>
        /// 应用程序的主入口点。
        /// </summary>
        [STAThread]
        static void Main()
        {
            //QRCodeBuild.BuildCode();
            Console.WriteLine( "1 转换word文档到pdf(将读取f盘下1.docx或1.doc，生成1.pdf)" );
            Console.WriteLine( "2 读取pdf文件写入水印(将读取f盘下1.pdf，生成2.pdf)" );
            Console.WriteLine( "3 读取pdf文件插入图片(将读取f盘下2.pdf，生成3.pdf)" );
            Console.WriteLine( "4 读取word文档，并用spire填充文本域" );
            Console.WriteLine( "5 读取word文档，并用office填充文本域" );
            Console.WriteLine( "6 百度识别ai图片,读取f盘下ai.png" );
            Console.WriteLine( "7 阿里识别ai图片,读取f盘下ai.png" );
            Console.WriteLine( "8 切分图片,9.png" );
            Console.WriteLine( "9 动态插入图片,100**.png" );
            Console.WriteLine( "请输入1、2 水印文本、3 ，回车" );
            var type = "3";// Console.ReadLine( );

            string templateFile = @"f:\template.docx";
            string templateOutputFile = @"f:\template_fill.docx";
            string sourceFile = @"f:\1 - 副本.doc"; string sourceFile1 = @"f:\101.doc";
            string outputFile = @"f:\1.pdf";
            string outputFile1 = @"f:\11.pdf";
            string pdfWaterOutputFile = @"f:\2.pdf";
            string pdfImageOutputFile = @"f:\3.pdf";
            string qfzSourceFile = @"f:\明源云空间SAAS开发商服务合同模板.docx";
            //qfzSourceFile = @"f:\明源云客硬件产品合同模板.docx";
            //qfzSourceFile = @"f:\明源云链SAAS开发商服务合同模板.docx";
            //qfzSourceFile = @"f:\明源云客SAAS开发商服务合同模板.docx";
            string qfzOutputFile = @"f:\9_N.docx";
            string imageFile = @"f:\qr.jpg";
            //aspose.AsposeWord2pdf(sourceFile1, outputFile1);
            //SpireOffice.word2pdf(sourceFile1, outputFile1);
            while ( type != "exit" )
            {
                if (type == "1")
                {
                    if (InteropOffice.ExportPdf(sourceFile1, outputFile, imageFile))
                    {
                        Console.WriteLine(outputFile + "生成成功");
                    }
                }
                else if (type.StartsWith("2") && type.IndexOf(" ") >= 0)
                {
                    var types = type.Split(' ');

                    if (types.Length == 1)
                    {
                        Console.WriteLine("无效水印文字");
                    }
                    else
                    {
                        string text = types[1];
                        SpireOffice.AddPdfWater(outputFile, pdfWaterOutputFile, text);
                    }
                }
                else if (type.StartsWith("3"))
                {
                    itextsharp.CreatePdf(outputFile, pdfImageOutputFile, imageFile);
                    //SpireOffice.AddPdfImg(outputFile, pdfImageOutputFile, imageFile);
                }
                else if (type.StartsWith("4"))
                {
                    SpireOffice.SpireFillTemplate(templateFile, templateOutputFile, "zds");
                }
                else if (type.StartsWith("5"))
                {
                    InteropOffice.OfficeFillTemplate(sourceFile, templateOutputFile, "zds");
                }
                else if (type.StartsWith("6"))
                {
                    BaiduOcr.GeneralBasicDemo();
                }
                else if (type.StartsWith("7"))
                {
                    AliOcr.Ocr();
                }
                else if (type.StartsWith("8"))
                {
                    var img = Image.FromFile(imageFile);
                    int wordPage = 8;
                    int mod = img.Width % wordPage;

                    var img_list = ImageTool.SplitImage(img, new System.Drawing.Rectangle(0, 0, img.Width / wordPage, img.Height), wordPage, mod);

                    int index = 10000;
                    foreach (var d in img_list)
                    {
                        d.Save(@"f:\\" + index++.ToString() + ".png");
                    }
                }
                else if (type.StartsWith("9"))
                {
                    InteropOffice.AddPicture(qfzSourceFile, qfzOutputFile);
                }
                else if (type.StartsWith("10"))
                {
                    word2pdf.setWatermark(outputFile, outputFile1, "21345");
                }
                else
                {
                    Console.WriteLine("无效命令");
                }
                type = Console.ReadLine( );
            }


        }


    }
}
