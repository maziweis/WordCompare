using Spire.Pdf;
using Spire.Pdf.General.Find;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace WordCompare
{
  public  class spire
    {
        public static void OfficeFillTemplate(string sourceFile, string outputFile, string propValue)
        {
            PdfDocument pdf = new PdfDocument();
            pdf.LoadFromFile(sourceFile);
            PdfTextFind[] AllMatchedText = null;
            //遍历页面
            foreach (PdfPageBase page in pdf.Pages)
            {
                //根据关键词查找页面中匹配的文本
                AllMatchedText = page.FindText("电子印章专用").Finds;
                foreach (PdfTextFind text in AllMatchedText)
                {
                    //给匹配的文本设置背景颜色
                    
                }

            }
            pdf.SaveToFile(outputFile);
        }
        }
}
