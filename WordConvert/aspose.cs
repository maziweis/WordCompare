using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Aspose.Words;
namespace WordConvert
{
    public class aspose
    {
        public static void AsposeWord2pdf(string sourceFile, string outputFile)
        {
            if (File.Exists(outputFile))
            {
                File.Delete(outputFile);
            }
            Document doc = new Document(sourceFile);
            //保存为PDF文件，此处的SaveFormat支持很多种格式，如图片，epub,rtf 等等         
            doc.Save(outputFile, SaveFormat.Pdf);

        }
    }
}

