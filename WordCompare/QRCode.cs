using QRCoder;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;

namespace WordCompare
{
    public class QRCodeBuild
    {
        public static void BuildCode()
        {
            // 生成二维码的内容
            string strCode = "http://www.baidu.com";
            QRCodeGenerator qrGenerator = new QRCoder.QRCodeGenerator();
            QRCodeData qrCodeData = qrGenerator.CreateQrCode(strCode, QRCodeGenerator.ECCLevel.Q);
            QRCode qrcode = new QRCode(qrCodeData);
            string filePath = "F:/qr.jpg";
            // qrcode.GetGraphic 方法可参考最下发“补充说明”
            Bitmap qrCodeImage = qrcode.GetGraphic(5, Color.Black, Color.White, null, 15, 6, false);
            //MemoryStream ms = new MemoryStream();
            //qrCodeImage.Save(ms, ImageFormat.Jpeg);
            qrCodeImage.Save(filePath);
            // 如果想保存图片 可使用  qrCodeImage.Save(filePath);

        }
    }
}
