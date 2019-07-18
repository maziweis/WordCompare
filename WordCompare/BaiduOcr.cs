using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Baidu.Aip.Ocr;

namespace WordCompare
{
    public class BaiduOcr
    {
        // 设置APPID/AK/SK
        static string APP_ID = "15726558";
        static string API_KEY = "lMquXQ5oxxgh9p4dx1alVEIN";
        static string SECRET_KEY = "4sQbRYlbwjdIGmI1EfmpgetjgTM9ryp3";
        static Baidu.Aip.Ocr.Ocr client = new Baidu.Aip.Ocr.Ocr( API_KEY , SECRET_KEY );

        static BaiduOcr()
        {            
            client.Timeout = 60000;  // 修改超时时间
        }

        public static void GeneralBasicDemo()
        {
            var image = File.ReadAllBytes( @"f:\ai.png" );
            // 调用通用文字识别, 图片参数为本地图片，可能会抛出网络等异常，请使用try/catch捕获
            var result = client.GeneralBasic( image );
            Console.WriteLine( result );
            // 如果有可选参数
            var options = new Dictionary<string , object>{
        {"language_type", "CHN_ENG"},
        {"detect_direction", "true"},
        {"detect_language", "true"},
        {"probability", "true"}
    };
            // 带参数调用通用文字识别, 图片参数为本地图片
            result = client.GeneralBasic( image , options );
            Console.WriteLine( result );
        }
    }
}