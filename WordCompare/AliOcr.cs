#region << 代 码 注 释 >>
/*----------------------------------------------------------------
* 项目名称 ：WordCompare
* 项目描述 ：
* 类 名 称 ：AliOcr
* 类 描 述 ：
* 所在的域 ：DESKTOP-L6TBB0E
* 命名空间 ：WordCompare
* 机器名称 ：DESKTOP-L6TBB0E 
* CLR 版本 ：4.0.30319.42000
* 作    者 ：zhouds
* 创建时间 ：2019/3/12 11:09:39
* 更新时间 ：2019/3/12 11:09:39
* 
* Ver      负责人        变更内容            变更日期
* ──────────────────────────────────────────────────────────────
* V1.0     周冬生    	 初版                2019/3/12 11:09:39 
*
* Copyright (c) 2018 MySoft Corporation. All rights reserved. 
*┌─────────────────────────────────────────────────────────────┐
*│　此技术信息为本公司机密信息，未经本公司书面同意禁止向第三方披露．                                                        │
*│　版权所有：明源云          　　　　　　　　　　　　　　                                                                  │
*└─────────────────────────────────────────────────────────────┘
//----------------------------------------------------------------*/
#endregion

using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Security;
using System.Security.Cryptography.X509Certificates;
using System.Text;

namespace WordCompare
{
    /// <summary>
    /// 类用途描述
    /// </summary>
    public static class AliOcr
    {
        // 设置APPID/AK/SK
        static string APP_ID = "25829048";
        static string APP_CODE = "5a303d377b794b5b8c011b8d3e3d11a4";
        static string SECRET_KEY = "8a0370f8a86d8d2f570266b71277dab5";

        static string host = "http://ocrapi-advanced.taobao.com";
        static string path = "/ocrservice/advanced";
        static string method = "POST";

        public static void Ocr()
        {           
            string querys = "";
            string base64 = ImageToBase64( @"f:\ai.png" );

            string bodys = $"{{\"img\":\"{base64}\",\"url\":\"\",\"prob\":false,\"charInfo\":false,\"rotate\":false,\"table\":false}}";
            string url = host + path;
            HttpWebRequest httpRequest = null;
            HttpWebResponse httpResponse = null;

            if ( 0 < querys.Length )
            {
                url = url + "?" + querys;
            }

            if ( host.Contains( "https://" ) )
            {
                ServicePointManager.ServerCertificateValidationCallback = new RemoteCertificateValidationCallback( CheckValidationResult );
                httpRequest = ( HttpWebRequest )WebRequest.CreateDefault( new Uri( url ) );
            }
            else
            {
                httpRequest = ( HttpWebRequest )WebRequest.Create( url );
            }
            httpRequest.Method = method;
            httpRequest.Headers.Add( "Authorization" , "APPCODE " + APP_CODE );
            //根据API的要求，定义相对应的Content-Type
            httpRequest.ContentType = "application/json; charset=UTF-8";
            if ( 0 < bodys.Length )
            {
                byte[ ] data = Encoding.UTF8.GetBytes( bodys );
                using ( Stream stream = httpRequest.GetRequestStream( ) )
                {
                    stream.Write( data , 0 , data.Length );
                }
            }
            try
            {
                httpResponse = ( HttpWebResponse )httpRequest.GetResponse( );
            }
            catch ( WebException ex )
            {
                httpResponse = ( HttpWebResponse )ex.Response;
            }

            Console.WriteLine( httpResponse.StatusCode );
            Console.WriteLine( httpResponse.Method );
            Console.WriteLine( httpResponse.Headers );
            Stream st = httpResponse.GetResponseStream( );
            StreamReader reader = new StreamReader( st , Encoding.GetEncoding( "utf-8" ) );
            Console.WriteLine( reader.ReadToEnd( ) );
            Console.WriteLine( "\n" );
        }

        public static bool CheckValidationResult( object sender , X509Certificate certificate , X509Chain chain , SslPolicyErrors errors )
        {
            return true;
        }

        /// <summary>
        /// Image 转成 base64
        /// </summary>
        /// <param name="fileFullName"></param>
        public static string ImageToBase64( string fileFullName )
        {
            try
            {
                Bitmap bmp = new Bitmap( fileFullName );
                MemoryStream ms = new MemoryStream( );
                bmp.Save( ms , System.Drawing.Imaging.ImageFormat.Jpeg );
                byte[ ] arr = new byte[ ms.Length ]; ms.Position = 0;
                ms.Read( arr , 0 , ( int )ms.Length ); ms.Close( );
                return Convert.ToBase64String( arr );
            }
            catch ( Exception ex )
            {
                return null;
            }
        }
    }
}
