#region << 代 码 注 释 >>
/*----------------------------------------------------------------
* 项目名称 ：WordCompare
* 项目描述 ：
* 类 名 称 ：ImageTool
* 类 描 述 ：
* 所在的域 ：DESKTOP-L6TBB0E
* 命名空间 ：WordCompare
* 机器名称 ：DESKTOP-L6TBB0E 
* CLR 版本 ：4.0.30319.42000
* 作    者 ：zhouds
* 创建时间 ：2019/3/19 9:40:15
* 更新时间 ：2019/3/19 9:40:15
* 
* Ver      负责人        变更内容            变更日期
* ──────────────────────────────────────────────────────────────
* V1.0     周冬生    	 初版                2019/3/19 9:40:15 
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

namespace WordCompare
{
    /// <summary>
    /// 图像工具类
    /// </summary>
    public class ImageTool
    {
        /// <summary>
        /// 切分图片
        /// </summary>
        /// <param name="source">源图</param>
        /// <param name="rect">切分区域</param>
        /// <param name="split_count">切分次数</param>
        /// <param name="mod">余数</param>
        /// <returns></returns>
        public static List<Image> SplitImage( Image source , Rectangle rect , int split_count , int mod )
        {
            if ( source == null || rect.IsEmpty || split_count == 0 )
            {
                return null;
            }

            List<Image> List = new List<Image>( );

            try
            {
                int x = 0;
                for ( int i = 0; i < split_count; i++ )
                {
                    x = rect.Width * i;
                    if ( i == split_count - 1 )
                    {
                        rect.Width = rect.Width + mod;
                    }

                    Bitmap bmSmall = new Bitmap( rect.Width , rect.Height , System.Drawing.Imaging.PixelFormat.Format32bppRgb );
                    using ( Graphics grSmall = Graphics.FromImage( bmSmall ) )
                    {
                        grSmall.DrawImage( source ,
                        rect ,
                        new System.Drawing.Rectangle( x , 0 , rect.Width , rect.Height ) ,
                        GraphicsUnit.Pixel );
                        grSmall.Dispose( );
                    }
                    List.Add( bmSmall );
                }

                return List;
            }
            catch ( Exception ex )
            {
                Console.WriteLine( ex.Message );
                return List;
            }
        }
    }
}
