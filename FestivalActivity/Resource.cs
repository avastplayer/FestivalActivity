using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Media.Imaging;
using System.Xml;

namespace FestivalActivity
{
    /// <summary>
    /// App.xaml 的交互逻辑
    /// </summary>
    public class Resource
    {
        public static string FilePath { get; private set; }
        public static string IconPath { get; private set; }
        public List<string> IconNameList { get; private set; }
        public DataTable ExcelDataTable { get; set; }

        public DataView ExcelDataItemsSource { get; private set; }
        public BitmapSource[] MainBackSource { get; private set; }
        public BitmapSource[] CommonCaseSource { get; private set; }
        public static BitmapSource[] BackImageSource { get; private set; }
        public BitmapSource CloseButtonSource { get; private set; }
        public BitmapSource TopLeftCornerSource { get; private set; }
        public BitmapSource TopRightCornerSource { get; private set; }
        public BitmapSource BottomLeftCornerSource { get; private set; }
        public BitmapSource BottomRightCornerSource { get; private set; }
        public BitmapSource PatternSource { get; private set; }
        public static BitmapSource IconImageBorderSource { get; private set; }
        public static BitmapSource ButtonSource { get; private set; }
        public BitmapSource LiMoChouSource { get; private set; }
        public BitmapSource[] CommonCaseSellSource { get; private set; }
        public BitmapSource[] CommonCase7Source { get; private set; }

        public Resource(bool isLoaded)
        {
            if (isLoaded) return;
            InitializeConfig();
            InitializeResource();
        }

        public void FreshConfig()
        {
            InitializeConfig();
        }

        public void WriteToExcel()
        {
            ExcelHelper excelHelper = new ExcelHelper(FilePath);
            excelHelper.DataTableToExcel(ExcelDataTable,"sheet1");
        }

        private static readonly Dictionary<string, Type> FieldMap = new Dictionary<string, Type>
        {
            ["id"] = typeof(int),
            ["任务名称"] = typeof(string),
            ["图标"] = typeof(string),
            ["角标"] = typeof(string),
            ["开启显示时间"] = typeof(string),
            ["结束显示时间"] = typeof(string),
            ["任务时间"] = typeof(string),
            ["活动类型"] = typeof(int),
            ["开启等级"] = typeof(int),
            ["等级"] = typeof(int),
            ["等级2"] = typeof(int),
            ["星期"] = typeof(int),
            ["类别"] = typeof(int),
            ["mapid"] = typeof(int),
            ["xpos"] = typeof(int),
            ["ypos"] = typeof(int),
            ["说明"] = typeof(string),
            ["排序"] = typeof(int)
        };

        private void InitializeConfig()
        {
            FilePath = GetAppConfig(nameof(FilePath));
            IconPath = GetAppConfig(nameof(IconPath));

            if (!Directory.Exists(IconPath))
            {
                MessageBox.Show($"\"{IconPath}\"未找到，请修改FestivalActivity.exe.config中的文件夹路径！");
                return;
            }

            if (!Directory.Exists(FilePath))
            {
                MessageBox.Show($"\"{FilePath}\"未找到，请修改FestivalActivity.exe.config中的文件夹路径！");
                return;
            }

            FilePath = FilePath + @"\j节日活动\c春节活动入口.xlsx";

            ExcelHelper excelHelper = new ExcelHelper(FilePath);
            ExcelDataTable = excelHelper.ExcelToDataTable("Sheet1", CreateTempletDataTable());

            ExcelDataItemsSource = ExcelDataTable.DefaultView;
        }

        private static Type GeFieldType(string fieldName) => FieldMap[fieldName];

        private static DataTable CreateTempletDataTable()
        {
            DataTable templetDataTable = new DataTable();

            foreach (string fieldName in FieldMap.Keys)
            {
                DataColumn column = new DataColumn(fieldName, GeFieldType(fieldName));
                templetDataTable.Columns.Add(column);
            }

            return templetDataTable;
        }

        private void InitializeResource()
        {
            
            MainBackSource = new BitmapSource[9];
            CommonCaseSource = new BitmapSource[9];
            BackImageSource = new BitmapSource[9];
            CommonCaseSellSource = new BitmapSource[9];//InfoCell边框
            CommonCase7Source = new BitmapSource[9];
            for (int i = 0; i < 9; i++)
            {
                MainBackSource[i] = SetImage(@"\component2", "MainBack", new Rectangle(i % 3 * (195 - 45 * (i % 3)), i / 3 * (195 - 45 * (i / 3)), i % 3 == 1 ? 60 : 150, i / 3 == 1 ? 60 : 150));
                CommonCaseSource[i] = SetImage(@"\component16", "commoncase_" + (i + 1));
                BackImageSource[i] = SetImage(@"\component17", "commoncaseteamN_" + (i + 1));
                CommonCaseSellSource[i] = SetImage(@"\component7", "commoncasesell", new Rectangle(i % 3 * (12 * (i % 3) + 18), i / 3 * (-8 * (i / 3) + 38), i % 3 == 1 ? 54 : 30, i / 3 == 1 ? 14 : 30));
                CommonCase7Source[i] = SetImage(@"\component5", "SmallBack", new Rectangle(i % 3 * (17 * (i % 3) + 13), i / 3 * (19 * (i / 3) + 11), i % 3 == 1 ? 64 : 30, i / 3 == 1 ? 68 : 30));
            }
            CloseButtonSource = SetImage(@"\component1", "CloseNormal");
            TopLeftCornerSource = SetImage(@"\MainControl49", "topleft");
            TopRightCornerSource = SetImage(@"\MainControl49", "topright");
            BottomLeftCornerSource = SetImage(@"\MainControl49", "botleft");
            BottomRightCornerSource = SetImage(@"\MainControl49", "botright");
            PatternSource = SetImage(@"\component2", "Pattern");
            IconImageBorderSource = SetImage(@"\BaseControl", "ItemInCell");
            ButtonSource = SetImage(@"\component4", "ButtonNormal");
            LiMoChouSource = SetImage(@"\MainControl9", "1meinv");//InfoCell李莫愁
        }

        public BitmapSource SetImage(string fileName, string iconName)
        {
            string iconPath = fileName.Contains("itemicon") ? (IconPath + "\\" + "item") : (IconPath + "\\" + fileName);
            IconNameList = GetIconName(iconPath);

            //当查找到图片时
            int findIconNumer = 0;
            foreach (var element in IconNameList)
            {
                if (!element.Contains(iconName)) continue;
                //加上文件后缀
                iconName = element;
                findIconNumer++;
            }
            if (findIconNumer == 0) return null;

            string soursePath = iconPath + "\\" + iconName;
            BitmapSource bitmapSource;
            if (iconName.Contains(".tga") || iconName.Contains(".TGA"))
            {
                Bitmap tga = Paloma.TargaImage.LoadTargaImage(soursePath);
                bitmapSource = System.Windows.Interop.Imaging.CreateBitmapSourceFromHBitmap(tga.GetHbitmap(),
                    IntPtr.Zero, Int32Rect.Empty,
                    BitmapSizeOptions.FromEmptyOptions());
            }
            else
            {
                bitmapSource = new BitmapImage(new Uri(soursePath, UriKind.Absolute));
            }
            return bitmapSource;
        }

        public BitmapSource SetImage(string fileName, string iconName, Rectangle rect)
        {
            string iconPath = fileName.Contains("itemicon") ? (IconPath + "\\" + "item") : (IconPath + "\\" + fileName);
            IconNameList = GetIconName(iconPath);

            //当查找到图片时
            int findIconNumer = 0;
            foreach (var element in IconNameList)
            {
                if (!element.Contains(iconName)) continue;
                //加上文件后缀
                iconName = element;
                findIconNumer++;
            }
            if (findIconNumer == 0) return null;

            string soursePath = iconPath + "\\" + iconName;
            Bitmap image;
            if (iconName.Contains(".tga") || iconName.Contains(".TGA"))
            {
                //按照坐标和尺寸裁切
                image = CutImage(Paloma.TargaImage.LoadTargaImage(soursePath), rect);
            }
            else
            {
                Bitmap originalImage = new Bitmap(soursePath);
                //按照坐标和尺寸裁切
                image = CutImage(originalImage, rect);
            }

            BitmapSource bitmapSource = System.Windows.Interop.Imaging.CreateBitmapSourceFromHBitmap(image.GetHbitmap(),
                IntPtr.Zero, Int32Rect.Empty,
                BitmapSizeOptions.FromEmptyOptions());
            return bitmapSource;
        }

        private static Bitmap CutImage(Image img, Rectangle rect)
        {
            Bitmap b = new Bitmap(rect.Width, rect.Height, PixelFormat.Format32bppArgb);
            Graphics g = Graphics.FromImage(b);
            g.DrawImage(img, 0, 0, rect, GraphicsUnit.Pixel);
            g.Dispose();
            return b;
        }

        private static string GetAppConfig(string strKey)
        {
            string file = Process.GetCurrentProcess().MainModule.FileName;
            Configuration config = ConfigurationManager.OpenExeConfiguration(file);
            return config.AppSettings.Settings.AllKeys.Any(key => key == strKey) ? config.AppSettings.Settings[strKey].Value : null;
        }

        private static List<string> GetIconName(string iconPath)
        {
            var files = Directory.GetFiles(iconPath);
            return files.Select(Path.GetFileName).ToList();
        }
    }
}