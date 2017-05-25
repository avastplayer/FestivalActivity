using System;
using System.Data;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;

namespace FestivalActivity
{
    /// <summary>
    /// Cell.xaml 的交互逻辑
    /// </summary>
    public partial class Cell
    {
        public DataTable ExcelDataTable { get; set; }
        public int Row { get; set; }

        public Cell(DataTable datatable, int row)
        {
            ExcelDataTable = datatable;
            Row = row;
            InitializeComponent();
            InitializeCell(Row);
        }

        public void InitializeCell(int row)
        {
            Resource resource = new Resource(true);

            //显示背景
            Image[] backImage = { BackImage1, BackImage2, BackImage3, BackImage4, BackImage5, BackImage6, BackImage7, BackImage8, BackImage9 };
            for (int i = 0; i < 9; i++)
            {
                backImage[i].Source = Resource.BackImageSource[i];
            }

            //显示图表边框
            IconImageBorder.Source = Resource.IconImageBorderSource;

            //显示图标
            string iconPath = ExcelDataTable.Rows[row]["图标"].ToString();
            string[] iconArray = iconPath.Split(new[] { "set:", "image:" }, StringSplitOptions.RemoveEmptyEntries);
            string iconFileName = iconArray[0];
            string iconName = iconArray[1];

            IconImage.Source = resource.SetImage(iconFileName, iconName);

            //显示角标
            string cornerPath = ExcelDataTable.Rows[row]["角标"].ToString();
            if (cornerPath != "")
            {
                string[] cornerArray = cornerPath.Split(new[] { "set:", "image:" }, StringSplitOptions.RemoveEmptyEntries);
                string cornerFileName = cornerArray[0];
                string cornerName = cornerArray[1];

                CornerImage.Source = resource.SetImage(cornerFileName, cornerName);
            }
            else
            {
                CornerImage.Source = null;
            }

            //显示活动名
            Binding taskNamebinding = new Binding
            {
                Source = ExcelDataTable.Rows[row],
                Path = new PropertyPath("[任务名称]"),
                Mode = BindingMode.TwoWay
            };
            TaskName.SetBinding(TextBlock.TextProperty, taskNamebinding);

            //显示活动时间
            Binding timebinding = new Binding
            {
                Source = ExcelDataTable.Rows[row],
                Path = new PropertyPath("[任务时间]"),
                Mode = BindingMode.TwoWay
            };
            Time.SetBinding(TextBlock.TextProperty, timebinding);

            //显示按钮
            Button.Source = Resource.ButtonSource;
        }
    }
}