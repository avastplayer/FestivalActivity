using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Xml;
using NPOI.SS.Formula.Functions;
using Color = System.Windows.Media.Color;
using ColorConverter = System.Windows.Media.ColorConverter;
using Image = System.Windows.Controls.Image;

namespace FestivalActivity
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow
    {
        public Cell[] ActivityCell { get; private set; }
        private readonly Resource _resource = new Resource(false);
        private List<int> _selectRow = new List<int>();

        public MainWindow()
        {
            InitializeComponent();

            InitializeFrameWindow();
        }

        private void InitializeFrameWindow()
        {
            ExcelData.ItemsSource = _resource.ExcelDataItemsSource;

            Image[] mainBack = { MainBack1, MainBack2, MainBack3, MainBack4, MainBack5, MainBack6, MainBack7, MainBack8, MainBack9 };
            Image[] commonCase = { CommonCase1, CommonCase2, CommonCase3, CommonCase4, CommonCase5, CommonCase6, CommonCase7, CommonCase8, CommonCase9 };
            for (int i = 0; i < 9; i++)
            {
                mainBack[i].Source = _resource.MainBackSource[i];
                commonCase[i].Source = _resource.CommonCaseSource[i];
            }
            CloseButton.Source = _resource.CloseButtonSource;
            TopLeftCorner.Source = _resource.TopLeftCornerSource;
            TopRightCorner.Source = _resource.TopRightCornerSource;
            BottomLeftCorner.Source = _resource.BottomLeftCornerSource;
            BottomRightCorner.Source = _resource.BottomRightCornerSource;
            Pattern.Source = _resource.PatternSource;

            InitializeInfoCell();
        }

        public void InitializeInfoCell()
        {
            Image[] commonCaseSell = { CommonCase_1, CommonCase_2, CommonCase_3, CommonCase_4, CommonCase_5, CommonCase_6, CommonCase_7, CommonCase_8, CommonCase_9 };
            Image[] commonCase7 = { CommonCase7_1, CommonCase7_2, CommonCase7_3, CommonCase7_4, CommonCase7_5, CommonCase7_6, CommonCase7_7, CommonCase7_8, CommonCase7_9 };
            for (int i = 0; i < 9; i++)
            {
                commonCaseSell[i].Source = _resource.CommonCaseSellSource[i];
                commonCase7[i].Source = _resource.CommonCase7Source[i];
            }
            LiMoChouPic.Source = _resource.LiMoChouSource;
            CloseButton1.Source = _resource.CloseButtonSource;
        }

        private DataRowView[] GetSelectedRows()
        {
            //当选中有多个单元格时，获取选中单元格所在行的数组
            //排除数组中相同的行
            if (ExcelData == null || ExcelData.SelectedCells.Count <= 0) return null;
            DataRowView[] dv = new DataRowView[ExcelData.SelectedCells.Count];
            for (int i = 0; i < ExcelData.SelectedCells.Count; i++)
            {
                dv[i] = ExcelData.SelectedCells[i].Item as DataRowView;
            }
            //因为选中的单元格可能在同一行的，需要排除重复的行
            return dv.Distinct().ToArray();
        }

        // <summary>
        // 更改数据
        // </summary>
        private void ExcelData_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            ScrollPane.Children.Clear();
            //获取选中行的数组
            DataRowView[] selectRowView = GetSelectedRows();
            if (selectRowView == null) return;

            _selectRow = (from t in selectRowView where t.Row.ItemArray[0].ToString() != "" select Convert.ToInt32(t.Row.ItemArray[0].ToString()) - 1).ToList();
            int cellNumber = _selectRow.Count;
            ActivityCell = new Cell[cellNumber];
            //显示超过6个需要设置滑动
            if (cellNumber > 8)
            {
                ScrollPane.Height = cellNumber / 2 * 101;
            }
            else
            {
                ScrollPane.Height = 484;
            }

            for (int i = 0; i < cellNumber; i++)
            {
                ActivityCell[i] = new Cell(_resource.ExcelDataTable, _selectRow[i]);
                Canvas.SetTop(ActivityCell[i], i / 2 * 100 + 1);
                Canvas.SetLeft(ActivityCell[i], i % 2 * 400 + 9);
                ScrollPane.Children.Add(ActivityCell[i]);
                //绑定点击事件
                ActivityCell[i].MouseUp += ActivityCell_MouseUp;
                //设置第一个选中的cell说明
                SetInfoMain(_selectRow[0]);
            }
        }

        private void ActivityCell_MouseUp(object sender, MouseButtonEventArgs e)
        {
            for (int i = 0; i < ActivityCell.Length; i++)
            {
                if (!ActivityCell[i].Equals(sender)) continue;
                SetInfoMain(_selectRow[i]);
            }

            InfoCell.Visibility = Visibility.Visible;
        }

        private void SetInfoMain(int row)
        {
            InfoMain.Inlines.Clear();

            string xmlString = _resource.ExcelDataTable.Rows[row]["说明"].ToString();

            XmlDocument xmldoc = new XmlDocument();
            xmldoc.LoadXml("<root>" + xmlString + "</root>");

            if (xmldoc.DocumentElement == null) return;
            XmlNodeList xmlNodeList = xmldoc.DocumentElement.ChildNodes;

            foreach (XmlNode xmlNode in xmlNodeList)
            {
                switch (xmlNode.Name)
                {
                    case "T":
                        if (xmlNode.Attributes["c"] != null)
                        {
                            string colorString = Regex.Replace(xmlNode.Attributes["c"].Value, "^ff|^FF", "#");
                            Color color = (Color)ColorConverter.ConvertFromString(colorString);
                            InfoMain.Inlines.Add(new Run(xmlNode.Attributes["t"].Value) { Foreground = new SolidColorBrush(color) });
                        }
                        else
                        {
                            InfoMain.Inlines.Add(new Run(xmlNode.Attributes["t"].Value));
                        }
                        break;

                    case "B":
                        InfoMain.Inlines.Add(new LineBreak());
                        break;
                }
            }
        }

        private void CloseButton1_OnMouseUp(object sender, MouseButtonEventArgs e)
        {
            InfoCell.Visibility = Visibility.Hidden;
        }
    }
}