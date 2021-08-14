using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace ExcelCellDecryption
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        /// <summary>
        /// 选择源文件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void SelectSource_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Multiselect = false;//该值确定是否可以选择多个文件
            dialog.Title = "请选择单元格异常文件";
            dialog.Filter = "Excel 文件(*.xlsx)|*.xlsx";
            if (dialog.ShowDialog() == true)
            {
                SourcePath.Text = dialog.FileName;
                TargetPath.Text = $"{System.IO.Path.GetDirectoryName(dialog.FileName)}\\{System.IO.Path.GetFileNameWithoutExtension(dialog.FileName)}_decrypt{System.IO.Path.GetExtension(dialog.FileName)}";
            }
        }

        /// <summary>
        /// 执行
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Implement_Click(object sender, RoutedEventArgs e)
        {
            if (!string.IsNullOrEmpty(SourcePath.Text) && !string.IsNullOrEmpty(TargetPath.Text))
            {

            }
            else
            {
                MessageBox.Show("请选择单元格加密的Excel文件！");
            }
        }
    }
}
