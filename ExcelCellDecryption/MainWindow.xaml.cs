using ExcelCellDecryption.Helper;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
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
using System.Xml;

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
        /// 解密静态变量
        /// </summary>
        private const string DECRYPT = "decrypt";

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
                TargetPath.Text = $"{System.IO.Path.GetDirectoryName(dialog.FileName)}\\{System.IO.Path.GetFileNameWithoutExtension(dialog.FileName)}_{DECRYPT}{System.IO.Path.GetExtension(dialog.FileName)}";
            }
        }

        /// <summary>
        /// 执行
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Implement_Click(object sender, RoutedEventArgs e)
        {
            if (!string.IsNullOrEmpty(SourcePath.Text) && !string.IsNullOrEmpty(TargetPath.Text) && File.Exists(SourcePath.Text))
            {
                var vTempPath = $"{System.IO.Path.GetDirectoryName(SourcePath.Text)}\\DecryptTemp_{DateTime.Now.Ticks}";
                var vNewFile = $"{vTempPath}\\{System.IO.Path.GetFileName(SourcePath.Text)}";
                // 创建临时文件夹
                if (!Directory.Exists(vTempPath))
                {
                    Directory.CreateDirectory(vTempPath);
                }
                // 拷贝需要处理的文件
                File.Copy(SourcePath.Text, vNewFile);
                // 修改后缀名称
                var vChangedName = System.IO.Path.ChangeExtension(vNewFile, "zip");
                File.Copy(vNewFile, vChangedName);
                // 解压文件
                var vZipPath = $"{System.IO.Path.GetDirectoryName(vChangedName)}\\{System.IO.Path.GetFileNameWithoutExtension(vChangedName)}";
                var vXMLPath = $"{vZipPath}\\xl\\sharedStrings.xml";
                var vIsDeCompressionZip = ZIPHelper.DeCompressionZip(vChangedName, vZipPath);
                if (vIsDeCompressionZip && File.Exists(vXMLPath))
                {
                    // 解析 Excel XML 文档
                    XmlDocument Document = new XmlDocument();
                    Document.Load(vXMLPath);
                    // TODO:处理数据

                    // 保存回 Excel Zip
                    var vSaveExcelZip = $"{vZipPath}_{DECRYPT}.zip";
                    List<string> listFolder = new List<string>();
                    listFolder.AddRange(FolderHelper.GetSpecifiedDirectoryFolders(vZipPath));
                    listFolder.AddRange(FolderHelper.GetSpecifiedDirectoryFiles(vZipPath));
                    var vIsCompressionZip = ZIPHelper.CompressionZip(vSaveExcelZip, listFolder);
                    if (vIsCompressionZip)
                    {
                        // 保存回 Excel 文件
                        File.Copy(vSaveExcelZip, TargetPath.Text);
                    }
                    else
                    {
                        MessageBox.Show("执行失败，文件压缩失败！");
                    }
                }
                else
                {
                    MessageBox.Show("执行失败，无法解压Excel文件！");
                }
                // 删除临时文件夹
                DirectoryInfo dir = new DirectoryInfo(vTempPath);
                dir.Delete(true);
            }
            else
            {
                MessageBox.Show("请选择单元格加密的Excel文件！");
            }
        }
    }
}
