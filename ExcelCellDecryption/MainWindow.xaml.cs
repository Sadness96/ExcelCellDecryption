using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
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
        /// 删除标识,包含则删除
        /// </summary>
        private List<string> listRemoveIdentification = new List<string>()
        {
            "html:Color=\"#FFFFF2\"",
            "html:Color=\"#FFFFF1\"",
            "html:Color=\"#FFFFCC\"",
            "html:Color=\"#FFFFFF\"",
            "html:Size=\"1\"",
            "html:Size=\"2\""
        };

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
            dialog.Filter = "Excel 电子表格 2003(*.xml)|*.xml";
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
            if (File.Exists(TargetPath.Text))
            {
                File.Delete(TargetPath.Text);
            }
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
                // 解析 Excel XML 文档
                XmlDocument doc = new XmlDocument();
                doc.Load(vNewFile);
                XmlNamespaceManager nsmgr = new XmlNamespaceManager(doc.NameTable);
                nsmgr.AddNamespace("ab", "http://www.w3.org/TR/REC-html40");
                nsmgr.AddNamespace("ss", "urn:schemas-microsoft-com:office:spreadsheet");
                // 删除掺杂的数据
                XmlNodeList nodeFonts = doc.SelectNodes("//ab:Font", nsmgr);
                for (int i = 0; i < nodeFonts.Count; i++)
                {
                    var vXmlNodeFont = nodeFonts[i];
                    bool bIsRemove = false;
                    foreach (var item in listRemoveIdentification)
                    {
                        if (vXmlNodeFont.OuterXml.Contains(item))
                        {
                            bIsRemove = true;
                            break;
                        }
                    }
                    if (bIsRemove)
                    {
                        var vParentNode = vXmlNodeFont.ParentNode;
                        vParentNode.RemoveChild(vXmlNodeFont);
                    }
                }
                // 合并整理后的数据
                XmlNodeList nodeDatas = doc.SelectNodes("//ss:Data", nsmgr);
                for (int i = 0; i < nodeDatas.Count; i++)
                {
                    var vXmlNodeData = nodeDatas[i];
                    var vXmlNodeFonts = vXmlNodeData.ChildNodes;
                    if (vXmlNodeFonts.Count >= 2)
                    {
                        // Data 中 Font 数量大于等于 2 需要合并
                        string strTxt = "";
                        XmlNode xmlNodeMain = null;
                        List<XmlNode> xmlNodesPrepare = new List<XmlNode>();
                        // 记录数据 拼接文本 记录主要 Font 和需要删除的 Font
                        for (int j = 0; j < vXmlNodeFonts.Count; j++)
                        {
                            var vXmlNodeFont = vXmlNodeFonts[j];
                            if (j == 0)
                            {
                                xmlNodeMain = vXmlNodeFont;
                            }
                            else
                            {
                                xmlNodesPrepare.Add(vXmlNodeFont);
                            }
                            strTxt += vXmlNodeFont.InnerText;
                        }
                        // 记录主要 Font,超过15位增加 "'"
                        if (strTxt.Length >= 15 && IsNumeric(strTxt) && !strTxt.First().Equals('\''))
                        {
                            xmlNodeMain.InnerText = $"'{strTxt}";
                        }
                        else
                        {
                            xmlNodeMain.InnerText = strTxt;
                        }
                        // 删除的 Font
                        var vParentNode = xmlNodeMain.ParentNode;
                        for (int k = 0; k < xmlNodesPrepare.Count; k++)
                        {
                            vParentNode.RemoveChild(xmlNodesPrepare[k]);
                        }
                    }
                }
                doc.Save(vNewFile);
                // 保存最终解密文件
                File.Copy(vNewFile, TargetPath.Text);
                // 删除临时文件夹
                DirectoryInfo dir = new DirectoryInfo(vTempPath);
                dir.Delete(true);
                MessageBox.Show("执行完成！");
            }
            else
            {
                MessageBox.Show("请选择单元格加密的Excel文件！");
            }
        }

        public static bool IsNumeric(string value)
        {
            return Regex.IsMatch(value, @"^[+-]?\d*[.]?\d*$");
        }
    }
}
