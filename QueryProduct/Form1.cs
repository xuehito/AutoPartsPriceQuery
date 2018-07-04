using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Microsoft.Win32;
using Application = Microsoft.Office.Interop.Excel.Application;
using DataTable = System.Data.DataTable;

namespace QueryProduct
{
    public partial class Form1 : Form
    {
        private readonly Stopwatch sw = new Stopwatch();

        public Form1()
        {
            InitializeComponent();

            Text = "配件价格查询软件v2.3";

            //ControlBox = false;隐藏窗体上自带控件

            label6.Text = "Copyright © 2014 - xuehito";
            label7.Text = "联系邮箱：xhito@foxmail.com";

            loadingProgress1.Stop();
        }

        public override sealed string Text
        {
            get { return base.Text; }
            set { base.Text = value; }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (label1.Text == string.Empty)
            {
                MessageBox.Show("请选择要采集的文件！", "提示");
                return;
            }
            if (textBox3.Text.Trim().Length <= 0)
            {
                MessageBox.Show("请输入要保存的文件名称！", "提示");
            }
            else
            {
                if (!File.Exists(label1.Text))
                {
                    MessageBox.Show("文件不存在");
                }
                else
                {
                    loadingProgress1.Start();
                    textBox1.AppendText("开始查询数据..." + Environment.NewLine);

                    button1.Enabled = false;
                    richTextBox2.Clear();

                    sw.Start();

                    textBox1.AppendText("开始计时..." + Environment.NewLine);

                    ReadExcel(label1.Text);
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            var openFileDialog = new OpenFileDialog();
            openFileDialog.InitialDirectory = "c:\\";
            openFileDialog.Filter = "文本文件|*.*|C#文件|*.cs|所有文件|*.*";
            openFileDialog.RestoreDirectory = true;
            openFileDialog.FilterIndex = 1;
            if (openFileDialog.ShowDialog() != DialogResult.OK) return;
            string fName = openFileDialog.FileName;
            label1.Text = fName;
        }

        /// <summary>
        ///     读取Excel
        /// </summary>
        public void ReadExcel(string filePath)
        {
            string num = string.Empty;

            try
            {
                ArrayList arrList = ExcelHelper.GetSheetNames(filePath);

                int i = 0;

                Task.Factory.StartNew(() => Product.QueryProduct(arrList, filePath, delegate(ProductInfo info)
                {
                    int i1 = i;

                    textBox1.Invoke(
                        () =>
                            textBox1.AppendText("序号：" + i1 + Environment.NewLine));

                    richTextBox2.Invoke(
                        () => richTextBox2.AppendText(info.Num + "  " + info.DPrice + Environment.NewLine));

                    i++;
                })).ContinueWith(task =>
                {
                    if (task.IsCompleted)
                    {
                        textBox1.Invoke(() =>
                        {
                            textBox1.AppendText("查询数据完成..." + Environment.NewLine);

                            MessageBox.Show("数据全部查询完成！总共：" + i + "条数据。" + Environment.NewLine, "提示");

                            string swtime = sw.Elapsed.Hours + ":" + sw.Elapsed.Minutes + ":" + sw.Elapsed.Seconds + ":" +
                                            sw.Elapsed.Milliseconds;

                            textBox1.AppendText("总共用时:" + swtime + Environment.NewLine);

                            loadingProgress1.Stop();
                        });

                        textBox1.Invoke(() => textBox1.AppendText("开始导出数据" + Environment.NewLine));
                        MessageBox.Show("开始导出数据" + Environment.NewLine, "提示");

                        List<ProductInfo> listInfo = Product.QueryList();
                        DataTable dts = new Class1()._ToDataTable(listInfo);
                        ExportExcel(dts);

                        textBox1.Invoke(() => button1.Enabled = true);
                    }
                });
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "-----配件编号" + num, "出错啦~~");
                loadingProgress1.Stop();
            }
        }

        /// <summary>
        ///     导数excel
        /// </summary>
        /// <param name="dt"></param>
        protected void ExportExcel(DataTable dt)
        {
            if (dt == null || dt.Rows.Count == 0) return;

            var xlApp = new Application();

            if (xlApp == null) return;
            CultureInfo CurrentCI = Thread.CurrentThread.CurrentCulture;
            Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");
            Workbooks workbooks = xlApp.Workbooks;
            Workbook workbook = workbooks.Add(XlWBATemplate.xlWBATWorksheet);
            var worksheet = (Worksheet) workbook.Worksheets[1];
            Range range;
            long totalCount = dt.Rows.Count;
            long rowRead = 0;
            float percent = 0;
            for (int i = 0; i < dt.Columns.Count; i++)
            {
                worksheet.Cells[1, i + 1] = dt.Columns[i].ColumnName;
                range = (Range) worksheet.Cells[1, i + 1];
                range.Interior.ColorIndex = 15;
                range.Font.Bold = true;
            }
            for (int r = 0; r < dt.Rows.Count; r++)
            {
                for (int i = 0; i < dt.Columns.Count; i++) worksheet.Cells[r + 2, i + 1] = dt.Rows[r][i].ToString();
                rowRead++;
                percent = ((float) (100*rowRead))/totalCount;
            }
            // xlApp.Visible = true;
            //string path = Path.Combine(Environment.CurrentDirectory, @"\Resources\"  + "xxx.xlsx");
            string date = DateTime.Now.Year + "-" + DateTime.Now.Month + "-" + DateTime.Now.Day;
            //RegistryKey folders = OpenRegistryPath(Registry.CurrentUser,
            //    @"\software\microsoft\windows\currentversion\explorer\shell folders");
            //string desktopPath = folders.GetValue("Desktop").ToString(); //桌面路径

            string path = @"D:\" + textBox3.Text + "-" + date + ".xlsx";

            workbook.SaveCopyAs(path);
            textBox1.Invoke(() => textBox1.AppendText("导出完成！" + Environment.NewLine));
            textBox1.Invoke(() => textBox1.AppendText("默认文件路径在D盘根目录下..." + Environment.NewLine));
            MessageBox.Show("导出完成！", "提示");
        }

        private RegistryKey OpenRegistryPath(RegistryKey root, string s)
        {
            s = s.Remove(0, 1) + @"\";
            while (s.IndexOf(@"\") != -1)
            {
                root = root.OpenSubKey(s.Substring(0, s.IndexOf(@"\")));
                s = s.Remove(0, s.IndexOf(@"\") + 1);
            }
            return root;
        }
    }
}