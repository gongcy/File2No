using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using NPOI;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;

namespace File2No
{
    public partial class Form1 : Form
    {
        IWorkbook wk = null;
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            this.ContextMenuStrip = contextMenuStrip1;
        }

        // btnSelect
        private void button1_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog dialog = new FolderBrowserDialog();
            dialog.Description = "请选择文件路径";
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                string foldPath = dialog.SelectedPath;
                textPath.Text = foldPath;
            }
        }

        // btnExport
        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                if (textPath.Text.Length == 0)
                {
                    throw new Exception();
                }
                DirectoryInfo theFolder = new DirectoryInfo(textPath.Text);
                if (!theFolder.Exists)
                {
                    throw new Exception();
                }
                FileInfo[] fileinfo = theFolder.GetFiles();
   
                string extension = System.IO.Path.GetExtension(textInp.Text);
                FileStream fs = File.OpenRead(textInp.Text);
                if (extension.Equals(".xls"))
                {
                    // for .xls
                    wk = new HSSFWorkbook(fs);
                }
                else
                {
                    // for .xlsx
                    wk = new XSSFWorkbook(fs);
                }
                fs.Close();
                ISheet sheet = wk.GetSheetAt(0);

                for (int i = 0; i < fileinfo.Length; i++)
                {
                    Regex reg = new Regex("201\\d+");
                    if (reg.IsMatch(fileinfo[i].Name))
                    {
                        string No = reg.Match(fileinfo[i].Name).ToString();
                        searchAndWrite(sheet, No);
                    }
                }
                sheet.AutoSizeColumn(0);
                fs = File.OpenWrite(textInp.Text);
                wk.Write(fs);
                fs.Flush();
                fs.Close();
                MessageBox.Show(null, "导出成功！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                
            }
            catch (Exception ex)
            {
                MessageBox.Show(null, "导出失败！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Console.Out.WriteLine(ex.Message);
            }
        }

        private void searchAndWrite(ISheet sheet, string No)
        {
            ICellStyle style = wk.CreateCellStyle();
            style.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
            style.ShrinkToFit = true;

            for (int i = 0; i <= sheet.LastRowNum; i++)
            {
                IRow row = sheet.GetRow(i);
                if (row != null)
                {
                    for (int j = 0; j <= row.LastCellNum; j++)
                    {
                        ICell cell = row.GetCell(j);
                        if (cell != null)
                        {
                            string value = cell.ToString();
                            if (value.Equals(No))
                            {
                                cell = row.CreateCell(row.LastCellNum);
                                cell.SetCellValue("Y");
                                cell.CellStyle = style;
                                // cell.SetCellType(CellType.String);
                            }
                        }
                    }
                }
            }
        }

        // btnInp
        private void button1_Click_1(object sender, EventArgs e)
        {
            OpenFileDialog od = new OpenFileDialog();
            od.Title = "请选择花名册文件";
            od.Filter = "EXCEL文件(*.xls)|*.xls|EXCEL文件(*.xlsx)|*.xlsx";
            if (od.ShowDialog() == DialogResult.OK)
            {
                FileInfo theFile = new FileInfo(od.FileName);
                string filePath =theFile.FullName;
                textInp.Text = filePath;
            }

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void textPath_TextChanged(object sender, EventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void 帮助ToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void 关于ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show(null, "点名程序\n如有问题，请联系gongcy126@126.com", "关于", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }
}
