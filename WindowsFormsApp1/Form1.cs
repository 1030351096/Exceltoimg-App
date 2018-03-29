using Aspose.Cells;
using Aspose.Cells.Rendering;
using System;
using System.Drawing;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            LicenseHelper.ModifyInMemory.ActivateMemoryPatching();
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            openFileDialog1.Multiselect = false;
            openFileDialog1.CheckPathExists = true;
            openFileDialog1.ShowDialog();
            string filename = openFileDialog1.SafeFileName;
            if (filename.Substring(filename.LastIndexOf("."), 4) == ".xls")
            {
                label2.ForeColor = Color.Blue;
                label2.Text = $"开始转换文件路径:{openFileDialog1.FileName}";
                ExcelToImg(openFileDialog1.FileName);
            }
            else
            {
                MessageBox.Show("请选择Excel文件,xls,xlsx");
            }
        }

        public async void ExcelToImg(string ExcelPath)
        {
            Workbook book = new Workbook(ExcelPath);
            var list = book.Worksheets;
            int count = 0;
            foreach (var item in list)
            {

                count ++;
                item.PageSetup.LeftMargin = 0;
                item.PageSetup.RightMargin = 0;
                item.PageSetup.BottomMargin = 0;
                item.PageSetup.TopMargin = 0;
                ImageOrPrintOptions imgOptions = new ImageOrPrintOptions();
                imgOptions.ImageFormat = System.Drawing.Imaging.ImageFormat.Jpeg;
                imgOptions.OnePagePerSheet = true;
                imgOptions.PrintingPage = PrintingPageType.IgnoreBlank;
                SheetRender sr = new SheetRender(item, imgOptions);
                string filepath = $@"{Path.GetDirectoryName(ExcelPath)}\{item.Name}.jpg";
                await Task.Run(() => sr.ToImage(0, filepath));
            }

            label2.ForeColor = Color.Green;
            label2.Text = $"转换完成:图片已存放到{Path.GetDirectoryName(ExcelPath)}";
        }


    }
}
