
using System;
using System.Text;
using System.IO;
using NPOI.HSSF.UserModel;
using NPOI.HPSF;
using NPOI.POIFS.FileSystem;
using NPOI.SS.UserModel;


namespace WindowsFormsApplication1
{
    class ReportFile
    {
        string currentPath = "D:\\";
        string fileName = "产品维修记录汇总表.xls";
        public HSSFWorkbook hssfworkbook;
        public ReportFile(string path)
        {
            currentPath = Path.GetDirectoryName(path);
            fileName = Path.GetFileName(path);
            InitializeWorkbook(path); 
        }
        public ReportFile(Stream stream)
        {
            InitializeWorkbook(stream);
        }
        public string updatefileName(string name)
        {
           return updatefileName("", name);
        }
        public string updatefileName(string city, string name)
        {
            fileName = string.Format("{0}{1}月 产品维修记录汇总表.xls", city+name, DateTime.Now.Month);
            return fileName;
        }
        public void updatefilePath(string path)
        {
            currentPath = path;
        }

        public string WriteToFile()
        {
            //Write the stream data of workbook to the root directory
            FileStream file = new FileStream(currentPath+fileName, FileMode.Create);
            hssfworkbook.Write(file);
            file.Close();
            return fileName;
        }

        public void setTitle(DateTime date)
        {
            string title = string.Format("  南宁  分公司(办事处)  {0}  年  {1}  月   ",date.Year,date.Month);
            HSSFRichTextString Richtitle = new HSSFRichTextString(title);
            IFont font = hssfworkbook.CreateFont();
            font.Underline = FontUnderlineType.Single;
            font.FontHeightInPoints = 16;
            font.FontName = "黑体";
            font.IsBold = true;
            IFont font1 = hssfworkbook.CreateFont();
            font1.FontHeightInPoints = 16;
            font1.FontName = "黑体";
            font1.IsBold = true;
            Richtitle.ApplyFont(0,6,font);
            Richtitle.ApplyFont(6, 15,font1);
            Richtitle.ApplyFont(15, 22, font);
            Richtitle.ApplyFont(22,24, font1);
            Richtitle.ApplyFont(24, 28, font);
            Richtitle.ApplyFont(28, 30, font1);
            ICell cell = hssfworkbook.GetSheet("Sheet1").GetRow(1).GetCell(1);
            cell.SetCellValue(Richtitle);
            
        }
        public void InitializeWorkbook()
        {
            InitializeWorkbook("D:\\abcc.xls");
        }
        public void InitializeWorkbook(string path)
        {
            //read the template via FileStream, it is suggested to use FileAccess.Read to prevent file lock.
            //book1.xls is an Excel-2007-generated file, so some new unknown BIFF records are added. 
            //FileStream file = new FileStream(@"template/book1.xls", FileMode.Open,FileAccess.Read);
            FileStream file = new FileStream(path, FileMode.Open, FileAccess.Read);
            InitializeWorkbook(file);
        }
        public void InitializeWorkbook(Stream stream)
        {
            hssfworkbook = new HSSFWorkbook(stream);

            //create a entry of DocumentSummaryInformation
            DocumentSummaryInformation dsi = PropertySetFactory.CreateDocumentSummaryInformation();
            dsi.Company = "GuoGuang Co.Ltd";
            hssfworkbook.DocumentSummaryInformation = dsi;

            //create a entry of SummaryInformation
            SummaryInformation si = PropertySetFactory.CreateSummaryInformation();
            si.Subject = "hoisung's project";
            hssfworkbook.SummaryInformation = si;
        }
    }
}
