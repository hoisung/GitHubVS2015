using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;

namespace WindowsFormsApplication1
{
    public enum orderingType:int { MaintenanceDate=1, Model};

    class ATMDataStruct
    {
        /// <summary>
        /// 产品类型
        /// </summary>
        public int Model = 1;//产品类型
        /// <summary>
        /// 机器编号
        /// </summary>
        public int ProductID = 0;//机器编号
        /// <summary>
        /// 故障现象
        /// </summary>
        public int ErrorSrc = 16;//故障现象
        /// <summary>
        /// 处理方法
        /// </summary>
        public int DealSrc = 17;//处理方法
        /// <summary>
        /// 用户单位
        /// </summary>
        public int ClientName = 29; //用户单位
        /// <summary>
        /// 工程师姓名
        /// </summary>
        public int Engineer = 20;//工程师姓名
        /// <summary>
        /// 工作日期
        /// </summary>
        public int MaintenanceDate = 11;//工作日期
        /// <summary>
        /// 工单编号
        /// </summary>
        public int PaiID =3;//工单编号
        /// <summary>
        /// 维修单编号
        /// </summary>
        public int ServiceID = 4;//维修单编号

        /// <summary>
        /// 工单类别  
        /// **非所需导出列数据，为程序筛选数据用**
        /// </summary>
        public int WorkStyle = 5;//工单类别  **非所需导出列数据，为程序筛选数据用**
        /// <summary>
        /// 所属银行      
        /// **非所需导出列数据**
        /// </summary>
        public int bank = 25;//所属银行       **非所需导出列数据**
        /// <summary>
        /// 服务方式 
        /// **非所需导出列数据，判断是否为电话解决类型，电话解决则不导出**
        /// </summary>
        public int serviceStyle = 5;//服务方式 **非所需导出列数据，判断是否为电话解决类型，电话解决则不导出**
    }
    /// <summary>
    /// 自助工单各数据所在列结构
    /// </summary>
    class GGWOWMDataStruct
    {
        /// <summary>
        /// 产品类型
        /// </summary>
        public int Model = 12;
        /// <summary>
        /// 机器编号
        /// </summary>
        public int ProductID = 11;
        /// <summary>
        /// 故障现象
        /// </summary>
        public int ErrorSrc = 18;
        /// <summary>
        /// 处理方法
        /// </summary>
        public int DealSrc = 19;
        /// <summary>
        /// 用户单位
        /// </summary>
        public int ClientName = 8;
        ///<summary>
        ///工程师姓名
        /// </summary> 
        public int Engineer = 31;
        /// <summary>
        /// 工作日期
        /// </summary>
        public int MaintenanceDate = 24;
        /// <summary>
        /// 工单编号
        /// </summary>
        public int PaiID = 1;
        /// <summary>
        /// 维修单编号
        /// </summary>
        public int ServiceID = 16;
        /// <summary>
        /// 工单类别  
        /// **非所需导出列数据，为程序筛选数据用**
        /// </summary>
        public int WorkStyle = 3;
        /// <summary>
        /// 所属银行       
        /// **非所需导出列数据**
        /// </summary>
        public int bank = 6;
        /// <summary>
        /// 服务方式 
        /// **非所需导出列数据，判断是否为电话解决类型，电话解决则不导出**
        /// </summary>
        public int serviceStyle = 17;
    }
    class XlsContentProcess
    {
        //**************field of exchange data***********************
        //*****xlsx from online to the report file(xls)***********
        XSSFWorkbook sourceWB;//源XLSX数据；
        ReportFileColumns RFcolumns = new ReportFileColumns();
        //*******************END*************************************

        ATMDataStruct AStruct = new ATMDataStruct();
        GGWOWMDataStruct GGStruct = new GGWOWMDataStruct();
        string savePath = "d:\\";
        string Station;
        List<ReportFileColumns> list = new List<ReportFileColumns>();
        ReportFile reportfile;

        public XlsContentProcess(string path)
        {

            savePath = path==null?savePath:path;
            byte[] filebuffer = Properties.Resources.baseXLSRes;
            Stream filestream = new MemoryStream(filebuffer);
            reportfile = new ReportFile(filestream);
            // ISheet sheet = sourceWB.GetSheet("Sheet1");
        }
        private bool isHaveOtherDataFile()
        {
            return File.Exists("维修档案查询统计表.xls");
        }
        private void processATMFXlsData(byte[] xlsdatas)
        {
           // FileStream fs = new FileStream("temp.xls", FileMode.Create);
             //fs.Write(xlsdatas, 0, xlsdatas.Length);
           //  fs.Close();
            Stream sm = new MemoryStream(xlsdatas);
            sourceWB = new XSSFWorkbook(sm);
            int lenght = sourceWB.GetSheet("Sheet1").LastRowNum;
            if (lenght < 1) return;
            for (int i = 1; i <= lenght; i++)
            {
                if (sourceWB.GetSheet("Sheet1").GetRow(i).GetCell(AStruct.serviceStyle).StringCellValue != "电话解决")
                    list.Add(AppendATMDataToColumnsFormat(i));
            }
        }
        private void processSetupXlsData(byte[] xlsdatas)
        {
            Stream sm = new MemoryStream(xlsdatas);
            sourceWB = new XSSFWorkbook(sm);
            int lenght = sourceWB.GetSheet("Sheet1").LastRowNum;
            if (lenght < 1) return;
            for (int i = 1; i <= lenght; i++)
            {
                    list.Add(AppendSetupDataToColumnsFormat(i));
            }
        }
        private void processGGWOWMXlsData(byte[] xlsdatas)
        {
            //FileStream fs = new FileStream("temp.xls", FileMode.Create);
           // fs.Write(xlsdatas, 0, xlsdatas.Length);
           // fs.Close();
                Stream sm = new MemoryStream(xlsdatas);
                HSSFWorkbook otherFileData = new HSSFWorkbook(sm);
                int lenght = otherFileData.GetSheetAt(0).LastRowNum;
            if (lenght < 1) return;
                ISheet sheet1 = otherFileData.GetSheetAt(0);
                for (int i = 1; i <= lenght; i++)
                {

                    if (sheet1.GetRow(i).GetCell(GGStruct.serviceStyle).StringCellValue != "电话解决")
                        list.Add(AppendOtherDataToColumnsFormat(sheet1, i));


                }
                //file.Close();
                //File.Delete("维修档案查询统计表.xls");
            
        }
        private void UpdateDataStruct(string ConfigPath)
        {
            INIOperater INIOper = new INIOperater(ConfigPath);
            if (INIOper.isFileExists())
            {
                FieldInfo[] finfoes = AStruct.GetType().GetFields();
                foreach (FieldInfo finfo in finfoes)
                {
                    string tem =INIOper.GetKey("ATMDataPosition", finfo.Name);
                    if (tem != "")
                    {
                        try
                        {
                            int tempNum = int.Parse(tem.Trim());
                            finfo.SetValue(AStruct, tempNum);
                        }
                        catch (Exception)
                        {
                            continue;
                        }
                    }
                    else INIOper.WriteKey("ATMDataPosition", finfo.Name, finfo.GetValue(AStruct).ToString());                       
                }

                FieldInfo[] finfies1 = GGStruct.GetType().GetFields();
                foreach (FieldInfo finfo in finfies1)
                {
                    string tem = INIOper.GetKey("GGWOWMDataPosition", finfo.Name);
                    if (tem != "")
                    {
                        try
                        {
                            int tempNum = int.Parse(tem.Trim());
                            finfo.SetValue(GGStruct, tempNum);
                        }
                        catch (Exception)
                        {
                            continue;
                        }
                    }
                    else INIOper.WriteKey("GGWOWMDataPosition", finfo.Name, finfo.GetValue(GGStruct).ToString());
                }
            }
         }
        public string process(byte[] ATMFDatas, byte[] GGWOWMData, string station, byte[] SETUPData,string ConfigPath)
        {
            Station = station;
            UpdateDataStruct(ConfigPath);
            processATMFXlsData(ATMFDatas);
            processGGWOWMXlsData(GGWOWMData);
            processSetupXlsData(SETUPData);
            orderingList(orderingType.Model);
            return Exportfile();
        }
        private void orderingList(orderingType type)
        {
            // ReportFile reportfile = new ReportFile("D:\\abcc.xls");
            int Rowindex = 3;
            if (list.Count >= 1) { 
            
            if (type == orderingType.MaintenanceDate)
                list = (from c in list orderby c.MaintenanceDate ascending select c).ToList();
            else
                list = (from c in list orderby c.Model ascending select c).ToList();
             
            foreach (ReportFileColumns RFC in list)
            {
                AppendToXlscell(reportfile, Rowindex, RFC);
                Rowindex++;
            }
            }
            reportfile.setTitle(DateTime.Now);
            
           // textBox1.Text = "月报已生成.\r\n****************\r\n";
           // textBox1.Text += string.Format("路  径: {0};\r\n\r\n文件名： {1};", savePath, currentNa);
        }
        private string Exportfile()
        {
            if (list.Count == 0) return "";
            string EName = list[0].Engineer;
            string currentNa = reportfile.updatefileName(Station,EName);
            reportfile.updatefilePath(savePath);
            return reportfile.WriteToFile();
        }

        //{ textBox1.Text = "工号为空，请输入工号!"; }


  /*      public void AppendToColumnsFormatForWebpage()
        {
            RFcolumns.PaiID = DFList.PaiID;
            RFcolumns.ServiceID = DFList.ServiceID;
            RFcolumns.ProductID = DFList.MID;
            RFcolumns.MaintenanceDate = DFList.AvTime;
            RFcolumns.ErrorSrc = DFList.ErrorScr;
            RFcolumns.DealSrc = DFList.DealScr;
            RFcolumns.Engineer = DFList.staffnum;
        }*/
        public ReportFileColumns AppendOtherDataToColumnsFormat(ISheet sheet, int rowindex)
        {
            IRow Row = sheet.GetRow(rowindex);
            ReportFileColumns temRFcolumns = new ReportFileColumns();
            temRFcolumns.Model = getCellValue(Row, GGStruct.Model);//Row.GetCell(1).StringCellValue;
                                                       //temRFcolumns.Model = "HT-" + temRFcolumns.Model;
            temRFcolumns.ProductID = getCellValue(Row, GGStruct.ProductID).Replace(temRFcolumns.Model + " ", ""); //Row.GetCell(2).StringCellValue;
            if (Row.GetCell(GGStruct.WorkStyle).StringCellValue == "PM")
            {
                temRFcolumns.ErrorSrc = "PM";
                temRFcolumns.DealSrc = "PM";
            }
            else if (Row.GetCell(GGStruct.WorkStyle).StringCellValue.Substring(0, 2) == "升级")
            {
                temRFcolumns.ErrorSrc = "升级";
                temRFcolumns.DealSrc = "升级";
            }
            else
            {
                temRFcolumns.ErrorSrc = getCellValue(Row, GGStruct.ErrorSrc); //Row.GetCell(16).StringCellValue;
                temRFcolumns.DealSrc = getCellValue(Row, GGStruct.DealSrc); //Row.GetCell(17).StringCellValue;
            }
            string bank = getCellValue(Row,GGStruct.bank);
            //bank = parseBankName(bank);
            temRFcolumns.ClientName = getCellValue(Row,GGStruct.ClientName); //Row.GetCell(29).StringCellValue;
            temRFcolumns.Engineer = getCellValue(Row,GGStruct.Engineer); //Row.GetCell(20).StringCellValue;
                                                           //temRFcolumns.MaintenanceDate = getCellValue(Row, 24); //Row.GetCell(11).StringCellValue;
            string temDate = getCellValue(Row, GGStruct.MaintenanceDate); //Row.GetCell(11).StringCellValue;
            temRFcolumns.MaintenanceDate =temDate.Remove(temDate.IndexOf(' ')).Replace("-", "/");
            temRFcolumns.PaiID = getCellValue(Row, GGStruct.PaiID); //Row.GetCell(3).StringCellValue;
            temRFcolumns.ServiceID = getCellValue(Row, GGStruct.ServiceID); //Row.GetCell(4).StringCellValue;
            return temRFcolumns;
        }
        public ReportFileColumns AppendATMDataToColumnsFormat(int rowindex)
        {
            IRow Row = sourceWB.GetSheet("Sheet1").GetRow(rowindex);
            ReportFileColumns temRFcolumns = new ReportFileColumns();
            temRFcolumns.Model = "HT" + getCellValue(Row, AStruct.Model);//Row.GetCell(1).StringCellValue;
                                                             //temRFcolumns.Model = "HT-" + temRFcolumns.Model;
            temRFcolumns.ProductID = getCellValue(Row, AStruct.ProductID); //Row.GetCell(2).StringCellValue;
            string tempCellValue = Row.GetCell(AStruct.WorkStyle).StringCellValue;
            if (tempCellValue == "PM")
            {
                temRFcolumns.ErrorSrc = "PM";
                temRFcolumns.DealSrc = "PM";
            }
            else if (tempCellValue.Substring(0, 2) == "升级")
            {
                temRFcolumns.ErrorSrc = "升级";
                temRFcolumns.DealSrc = "升级";
            }
            else if( tempCellValue.IndexOf("协助对账")!=-1)
            {
              
                    temRFcolumns.ErrorSrc = tempCellValue.Substring(4);
                    temRFcolumns.DealSrc = "协助对账";
                
            }
            else
            {
                temRFcolumns.ErrorSrc = getCellValue(Row, AStruct.ErrorSrc); //Row.GetCell(16).StringCellValue;
                temRFcolumns.DealSrc = getCellValue(Row, AStruct.DealSrc); //Row.GetCell(17).StringCellValue;
            }
            string bank = getCellValue(Row, AStruct.bank);
            bank = parseBankName(bank);
            temRFcolumns.ClientName = bank + getCellValue(Row, AStruct.ClientName); //Row.GetCell(29).StringCellValue;
            temRFcolumns.Engineer = getCellValue(Row, AStruct.Engineer); //Row.GetCell(20).StringCellValue;
            temRFcolumns.MaintenanceDate = getCellValue(Row, AStruct.MaintenanceDate,true); //Row.GetCell(11).Stri
            temRFcolumns.PaiID = getCellValue(Row, AStruct.PaiID); //Row.GetCell(3).StringCellValue;
            temRFcolumns.ServiceID = getCellValue(Row, AStruct.ServiceID); //Row.GetCell(4).StringCellValue;
            return temRFcolumns;
        }
        public ReportFileColumns AppendSetupDataToColumnsFormat(int rowindex)
        {
            IRow Row = sourceWB.GetSheet("Sheet1").GetRow(rowindex);
            ReportFileColumns temRFcolumns = new ReportFileColumns();
            temRFcolumns.Model = "HT" + getCellValue(Row, 2);//Row.GetCell(1).StringCellValue;
                                                             //temRFcolumns.Model = "HT-" + temRFcolumns.Model;
            temRFcolumns.ProductID = getCellValue(Row, 1); //Row.GetCell(2).StringCellValue;
            
                temRFcolumns.ErrorSrc = "新装机";
                temRFcolumns.DealSrc = "新装机";

            string bank = getCellValue(Row, 24);
            bank = parseBankName(bank);
            temRFcolumns.ClientName = bank + getCellValue(Row, 26); //Row.GetCell(29).StringCellValue;
            temRFcolumns.Engineer = getCellValue(Row, 33); //Row.GetCell(20).StringCellValue;
           
            temRFcolumns.MaintenanceDate = getCellValue(Row,4,true); //Row.GetCell(11).Stri
            temRFcolumns.PaiID = getCellValue(Row, 43); //Row.GetCell(3).StringCellValue;
            temRFcolumns.ServiceID = getCellValue(Row, 0); //Row.GetCell(4).StringCellValue;
            return temRFcolumns;
        }
        private string parseBankName(string bank)
        {
            if (bank == "邮政公司")
                return "邮政";
            else if (bank == "广西农信")
                return "农信";
            else if (bank == "中国银行")
            {
                return "中行";
            }
            return "";
        }
        public string getCellValue(IRow row, int index)
        {
          
            return getCellValue(row,index,false);
        }
        public string getCellValue(IRow row, int index,bool isdate)
        {
            string value = "";
            ICell cell = row.GetCell(index);
            if (cell != null)
                if (cell.CellType == CellType.Numeric)
                {
                    double NumValue = cell.NumericCellValue;
                    if (isdate |(NumValue-(int)NumValue!=0))
                        value = cell.DateCellValue.ToString("yyyy/MM/dd");
                    else
                        value = cell.NumericCellValue.ToString();
                }
                else
                    value = cell.StringCellValue;
            value = (value == null) ? "" : value;
            return value;
        }




        private void AppendToXlscell(ReportFile RF, int index, ReportFileColumns RFC)
        {
            ISheet sheet1 = RF.hssfworkbook.GetSheet("Sheet1");
            int Rowindex = index;
            IRow newRow = sheet1.CreateRow(Rowindex);
            newRow.HeightInPoints = 20;
            ICellStyle cellstyle = RF.hssfworkbook.CreateCellStyle();
            cellstyle.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
            cellstyle.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
            cellstyle.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
            cellstyle.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
            cellstyle.VerticalAlignment = VerticalAlignment.Center;
          //  ICellStyle DatecellStype = RF.hssfworkbook.CreateCellStyle();
           // IDataFormat format = RF.hssfworkbook.CreateDataFormat();
            for (int i = 0; i < 13; i++)
            {

                ICell cell = newRow.CreateCell(i);

                cell.SetCellValue((string)RFC.GetType().GetFields()[i].GetValue(RFC));
                cell.CellStyle = cellstyle;
                /*  if (i == 10)
                  {
                      cell.SetCellValue((DateTime)RFC.GetType().GetFields()[i].GetValue(RFC));

                      DatecellStype.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
                      DatecellStype.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
                      DatecellStype.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
                      DatecellStype.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
                      DatecellStype.VerticalAlignment = VerticalAlignment.Center;
                      DatecellStype.DataFormat = format.GetFormat("m/d/yy");
                      cell.CellStyle = DatecellStype;
                      // cell.CellStyle.WrapText = true;
                  }
                  else
                  {*/


            }
            sheet1.ForceFormulaRecalculation = true;

        }

        private void AppendToXlscell(ReportFile RF)
        {
            ISheet sheet1 = RF.hssfworkbook.GetSheet("Sheet1");
            int RowCount = sheet1.LastRowNum + 1;
            HSSFWorkbook hssfworkbook = new HSSFWorkbook();
            ICellStyle cellstyle = hssfworkbook.CreateCellStyle();
            cellstyle.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
            cellstyle.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
            cellstyle.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
            cellstyle.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
            IRow newRow = sheet1.CreateRow(RowCount);
            for (int i = 0; i < 13; i++)
            {
                ICell cell = newRow.CreateCell(i);
                cell.SetCellValue((string)RFcolumns.GetType().GetFields()[i].GetValue(RFcolumns));
                cell.CellStyle = cellstyle;
                if (i == 5 | i == 6)
                {
                    cell.CellStyle.WrapText = true;
                }
            }
            sheet1.ForceFormulaRecalculation = true;
            RF.WriteToFile();
        }
    }
}
