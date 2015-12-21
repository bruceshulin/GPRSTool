using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace GPRSTOOL
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnImportExcel_Click(object sender, EventArgs e)
        {
            listGprs = new List<GPRSparam>();
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "*.*|*.*";
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                txtAddress.Text = ofd.FileName;
                Thread th = new Thread(new ParameterizedThreadStart(importExcel));
                th.Start(ofd.FileName);
            }
        }


        private void importExcel(object obj)
        {
            string filename = (string)obj;
            EpplusExcel2007Read(filename);
        }

        List<GPRSparam> listGprs = new List<GPRSparam>();

        private void EpplusExcel2007Read(string path)
        {
            try
            {
                //实例化一个计时器
                Stopwatch watch = new Stopwatch();
                //开始计时/*此处为要计算的运行代码
                watch.Start();
                //保存表头信息
                Dictionary<int, string> dictHeader = new Dictionary<int, string>();
                //文件信息
                FileInfo newFile = new FileInfo(path);
                using (ExcelPackage package = new ExcelPackage(newFile))
                {
                    string time = watch.ElapsedMilliseconds.ToString();
                    Console.WriteLine("加载完文件时间:" + time);
                    Lv.Log.Write("加载完文件时间: " + time, Lv.Log.MessageType.Info);

                    int vSheetCount = package.Workbook.Worksheets.Count; //获取总Sheet页
                    int page = 1;
                    for (int pagei = 1; pagei <= page; pagei++)
                    {
                        ExcelWorksheet worksheet = package.Workbook.Worksheets[2];//选定 指定页
                        time = watch.ElapsedMilliseconds.ToString();
                        Console.WriteLine("到打开表时间:" + time);
                        Lv.Log.Write("到打开表时间: " + time, Lv.Log.MessageType.Info);
                        watch.Stop();//结束计时

                        int colStart = worksheet.Dimension.Start.Column;//工作区开始列
                        int colEnd = worksheet.Dimension.End.Column;    //工作区结束列
                        int rowStart = worksheet.Dimension.Start.Row;   //工作区开始行号
                        int rowEnd = worksheet.Dimension.End.Row;       //工作区结束行号
                        //现在用不到，以后用得到

                        //1　每个表的样式都不一样，如果统一起来处理的话不好处理，怎么处理呢？
                        //传每个表的类型过来

                        int off = 0;    //起始行偏移量

                        //将每列标题添加到字典中
                        for (int i = colStart; i <= colEnd; i++)
                        {
                            if (worksheet.Cells[i,rowStart ].Value == null)
                            {
                                continue;
                            }
                            string titlestr = worksheet.Cells[ i , rowStart ].Value.ToString();
                            if (titlestr == null || titlestr == "")
                            {
                                continue;
                            }
                            dictHeader[i] = titlestr.Replace(" ", "");
                        }
                        off += 1;
                        int count = 0;
                        //遍历每一列
                        for (int col = colStart + off; col <= colEnd; col++)
                        {
                            Dictionary<int,string> dictHeadervalue = new Dictionary<int,string>();
                            count++;
                            //遍历每一列的单元格
                            for (int row = rowStart; row <= rowEnd; row++)
                            {
                                string text = "";
                                //得到单元格信息
                                ExcelRange cell = null;
                                try
                                {
                                    cell = worksheet.Cells[row, col];
                                }
                                catch (Exception err)
                                {
                                    text = "";
                                    Console.WriteLine("" + err.Message);
                                    Lv.Log.Write("提取单元数据出错　row" + col.ToString() + " col" + row.ToString() + err.Message, Lv.Log.MessageType.Error);
                                }
                                if (cell.Value == null)
                                {
                                    text = "";
                                }else
                                { 
                                    text = cell.RichText.Text;
                                }
                                // dictHeadervalue[col] = text;         //标题，值 不用标题了可以做判断
                                //对每一个网络参数进行修正操作
                                if (row>=34)
                                {
                                    break;
                                }
                                text = revisedValue(text, dictHeader[row]);
                                dictHeadervalue.Add(row,text);
                            }
                            
                            //保存 当前列的gprs mms
                            ConvertGprs(dictHeadervalue,dictHeader);
                        }
                        Console.WriteLine("总处理列数:" + count);
                        Lv.Log.Write("总处理行数: " + count, Lv.Log.MessageType.Info);
                    }
                }
            }
            catch (Exception err)
            {
                Console.WriteLine("加载excel出错了　" + err.Message);
                Lv.Log.Write("加载excel出错了　 " + err.Message, Lv.Log.MessageType.Error);
            }
        }

        private void ConvertGprs(Dictionary<int, string> dictHeadervalue, Dictionary<int, string> dictHeader)
        {
            //gprs
            GPRSparam gprs = new GPRSparam();
            gprs.Type = GPRSTYPE.GPRS;
            gprs.Typestr = "gprs";
            gprs.Country = GetValue(dictHeadervalue, 1);
            gprs.Mvno_type = GetValue(dictHeadervalue, 3);
            gprs.Mvno_match_data = GetValue(dictHeadervalue, 4);
            gprs.Idledisplay = GetValue(dictHeadervalue, 5);
            gprs.Name = GetValue(dictHeadervalue, 6);
            gprs.Apn = GetValue(dictHeadervalue, 7);
            gprs.Proxy = GetValue(dictHeadervalue, 8);
            gprs.Port = GetValue(dictHeadervalue, 9);
            gprs.Username = GetValue(dictHeadervalue, 10);
            gprs.Password = GetValue(dictHeadervalue, 11);
            gprs.Server = GetValue(dictHeadervalue, 12);
            gprs.Mmsc = GetValue(dictHeadervalue, 13);
            gprs.Mmsproxy = GetValue(dictHeadervalue, 14);
            gprs.Mmsport = GetValue(dictHeadervalue, 15);
            gprs.Mcc = GetValue(dictHeadervalue, 16);
            gprs.Mnc = GetValue(dictHeadervalue, 17);
            gprs.Authtype = GetValue(dictHeadervalue, 18);
            gprs.Apntype = GetValue(dictHeadervalue, 19);

            listGprs.Add(gprs);
            //mms
            GPRSparam mms = new GPRSparam();
            mms.Typestr = "mms";
            mms.Type = GPRSTYPE.MMS;
            mms.Country = GetValue(dictHeadervalue, 1);
            mms.Mvno_type = GetValue(dictHeadervalue, 3);
            mms.Mvno_match_data = GetValue(dictHeadervalue, 4);
            mms.Idledisplay = GetValue(dictHeadervalue, 5);
            mms.Name = GetValue(dictHeadervalue, 21);
            mms.Apn = GetValue(dictHeadervalue, 22);
            mms.Proxy = GetValue(dictHeadervalue, 23);
            mms.Port = GetValue(dictHeadervalue, 24);
            mms.Username = GetValue(dictHeadervalue, 25);
            mms.Password = GetValue(dictHeadervalue, 26);
            mms.Server = GetValue(dictHeadervalue, 27);
            mms.Mmsc = GetValue(dictHeadervalue, 28);
            mms.Mmsproxy = GetValue(dictHeadervalue, 29);
            mms.Mmsport = GetValue(dictHeadervalue, 30);
            mms.Mcc = GetValue(dictHeadervalue,31);
            mms.Mnc = GetValue(dictHeadervalue, 32);
            mms.Authtype = GetValue(dictHeadervalue, 33);
            mms.Apntype = GetValue(dictHeadervalue, 34);

            listGprs.Add(mms);
        }

        private string GetValue(Dictionary<int, string> dictHeadervalue, int index)
        {
            if (dictHeadervalue.ContainsKey(index))
            {
                return dictHeadervalue[index];
            }
            else
            {
                return "";
            }
        }

        private int GetHeader(Dictionary<int, string> dictHeader, string name)
        {
            foreach (var item in dictHeader)
            {
                if (item.Value == name)
                {
                    return item.Key;
                }
            }
            return -1;
        }



        /// <summary>
        /// 修正值
        /// </summary>
        /// <param name="text"></param>
        /// <param name="p"></param>
        /// <returns></returns>
        private string revisedValue(string text, string title)
        {
            switch (title)
            {
                case "Tecno/itel": break;
                case "GPRS/EDGE/Internet": break;
                case "mvno_type": break;
                case "mvno_match_data": break;
                case "insertSIMidledisplay": break;
                case "NAME": break;
                case "APN": break;
                case "PROXY":
                    if (text == "")
                    {
                        text = "0.0.0.0";
                    }
                    System.Text.RegularExpressions.Match mc = System.Text.RegularExpressions.Regex.Match(text, @"[\d]{1,3}.[\d]{1,3}.[\d]{1,3}.[\d]{1,3}", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                    if (mc.Success == false)
                    {
                        MessageBox.Show("不是ip地址" + text);
                    }
                    break;
                case "PORT":
                    if (text == "")
                    {
                        text = "0";
                    }
                    break;
                case "USERNAME": break;
                case "PASSWORD": break;
                case "SERVER": break;
                case "MMSC":
                    //homepage
                    break;
                case "MMSPROXY": 
                    if (text == "")
                    {
                        text = "0.0.0.0";
                    }
                    System.Text.RegularExpressions.Match mmsmc = System.Text.RegularExpressions.Regex.Match(text, @"[\d]{1,3}.[\d]{1,3}.[\d]{1,3}.[\d]{1,3}", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                    if (mmsmc.Success == false)
                    {
                        MessageBox.Show("mms不是ip地址" + text);
                    }
                    break;
                case "MMSPORT":
                    if (text == "")
                    {
                        text = "0";
                    }
                    break;
                case "MCC": break;
                case "MNC": break;
                case "AUTHENTICATIONTYPE":
                    string tmptext = text.ToLower();
                    if (tmptext == "")
                    {
                        text = "0";
                    }
                    else if (tmptext.Contains("none") == true)
                    {
                        text = "0";
                    }
                    else if (tmptext.Contains("not set") == true)
                    {
                        text = "0";
                    }
                    else if (tmptext.Contains("pap") == true)
                    {
                        text = "0";
                    }
                    else if (tmptext.Contains("chap") == true)
                    {
                        text = "1";
                    }
                    else
                    {
                        text = "1";
                    }
                    break;
                case "APNTYPE": break;
                case "MMS": break;
                default:
                    break;
            }

            Dictionary<string,string> listfilter = new Dictionary<string,string>();
            listfilter.Add("For Browser only","");
            foreach (var item in listfilter)
            {
                if (text.Contains(item.Key) == true)
                {
                    text = item.Value;
                }
            }
            return text.Trim();
        }

        private void btnOutPutExcel_Click(object sender, EventArgs e)
        {
            string path ="GPRS" + DateTime.Now.Minute.ToString() + DateTime.Now.Second.ToString() + ".xlsx";
            Thread th = new Thread(new ParameterizedThreadStart(NPOIOutExcel));
            th.Start(path);
        }
        private void NPOIOutExcel(object obj)
        {
            if (listGprs.Count < 1)
            {
                MessageBox.Show("请加载gprs络参数文件!");
                return;
            }
            string[] urls = this.textBox1.Text.Split('\n');
            if (urls.Length == 0)
            {
                urls[0] = "homepage";
            }
            for (int i = 0; i < urls.Length; i++)
            {
                urls[i] = urls[i].Replace('\r', ' ').Trim();
            }
            foreach (string url in urls)
            {
                string M = System.Text.RegularExpressions.Regex.Match(url, @"&M=(?<m>[\w]*)", System.Text.RegularExpressions.RegexOptions.IgnoreCase).Groups["m"].Value;
                string Z = System.Text.RegularExpressions.Regex.Match(url, @"&Z=(?<z>[\w]*)", System.Text.RegularExpressions.RegexOptions.IgnoreCase).Groups["z"].Value;
                string path = Z + M + ".xls";
                if (path == "")
                {
                    path = "test";
                }
                System.IO.FileInfo file = new System.IO.FileInfo((string)path);

                try
                {
                    DataTable dt = new DataTable();
                    HSSFWorkbook workbook = new HSSFWorkbook();
                    ISheet sheet = workbook.CreateSheet("NTAC");
                    ISheet sheet1 = workbook.CreateSheet("NTACHdr");
                    ISheet sheet2 = workbook.CreateSheet("Sheet3");

                    ICellStyle HeadercellStyle = workbook.CreateCellStyle();
                    HeadercellStyle.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
                    HeadercellStyle.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
                    HeadercellStyle.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
                    HeadercellStyle.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
                    HeadercellStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
                    //字体
                    NPOI.SS.UserModel.IFont headerfont = workbook.CreateFont();
                    headerfont.Boldweight = (short)FontBoldWeight.Bold;
                    HeadercellStyle.SetFont(headerfont);

                    List<string> tabletitle = new List<string>();
                    tabletitle.Add("Setting Name");
                    tabletitle.Add("MCC");
                    tabletitle.Add("MNC");
                    tabletitle.Add("Account Type");
                    tabletitle.Add("APN");
                    tabletitle.Add("Access Type");
                    tabletitle.Add("Access Option");
                    tabletitle.Add("Proxy Server IP");
                    tabletitle.Add("Port");
                    tabletitle.Add("User Name");
                    tabletitle.Add("Password");
                    tabletitle.Add("First DNS");
                    tabletitle.Add("Second DNS");
                    tabletitle.Add("Home Page");
                    tabletitle.Add("Auth Type");
                    tabletitle.Add("Reserved");
                    //用column name 作为列名
                    int icolIndex = 0;
                    IRow headerRow = sheet.CreateRow(0);
                    foreach (string item in tabletitle)
                    {
                        ICell cell = headerRow.CreateCell(icolIndex);
                        cell.SetCellValue(item);
                        cell.CellStyle = HeadercellStyle;
                        icolIndex++;
                    }

                    ICellStyle cellStyle = workbook.CreateCellStyle();

                    //为避免日期格式被Excel自动替换，所以设定 format 为 『@』 表示一率当成text來看
                    cellStyle.DataFormat = HSSFDataFormat.GetBuiltinFormat("@");
                    cellStyle.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
                    cellStyle.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
                    cellStyle.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
                    cellStyle.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;

                    NPOI.SS.UserModel.IFont cellfont = workbook.CreateFont();
                    cellfont.Boldweight = (short)FontBoldWeight.Normal;
                    cellStyle.SetFont(cellfont);

                    //建立内容行
                    int iRowIndex = 1;
                    int iCellIndex = 0;
                    foreach (GPRSparam item in listGprs)
                    {
                        IRow DataRow = sheet.CreateRow(iRowIndex);
                        List<string> itemValues = new List<string>();
                        if (item.Mcc == "" || item.Mnc == "" || item.Name == "")
                        {
                            continue;
                        }
                        itemValues.Add(item.Name);
                        itemValues.Add(item.Mcc);
                        itemValues.Add(item.Mnc);
                        itemValues.Add(item.Type == GPRSTYPE.GPRS ? "0" : "1");//gprs 0 mms 1 java 2 dcd 3
                        itemValues.Add(item.Apn);
                        itemValues.Add("1");
                        itemValues.Add("0");
                        itemValues.Add(item.Type == GPRSTYPE.GPRS ? item.Proxy : item.Mmsproxy);        //短信的和gprs的不一样
                        itemValues.Add(item.Type == GPRSTYPE.GPRS ? item.Port : item.Mmsport);         //短信的和gprs的不一样
                        itemValues.Add(item.Username);
                        itemValues.Add(item.Password);
                        itemValues.Add("0.0.0.0");
                        itemValues.Add("0.0.0.0");
                        if (item.Type == GPRSTYPE.GPRS)
                        {
                            if (item.Server != "")
                            {
                                itemValues.Add(item.Server);      //homepage
                            }
                            else
                            {
                                itemValues.Add(url);
                            }
                        }
                        else
                        {
                            if (item.Mmsc != "")
                            {
                                itemValues.Add(item.Mmsc);      //homepage
                            }
                            else
                            {
                                itemValues.Add(url);
                            }
                        }
                        itemValues.Add(item.Authtype);    //#PAP 0 CHAP 1  这个更据表上来
                        itemValues.Add("0");

                        foreach (string itemvalues in itemValues)
                        {
                            ICell cell = DataRow.CreateCell(iCellIndex);
                            cell.SetCellValue(itemvalues);
                            cell.CellStyle = cellStyle;
                            iCellIndex++;
                        }
                        iCellIndex = 0;
                        iRowIndex++;
                    }
                    //加载固定网络参数
                    Dictionary<int, List<string>> fixGprs = new Dictionary<int, List<string>>();
                    ReadFixGprs(ref fixGprs, url);
                    foreach (var itemfix in fixGprs)
	                {
                        IRow DataRow = sheet.CreateRow(iRowIndex);
                        foreach (string itemvalues in itemfix.Value)
                        {
                            ICell cell = DataRow.CreateCell(iCellIndex);
                            cell.SetCellValue(itemvalues);
                            cell.CellStyle = cellStyle;
                            iCellIndex++;
                        }
                        iCellIndex = 0;
                        iRowIndex++;
                    }

                    //自适应列宽度
                    for (int i = 0; i < icolIndex; i++)
                    {
                        sheet.AutoSizeColumn(i);
                    }

                    //写Excel
                    FileStream fileIO = new FileStream(path, FileMode.OpenOrCreate);
                    workbook.Write(fileIO);
                    fileIO.Flush();
                    fileIO.Close();
                    //MESSAGE
                    MessageBox.Show("导出完成！ 文件：" + path);
                }
                catch (Exception ex)
                {
                    //MESSAGE
                    MessageBox.Show("导出出错" + ex.Message);
                }
                finally { }
            }
        }

        /// <summary>
        /// 加入固定网络参数
        /// </summary>
        /// <param name="itemValues"></param>
        private void ReadFixGprs(ref Dictionary<int, List<string>> dicgprs, string url)
        {
            int row = 0;
            List<string> itemValues = new List<string>();
            //--1
            itemValues.Add("中国移动梦网");
            itemValues.Add("460");
            itemValues.Add("0");
            itemValues.Add("0");//gprs 0 mms 1 java 2 dcd 3
            itemValues.Add("cmwap");
            itemValues.Add("1");
            itemValues.Add("0");
            itemValues.Add("10.0.0.172");        //短信的和gprs的不一样
            itemValues.Add("80");         //短信的和gprs的不一样
            itemValues.Add("");
            itemValues.Add("");
            itemValues.Add("0.0.0.0");
            itemValues.Add("0.0.0.0");
            itemValues.Add(url);        //homepage
            itemValues.Add("0");    //#PAP 0 CHAP 1  这个更据表上来
            itemValues.Add("0");
            
            dicgprs.Add(row, itemValues);
            itemValues = new List<string>();
            row++;
            //--2
            itemValues.Add("中国移动彩信");
            itemValues.Add("460");
            itemValues.Add("0");
            itemValues.Add("1");//gprs 0 mms 1 java 2 dcd 3
            itemValues.Add("cmwap");
            itemValues.Add("1");
            itemValues.Add("0");
            itemValues.Add("10.0.0.172");        //短信的和gprs的不一样
            itemValues.Add("80");         //短信的和gprs的不一样
            itemValues.Add("");
            itemValues.Add("");
            itemValues.Add("0.0.0.0");
            itemValues.Add("0.0.0.0");
            itemValues.Add("http://mmsc.monternet.com");        //homepage
            itemValues.Add("0");    //#PAP 0 CHAP 1  这个更据表上来
            itemValues.Add("0");

            dicgprs.Add(row, itemValues);
            itemValues = new List<string>();
            row++;

            //#-3
            itemValues.Add( "中国移动互联网");
            itemValues.Add( "460");
            itemValues.Add( "0");
            itemValues.Add( "0");//#在软件上有四个选项browser 0 mms 1 java 2 dcd 3//这里都是GPRS的网络参数，所以都是 0
            itemValues.Add( "cmnet");
            itemValues.Add( "1");//#在软件上有两个选项 wap1.2(wsp) 0 wap2.0(http) 1 这里填1，大部分都是这相，所以用这个
            itemValues.Add( "0");//#editoption 0  ReadOnly 1  以前全都是0
            itemValues.Add( "0.0.0.0");
            itemValues.Add( "0");
            itemValues.Add( "");
            itemValues.Add( "");
            itemValues.Add( "0.0.0.0");//#全都是 0.0.0.0
            itemValues.Add( "0.0.0.0");//#全都是 0.0.0.0
            itemValues.Add( url);//#每个项目都需要设置自身的homepage 
            itemValues.Add( "0"); //#PAP 0 CHAP 1  这个更据表上来
            itemValues.Add( "0");	//#保留字段，都是0

            dicgprs.Add(row, itemValues);
            itemValues = new List<string>();
            row++;
            //#--4
            itemValues.Add("中国移动快讯");
            itemValues.Add("460");
            itemValues.Add("0");
            itemValues.Add("3");//#在软件上有四个选项browser 0 mms 1 java 2 dcd 3//这里都是GPRS的网络参数，所以都是 0
            itemValues.Add("cmwap");
            itemValues.Add("1");//#在软件上有两个选项 wap1.2(wsp) 0 wap2.0(http) 1 这里填1，大部分都是这相，所以用这个
            itemValues.Add("0");//#editoption 0  ReadOnly 1  以前全都是0
            itemValues.Add("10.0.0.172");
            itemValues.Add("80");
            itemValues.Add("");
            itemValues.Add("");
            itemValues.Add("0.0.0.0");//#全都是 0.0.0.0
            itemValues.Add("0.0.0.0");//#全都是 0.0.0.0
            itemValues.Add(url);//#每个项目都需要设置自身的homepage 
            itemValues.Add("0"); //#PAP 0 CHAP 1  这个更据表上来
            itemValues.Add("0");	//#保留字段，都是0

            dicgprs.Add(row, itemValues);
            itemValues = new List<string>();
            row++;

            //#-5
           itemValues.Add("中国联通WAP");
           itemValues.Add("460");
           itemValues.Add("1");
           itemValues.Add("0");//#在软件上有四个选项browser 0 mms 1 java 2 dcd 3//这里都是GPRS的网络参数，所以都是 0
           itemValues.Add("uniwap");
           itemValues.Add("1");//#在软件上有两个选项 wap1.2(wsp) 0 wap2.0(http) 1 这里填1，大部分都是这相，所以用这个
           itemValues.Add("0");//#editoption 0  ReadOnly 1  以前全都是0
           itemValues.Add("10.0.0.172");
           itemValues.Add("80");
           itemValues.Add("");
           itemValues.Add("");
           itemValues.Add("0.0.0.0");//#全都是 0.0.0.0
           itemValues.Add("0.0.0.0");//#全都是 0.0.0.0
           itemValues.Add(url);//#每个项目都需要设置自身的homepage
           itemValues.Add("0"); //#PAP 0 CHAP 1  这个更据表上来
           itemValues.Add("0");//	#保留字段，都是0

           dicgprs.Add(row, itemValues);
           itemValues = new List<string>();
           row++;
            // #-6
            itemValues.Add("中国联通彩信");
            itemValues.Add("460");
            itemValues.Add("1");
            itemValues.Add("1");//#在软件上有四个选项browser 0 mms 1 java 2 dcd 3//这里都是GPRS的网络参数，所以都是 0
            itemValues.Add("uniwap");
            itemValues.Add("1");//#在软件上有两个选项 wap1.2(wsp) 0 wap2.0(http) 1 这里填1，大部分都是这相，所以用这个
            itemValues.Add("0");//#editoption 0  ReadOnly 1  以前全都是0
            itemValues.Add("10.0.0.172");
            itemValues.Add("80");
            itemValues.Add("");
            itemValues.Add("");
            itemValues.Add("0.0.0.0");//#全都是 0.0.0.0
            itemValues.Add("0.0.0.0");//#全都是 0.0.0.0
            itemValues.Add("http://mmsc.myuni.com.cn");//#每个项目都需要设置自身的homepage
            itemValues.Add("0"); //#PAP 0 CHAP 1  这个更据表上来
            itemValues.Add("0");//#保留字段，都是0

            dicgprs.Add(row, itemValues);
            itemValues = new List<string>();
            row++;
            //#-7
            itemValues.Add("中国联通互联网");
            itemValues.Add("460");
            itemValues.Add("1");
            itemValues.Add("0");//#在软件上有四个选项browser 0 mms 1 java 2 dcd 3//这里都是GPRS的网络参数，所以都是 0
            itemValues.Add("uninet");
            itemValues.Add("1");//#在软件上有两个选项 wap1.2(wsp) 0 wap2.0(http) 1 这里填1，大部分都是这相，所以用这个
            itemValues.Add("0");//#editoption 0  ReadOnly 1  以前全都是0
            itemValues.Add("0.0.0.0");
            itemValues.Add("0");
            itemValues.Add("");
            itemValues.Add("");
            itemValues.Add("0.0.0.0");//#全都是 0.0.0.0
            itemValues.Add("0.0.0.0");//#全都是 0.0.0.0
            itemValues.Add(url);//#每个项目都需要设置自身的homepage
            itemValues.Add("0"); //#PAP 0 CHAP 1  这个更据表上来
            itemValues.Add("0");	//#保留字段，都是0

            dicgprs.Add(row, itemValues);
            itemValues = new List<string>();
            row++;
        }

        private void OutPutExcel(object obj)
        {
            

            string[] urls = this.textBox1.Text.Split('\n');
            if (urls.Length == 0 )
            {
                urls[0] = "homepage";
            }
            for (int i = 0; i < urls.Length; i++)
            {
                urls[i] = urls[i].Replace('\r',' ').Trim();
            }
            foreach (string url in urls)
            {
               // url = url.Replace('\r','');
                string M = System.Text.RegularExpressions.Regex.Match(url, @"&M=(?<m>[\w]*)", System.Text.RegularExpressions.RegexOptions.IgnoreCase).Groups["m"].Value;    
                string Z = System.Text.RegularExpressions.Regex.Match(url, @"&Z=(?<z>[\w]*)", System.Text.RegularExpressions.RegexOptions.IgnoreCase).Groups["z"].Value;
                string path = Z + M + ".xlsx";
                System.IO.FileInfo file = new System.IO.FileInfo((string)path);
                int row = 2;
                OfficeOpenXml.ExcelPackage ep = new OfficeOpenXml.ExcelPackage(file);
                OfficeOpenXml.ExcelWorkbook wb = ep.Workbook;
                OfficeOpenXml.ExcelWorksheet ws = wb.Worksheets.Add("NTAC");
                /*
                 //配置文件属性
                wb.Properties.Category = "类别";
                wb.Properties.Author = "作者";
                wb.Properties.Comments = "备注";
                wb.Properties.Company = "公司";
                wb.Properties.Keywords = "关键字";
                wb.Properties.Manager = "管理者";
                wb.Properties.Status = "内容状态";
                wb.Properties.Subject = "主题";
                wb.Properties.Title = "标题";
                wb.Properties.LastModifiedBy = "最后一次保存者";
                 */
                //写标题
                ws.Cells[1, 1].Value = "Setting Name";
                ws.Cells[1, 2].Value = "MCC";
                ws.Cells[1, 3].Value = "MNC";
                ws.Cells[1, 4].Value = "Account Type";
                ws.Cells[1, 5].Value = "APN";
                ws.Cells[1, 6].Value = "Access Type";
                ws.Cells[1, 7].Value = "Access Option";
                ws.Cells[1, 8].Value = "Proxy Server IP";
                ws.Cells[1, 9].Value = "Port";
                ws.Cells[1, 10].Value = "User Name";
                ws.Cells[1, 11].Value = "Password";
                ws.Cells[1, 12].Value = "First DNS";
                ws.Cells[1, 13].Value = "Second DNS";
                ws.Cells[1, 14].Value = "Home Page";
                ws.Cells[1, 15].Value = "Auth Type";
                ws.Cells[1, 16].Value = "Reserved";

                //写内容

                foreach (GPRSparam item in listGprs)
                {
                    if (item.Mcc == "" || item.Mnc == "" || item.Name == "")
                    {
                        continue;
                    }
                    ws.Cells[row, 1].Value = item.Name;
                    ws.Cells[row, 2].Value = item.Mcc;
                    ws.Cells[row, 3].Value = item.Mnc;
                    ws.Cells[row, 4].Value = item.Type == GPRSTYPE.GPRS ? "0" : "1";//gprs 0 mms 1 java 2 dcd 3
                    ws.Cells[row, 5].Value = item.Apn;
                    ws.Cells[row, 6].Value = "1";
                    ws.Cells[row, 7].Value = "0";
                    ws.Cells[row, 8].Value = item.Type == GPRSTYPE.GPRS ? item.Proxy : item.Mmsproxy;        //短信的和gprs的不一样
                    ws.Cells[row, 9].Value = item.Type == GPRSTYPE.GPRS ? item.Port : item.Mmsport;         //短信的和gprs的不一样
                    ws.Cells[row, 10].Value = item.Username;
                    ws.Cells[row, 11].Value = item.Password;
                    ws.Cells[row, 12].Value = "0.0.0.0";
                    ws.Cells[row, 13].Value = "0.0.0.0";
                    if (item.Type == GPRSTYPE.GPRS)
                    {
                        if (item.Server != "")
                        {
                            ws.Cells[row, 14].Value = item.Server;      //homepage
                        }
                        else
                        {
                            ws.Cells[row, 14].Value = url;
                        }
                    }
                    else
                    {
                        if (item.Mmsc != "")
                        {
                            ws.Cells[row, 14].Value = item.Mmsc;      //homepage
                        }
                        else
                        {
                            ws.Cells[row, 14].Value = url;
                        }
                    }
                    ws.Cells[row, 15].Value = item.Authtype;    //#PAP 0 CHAP 1  这个更据表上来
                    ws.Cells[row, 16].Value = "0";
                    row++;
                }
                ADDFIXITEM(ws, row, url);
                ws = wb.Worksheets.Add("NTACHdr");
                ws.Cells[1, 1].Value = "magic";
                ws.Cells[1, 2].Value = "version";
                ws.Cells[1, 3].Value = "nCount";
                ws = wb.Worksheets.Add("Sheet3");
                ep.Save();
            }
            MessageBox.Show("导出完成！");
        }

        private void ADDFIXITEM(ExcelWorksheet ws,int row,string url)
        {
            //--1
            ws.Cells[row, 1].Value = "中国移动梦网";
            ws.Cells[row, 2].Value = "460";
            ws.Cells[row, 3].Value = "0";
            ws.Cells[row, 4].Value = "0";//gprs 0 mms 1 java 2 dcd 3
            ws.Cells[row, 5].Value = "cmwap";
            ws.Cells[row, 6].Value = "1";
            ws.Cells[row, 7].Value = "0";
            ws.Cells[row, 8].Value = "10.0.0.172";        //短信的和gprs的不一样
            ws.Cells[row, 9].Value = "80";         //短信的和gprs的不一样
            ws.Cells[row, 10].Value = "";
            ws.Cells[row, 11].Value = "";
            ws.Cells[row, 12].Value = "0.0.0.0";
            ws.Cells[row, 13].Value = "0.0.0.0";
            ws.Cells[row, 14].Value = url;        //homepage
            ws.Cells[row, 15].Value = "0";    //#PAP 0 CHAP 1  这个更据表上来
            ws.Cells[row, 16].Value = "0";
            row++;
            //--2
            ws.Cells[row, 1].Value = "中国移动彩信";
            ws.Cells[row, 2].Value = "460";
            ws.Cells[row, 3].Value = "0";
            ws.Cells[row, 4].Value = "1";//gprs 0 mms 1 java 2 dcd 3
            ws.Cells[row, 5].Value = "cmwap";
            ws.Cells[row, 6].Value = "1";
            ws.Cells[row, 7].Value = "0";
            ws.Cells[row, 8].Value = "10.0.0.172";        //短信的和gprs的不一样
            ws.Cells[row, 9].Value = "80";         //短信的和gprs的不一样
            ws.Cells[row, 10].Value = "";
            ws.Cells[row, 11].Value = "";
            ws.Cells[row, 12].Value = "0.0.0.0";
            ws.Cells[row, 13].Value = "0.0.0.0";
            ws.Cells[row, 14].Value = "http://mmsc.monternet.com";        //homepage
            ws.Cells[row, 15].Value = "0";    //#PAP 0 CHAP 1  这个更据表上来
            ws.Cells[row, 16].Value = "0";
            row++;

            //#-3
            ws.Cells[row, 1].Value =  "中国移动互联网";
            ws.Cells[row, 2].Value =  "460";
            ws.Cells[row, 3].Value =  "0";
            ws.Cells[row, 4].Value =  "0";//#在软件上有四个选项browser 0 mms 1 java 2 dcd 3//这里都是GPRS的网络参数，所以都是 0
            ws.Cells[row, 5].Value =  "cmnet";
            ws.Cells[row, 6].Value =  "1";//#在软件上有两个选项 wap1.2(wsp) 0 wap2.0(http) 1 这里填1，大部分都是这相，所以用这个
            ws.Cells[row, 7].Value =  "0";//#editoption 0  ReadOnly 1  以前全都是0
            ws.Cells[row, 8].Value =  "0.0.0.0";
            ws.Cells[row, 9].Value =  "0";
            ws.Cells[row, 10].Value = "";
            ws.Cells[row, 11].Value = "";
            ws.Cells[row, 12].Value = "0.0.0.0";//#全都是 0.0.0.0
            ws.Cells[row, 13].Value = "0.0.0.0";//#全都是 0.0.0.0
            ws.Cells[row, 14].Value = url;//#每个项目都需要设置自身的homepage 
            ws.Cells[row, 15].Value = "0"; //#PAP 0 CHAP 1  这个更据表上来
            ws.Cells[row, 16].Value = "0";	//#保留字段，都是0
            row++;
            //#--4
            ws.Cells[row, 1].Value =   "中国移动快讯";
            ws.Cells[row, 2].Value =   "460";
            ws.Cells[row, 3].Value =   "0";
            ws.Cells[row, 4].Value =   "3";//#在软件上有四个选项browser 0 mms 1 java 2 dcd 3//这里都是GPRS的网络参数，所以都是 0
            ws.Cells[row, 5].Value =   "cmwap";
            ws.Cells[row, 6].Value =   "1";//#在软件上有两个选项 wap1.2(wsp) 0 wap2.0(http) 1 这里填1，大部分都是这相，所以用这个
            ws.Cells[row, 7].Value =   "0";//#editoption 0  ReadOnly 1  以前全都是0
            ws.Cells[row, 8].Value =   "10.0.0.172";
            ws.Cells[row, 9].Value =   "80";
            ws.Cells[row, 10].Value =  "";
            ws.Cells[row, 11].Value =  "";
            ws.Cells[row, 12].Value =  "0.0.0.0";//#全都是 0.0.0.0
            ws.Cells[row, 13].Value =  "0.0.0.0";//#全都是 0.0.0.0
            ws.Cells[row, 14].Value =  url;//#每个项目都需要设置自身的homepage 
            ws.Cells[row, 15].Value =  "0"; //#PAP 0 CHAP 1  这个更据表上来
            ws.Cells[row, 16].Value =  "0";	//#保留字段，都是0
            row++;

            //#-5
            ws.Cells[row, 1].Value =   "中国联通WAP";
            ws.Cells[row, 2].Value =   "460";
            ws.Cells[row, 3].Value =   "1";
            ws.Cells[row, 4].Value =   "0";//#在软件上有四个选项browser 0 mms 1 java 2 dcd 3//这里都是GPRS的网络参数，所以都是 0
            ws.Cells[row, 5].Value =   "uniwap";
            ws.Cells[row, 6].Value =   "1";//#在软件上有两个选项 wap1.2(wsp) 0 wap2.0(http) 1 这里填1，大部分都是这相，所以用这个
            ws.Cells[row, 7].Value =   "0";//#editoption 0  ReadOnly 1  以前全都是0
            ws.Cells[row, 8].Value =   "10.0.0.172";
            ws.Cells[row, 9].Value =   "80";
            ws.Cells[row, 10].Value =  "";
            ws.Cells[row, 11].Value =  "";
            ws.Cells[row, 12].Value =  "0.0.0.0";//#全都是 0.0.0.0
            ws.Cells[row, 13].Value =  "0.0.0.0";//#全都是 0.0.0.0
            ws.Cells[row, 14].Value =  url;//#每个项目都需要设置自身的homepage
            ws.Cells[row, 15].Value =  "0"; //#PAP 0 CHAP 1  这个更据表上来
            ws.Cells[row, 16].Value =  "0";//	#保留字段，都是0
            row++;
           // #-6
            ws.Cells[row, 1].Value =    "中国联通彩信";
            ws.Cells[row, 2].Value =    "460";
            ws.Cells[row, 3].Value =    "1";
            ws.Cells[row, 4].Value =    "1";//#在软件上有四个选项browser 0 mms 1 java 2 dcd 3//这里都是GPRS的网络参数，所以都是 0
            ws.Cells[row, 5].Value =    "uniwap";
            ws.Cells[row, 6].Value =    "1";//#在软件上有两个选项 wap1.2(wsp) 0 wap2.0(http) 1 这里填1，大部分都是这相，所以用这个
            ws.Cells[row, 7].Value =    "0";//#editoption 0  ReadOnly 1  以前全都是0
            ws.Cells[row, 8].Value =    "10.0.0.172";
            ws.Cells[row, 9].Value =    "80";
            ws.Cells[row, 10].Value =   "";
            ws.Cells[row, 11].Value =   "";
            ws.Cells[row, 12].Value =   "0.0.0.0";//#全都是 0.0.0.0
            ws.Cells[row, 13].Value =   "0.0.0.0";//#全都是 0.0.0.0
            ws.Cells[row, 14].Value =   "http://mmsc.myuni.com.cn";//#每个项目都需要设置自身的homepage
            ws.Cells[row, 15].Value =   "0"; //#PAP 0 CHAP 1  这个更据表上来
            ws.Cells[row, 16].Value =   "0";//#保留字段，都是0
            row++;
            //#-7
            ws.Cells[row, 1].Value =   "中国联通互联网";
            ws.Cells[row, 2].Value =   "460";
            ws.Cells[row, 3].Value =   "1";
            ws.Cells[row, 4].Value =   "0";//#在软件上有四个选项browser 0 mms 1 java 2 dcd 3//这里都是GPRS的网络参数，所以都是 0
            ws.Cells[row, 5].Value =   "uninet";
            ws.Cells[row, 6].Value =   "1";//#在软件上有两个选项 wap1.2(wsp) 0 wap2.0(http) 1 这里填1，大部分都是这相，所以用这个
            ws.Cells[row, 7].Value =   "0";//#editoption 0  ReadOnly 1  以前全都是0
            ws.Cells[row, 8].Value =   "0.0.0.0";
            ws.Cells[row, 9].Value =   "0";
            ws.Cells[row, 10].Value =  "";
            ws.Cells[row, 11].Value =  "";
            ws.Cells[row, 12].Value =  "0.0.0.0";//#全都是 0.0.0.0
            ws.Cells[row, 13].Value =  "0.0.0.0";//#全都是 0.0.0.0
            ws.Cells[row, 14].Value =  url;//#每个项目都需要设置自身的homepage
            ws.Cells[row, 15].Value =  "0"; //#PAP 0 CHAP 1  这个更据表上来
            ws.Cells[row, 16].Value =  "0";	//#保留字段，都是0
            row++;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string path = "apns-conf-transsion.xml";
            loadFixTxt(path);

            StringBuilder sb = new StringBuilder();
            sb.AppendLine("<apns version=\"8\">");
            sb.AppendLine("<apns version=\"8\">");
            foreach (GPRSparam item in listGprs)
            {
                if (item.Mcc == "" || item.Mnc == "" || item.Name == "")
                {
                    continue;
                }
                sb.AppendLine("    <apn carrier=\"" + item.Name + "\"");
                sb.AppendLine("        mcc=\"" + item.Mcc + "\"");
                sb.AppendLine("        mnc=\"" + item.Mnc + "\"");
                sb.AppendLine("        apn=\"" + item.Apn + "\"");
                sb.AppendLine("        proxy=\"" + item.Proxy + "\"");
                sb.AppendLine("        port=\"" + item.Port + "\"");
                sb.AppendLine("        server=\"" + item.Server + "\"");
                sb.AppendLine("        user=\"" + item.Username + "\"");
                sb.AppendLine("        password=\"" + item.Password + "\"");
                sb.AppendLine("        type=\"" + item.Apntype + "\"");
                sb.AppendLine("        authtype=\"" + item.Authtype + "\"");
                sb.AppendLine("        preload=\"1\"");
                sb.AppendLine("    />");
            }
            sb.AppendLine("</apns>");
            System.IO.File.AppendAllText(path, sb.ToString());
        }

        private void loadFixTxt(string path)
        {
            StringBuilder cotent = new StringBuilder();
            cotent.Append(System.IO.File.ReadAllText("head.txt"));
            System.IO.File.AppendAllText(path, cotent.ToString());
        }


    }
}
