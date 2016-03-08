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
using System.Text.RegularExpressions;
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

        string SheetCount = "";
        private void btnImportExcel_Click(object sender, EventArgs e)
        {
            if (txtAddress.Text != "" && System.IO.File.Exists(txtAddress.Text) == true)
            {
                SheetCount = txtSheetPage.Text;
                listGprs = new List<GPRSparam>();
                Thread th = new Thread(new ParameterizedThreadStart(importExcel));
                th.Start(txtAddress.Text);
                this.btnImportExcel.Enabled = false;
            }
            else 
            {
                MessageBox.Show("地址有误请重新选择!");
            }
        }


        private void importExcel(object obj)
        {
            string filename = (string)obj;
            EpplusExcel2007Read(filename);
            this.Invoke((EventHandler)delegate
            {
                this.btnImportExcel.Enabled = true;
                this.btnImportExcel.BackColor = Color.Green;
                this.btnOutPutExcel.Enabled = true;
            });
        }

        List<GPRSparam> listGprs = new List<GPRSparam>();

        string pretmpmms = "mms_";//短信前缀
        private void EpplusExcel2007Read(string path)
        {
            try
            {
                //实例化一个计时器
                Stopwatch watch = new Stopwatch();
                //开始计时/*此处为要计算的运行代码
                watch.Start();
              
                //文件信息
                FileInfo newFile = new FileInfo(path);
                using (ExcelPackage package = new ExcelPackage(newFile))
                {
                    string time = watch.ElapsedMilliseconds.ToString();
                    Console.WriteLine("加载完文件时间:" + time);
                    Lv.Log.Write("加载完文件时间: " + time, Lv.Log.MessageType.Info);
                    List<int> listCountPage = new List<int>();
                    SheetCount = SheetCount.Replace("，", ",");
                    string[] sheepsplit = SheetCount.Split(',');
                    foreach (string item in sheepsplit)
                    {
                        int temppage=0;
                        int.TryParse(item, out temppage);
                        listCountPage.Add(temppage);
                    }
                    if (listCountPage.Count == 0)
                    {
                        listCountPage.Add(2);
                    }
                    int vSheetCount = package.Workbook.Worksheets.Count; //获取总Sheet页

                    for (int pagei = 1; pagei <= vSheetCount; pagei++)
                    {
                        if (listCountPage.IndexOf(pagei) == -1)
                        {
                            continue;
                        }
                        ExcelWorksheet worksheet = package.Workbook.Worksheets[pagei];//选定 指定页
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
                        //保存表头信息
                        Dictionary<string, int> dictHeader = new Dictionary<string, int>();
                        List<string> listHeader = new List<string>();
                        int off = 0;    //起始行偏移量
                        bool isMMS = false;
                        //将每列标题添加到字典中
                        for (int i = colStart; i <= colEnd; i++)
                        {
                            if (worksheet.Cells[i,rowStart ].Value == null)
                            {
                                continue;
                            }
                            string titlestr = worksheet.Cells[ i , rowStart ].Value.ToString();
                            titlestr = titlestr.Trim();
                            if (titlestr == null || titlestr == "" )
                            {
                                continue;
                            }
                            if (titlestr == "MMS")
                            {
                                isMMS = true;
                            }
                            if (isMMS == false)
                            {
                                dictHeader.Add(titlestr, i);
                                listHeader.Add(titlestr);
                            }
                            else
                            {
                                dictHeader.Add(pretmpmms + titlestr, i);
                                listHeader.Add(pretmpmms + titlestr);
                            }
                        }
                        off += 1;
                        int count = 0;
                        if (System.IO.File.Exists("test.txt"))
                        {
                            System.IO.File.Delete("test.txt");    
                        }
                        
                        if (System.IO.File.Exists("mmstest.txt"))
                        {
                            System.IO.File.Delete("mmstest.txt");
                        }
                        
                        //针对国家里面合并列的情况需要临时变量保存
                        string country = "";
                        //针对印度版本里合并列的情况需要设置一个跳过列变量
                        int jumpcol = 0;
                        int mmsjumpcol = 0;
                        //遍历每一列
                        for (int col = colStart + off; col <= colEnd; col++)
                        {
                            count++;

                            //对印度版本的网络参数增加两个临时变量,重复用|号分隔
                            string repatemcc = "";
                            string repatemnc = "";
                            string mmsrepatemcc = "";
                            string mmsrepatemnc = "";

                            Dictionary<string, string> dictHeadervalue = new Dictionary<string, string>();
                            GPRSparam gprs = new GPRSparam();
                            GPRSparam mms = new GPRSparam();



                            if (jumpcol > 0)
                            {
                                col += jumpcol; //跳过合并列
                                count += jumpcol;
                                jumpcol = 0;
                                
                            }

                            
                            //遍历每一列的单元格
                            for (int row = rowStart; row <= rowEnd; row++)
                            {
                                
                                if (listHeader[row - 1] == pretmpmms + "MMS" && mmsjumpcol > 0) //为了把合并的MMS跳过
                                {
                                    break;
                                }

                                string text = "";
                                //得到单元格信息
                                ExcelRange cell = null;
                                try
                                {
                                    cell = worksheet.Cells[row, col];
                                    text = GetMegerValue(worksheet, row, col);


                                }
                                catch (Exception err)
                                {
                                    text = "";
                                    Console.WriteLine("" + err.Message);
                                    Lv.Log.Write("提取单元数据出错　row" + col.ToString() + " col" + row.ToString() + err.Message, Lv.Log.MessageType.Error);
                                }

                                if (listHeader[row - 1] == "MCC")
                                {
                                    //mcc
                                    //判断上一行是否合并,如果合并合并了多少行.
                                    string range = worksheet.MergedCells[row - 1, col];

                                    if (range != null && range.Contains(":"))    //说明合并了多列
                                    {
                                        jumpcol = GetMegerColSum(range);
                                    }
                                    if (jumpcol > 0)
                                    {
                                        text = GetMccMncValue(worksheet, row, col, text, jumpcol);
                                        //mcc
                                        repatemcc = text;
                                        text = "";
                                    }
                                    else
                                    {
                                        string tmpmncvalue = GetMegerValue(worksheet, row + 1, col); //如果mnc里面包含多个，那么这个时候需要记录mcc这样才能进行重复计算
                                        if (tmpmncvalue.Contains(",") == true)
                                        {
                                            repatemcc = text;
                                        }
                                        //其他情况不处理
                                    }
                                }
                                if (listHeader[row - 1] == "MNC")
                                {
                                    if (jumpcol > 0)
                                    {
                                        text = GetMccMncValue(worksheet, row, col, text, jumpcol);
                                        repatemnc = text;
                                        text = "";
                                    }
                                    else
                                    {
                                        if (text.Contains(",") == true)
                                        {
                                            repatemnc = text;
                                            text = "";
                                        }
                                        //其他情况不处理
                                    }
                                    
                                }
                                if (listHeader[row - 1] == pretmpmms + "MCC")
                                {
                                    //mcc
                                    //判断上一行是否合并,如果合并合并了多少行.
                                    string range = worksheet.MergedCells[row - 1, col];

                                    if (range != null && range.Contains(":"))    //说明合并了多列
                                    {
                                        mmsjumpcol = GetMegerColSum(range);
                                    }
                                    if (mmsjumpcol > 0)
                                    {
                                        text = GetMccMncValue(worksheet, row, col, text, jumpcol);
                                        mmsrepatemcc = text;
                                        text = "";
                                        if (mmsjumpcol == jumpcol)
                                        {
                                            mmsjumpcol = 0;
                                        }
                                        else
                                        {
                                            mmsjumpcol += 1;        //本次加1，为了第二列的时候能跳过
                                        }
                                    }
                                    else
                                    {
                                        string tmpmncvalue = GetMegerValue(worksheet, row + 1, col); //如果mnc里面包含多个，那么这个时候需要记录mcc这样才能进行重复计算
                                        if (tmpmncvalue.Contains(",") == true)
                                        {
                                            mmsrepatemcc = text;
                                        }
                                        //其他情况不处理
                                    }
                                }
                                if (listHeader[row - 1] == pretmpmms + "MNC")
                                {
                                    if (jumpcol > 0 || mmsjumpcol >0)
                                    {
                                        text = GetMccMncValue(worksheet, row, col, text, jumpcol);
                                        mmsrepatemnc = text;
                                        text = "";
                                    }
                                    else
                                    {
                                        if (text.Contains(",") == true)
                                        {
                                            mmsrepatemnc = text;
                                            text = "";
                                        }
                                        //其他情况不处理
                                    }
                                }
                                //end 多mms多mnc
                                // dictHeadervalue[col] = text;         //标题，值 不用标题了可以做判断
                                //对每一个网络参数进行修正操作
                                if (listHeader[row - 1].Contains("注：") == true)
                                {
                                    //说明到底了   1碰到一个问题，还真有人把这个注给删除了，导致程序跳过了原有的数据区域
                                    break;
                                }
                                text = revisedValue(text, listHeader[row - 1]);  //数据检查


                                if (listHeader.Count - 1 > row)
                                {
                                    dictHeadervalue.Add(listHeader[row - 1], text);
                                    Console.WriteLine(listHeader[row - 1] + "     " + row.ToString());
                                }
                                else
                                {
                                    break;
                                }
                            }
                            if (mmsjumpcol > 0)
                            {
                                mmsjumpcol--;
                            }
                            //前面对列的记录 同时还要记算出他合并了多少行，然后对其他行进行数据提取操作
                            //从目前就只有mcc和mnc有需要提取，只对这两个进行操作


                            foreach (var item in dictHeadervalue)
                            {
                                setValue(ref gprs, ref mms, item.Key, item.Value);
                            }
                            //设置proxy 为空用mmsproxy
                            if (gprs.Proxy ==  "0.0.0.0" && gprs.Mmsproxy != "0.0.0.0")
                            {
                                gprs.Proxy = gprs.Mmsproxy;
                                gprs.Port = gprs.Mmsport;
                            }
                            //设置MMSC 为空用homepage
                            if (gprs.Homepage == "" && gprs.Mmsc != "")
                            {
                                gprs.Homepage = gprs.Mmsc;
                            }
                            //印度多mcc mnc进入
                            if (repatemcc != "" || repatemnc != "")
                            {
                                //SaveRecode(repatemcc, repatemnc, mmsrepatemcc, mmsrepatemnc, gprs, mms);
                                SaveRecodeGprs(repatemcc, repatemnc, gprs);
                            }
                            else
                            {
                               GprsAdd(gprs);
                            }
                            if ( mmsrepatemcc != "" || mmsrepatemnc != "")
                            {
                                SaveRecodeGprs(mmsrepatemcc, mmsrepatemnc, mms);
                            }
                            else
                            {
                                //保存 当前列的gprs mms
                                //ConvertGprs(dictHeadervalue,dictHeader);
                              
                                GprsAdd(mms);
                            }
                        }
                        Console.WriteLine("总处理列数:" + count);
                        Lv.Log.Write("总处理行数: " + count, Lv.Log.MessageType.Info);
                    }
                }
            }
            catch (Exception err)
            {
                Console.WriteLine("加载excel出错了　" + err.Message);
                MessageBox.Show("加载excel出错了　" + err.Message+" 如找不到原因，可致邮452113521@qq.com");
                Lv.Log.Write("加载excel出错了　 " + err.Message, Lv.Log.MessageType.Error);
                return;
            }

            MessageBox.Show("已正常打开网络参数表");
        }
        /// <summary>
        /// 对保存的GPRS进行判断，然后保存
        /// </summary>
        /// <param name="gprs"></param>
        private void GprsAdd(GPRSparam gprs)
        {
            //throw new NotImplementedException();
            if (gprs.Mcc != "" && gprs.Mnc != "")
            {
                listGprs.Add(gprs);   
                
                //对比检查数据
                if (gprs.Type == GPRSTYPE.GPRS)
                {
                    string content = gprs.Mnc + "\r\n";
                    System.IO.File.AppendAllText("test.txt", content);
                }

                //对比检查数据
                if (gprs.Type == GPRSTYPE.MMS)
                {
                    string content = gprs.Mnc + "\r\n";
                    System.IO.File.AppendAllText("mmstest.txt", content);
                }

            }
            
            
        }

        private string GetMccMncValue(ExcelWorksheet worksheet, int row, int col, string text,int jumpcol)
        {
            //throw new NotImplementedException();
            string tmprange = worksheet.MergedCells[row, col]; // 用来判断mcc是否也进行了合并
            if (tmprange != null && tmprange.Contains(":") == true)
            {
                List<string> listtext = new List<string>();//记录上一个列表里的mcc 
                listtext.Add(text);
                for (int i = 1; i < jumpcol + 1; i++)
                {
                    string tmptext = GetMegerValue(worksheet, row, col + i);
                    if (tmptext != listtext[i - 1])
                    {
                        text += "|" + tmptext;
                    }
                }
            }
            else
            {
                //没有合并的情况下
                for (int i = 1; i < jumpcol + 1; i++)
                {
                    text += "|" + GetMegerValue(worksheet, row, col + i);
                }
            }
            return text;
        }

        private void SaveRecodeGprs(string repatemcc, string repatemnc, GPRSparam gprs)
        {
            //throw new NotImplementedException();
            string[] mccsplit = repatemcc.Split('|');
            string[] mncsplit = repatemnc.Split('|');
            for (int i = 0; i < mccsplit.Length; i++)
            {
                string[] mncsig = mncsplit[i].Split(',');
                for (int j = 0; j < mncsig.Length; j++)
                {
                    GPRSparam tmpgprs = (GPRSparam)gprs.Clone();
                    tmpgprs.Mcc = mccsplit[i];
                    tmpgprs.Mnc = mncsig[j];
                    //listGprs.Add(tmpgprs);
                    GprsAdd(tmpgprs);
                }
            }
        }

        //计算出合并了多行列
        private int GetMegerColSum(string range)
        {
            //方法1，但是这种方法不精准
/*
            //throw new NotImplementedException(); 返回异常
            string[] strsplit = range.Split(':');
            byte[] arry1 =  System.Text.ASCIIEncoding.ASCII.GetBytes(strsplit[0]);
            byte[] arry2 = System.Text.ASCIIEncoding.ASCII.GetBytes(strsplit[1]);

            int jump = sumArray(arry2) - sumArray(arry1);
            return jump;
 * 
 * 
*/
            //方法二 计算出行的值，然后进行减运算就得到结果了
            string[] strsplit = range.Split(':');
            string col1 = System.Text.RegularExpressions.Regex.Match(strsplit[0], "(?<col>[A-Z]*)").Result("$1").ToString();
            int col1hao = convertR1c1(col1);
            string col2 = System.Text.RegularExpressions.Regex.Match(strsplit[1], "(?<col>[A-Z]*)").Result("$1").ToString();
            int col2hao = convertR1c1(col2);
            return System.Math.Abs( col2hao - col1hao);
        }

        /// <summary>
        /// 把ABC转换成行号
        /// </summary>
        /// <param name="col1"></param>
        /// <returns></returns>
        private int convertR1c1(string col1)
        {
            int sum = 0;
            //throw new NotImplementedException();
            byte[] arry1 = System.Text.ASCIIEncoding.ASCII.GetBytes(col1);
            int local = 0;
            for (int i = arry1.Length-1; i >= 0; i--)
            {
                if (local == 0)
                {
                    sum += (arry1[i] - 64);
                }
                else
                {
                    double tmppow = System.Math.Pow(26, local);
                    sum += (int)(tmppow * (double)(arry1[i] - 64));
                }
                local++;
            }
            return sum;
        }

        private int sumArray(byte[] arry1)
        {
            int sum = 0;
            for (int i = 0; i < arry1.Length; i++)
            {
                sum += arry1[i];
            }
            return sum;
        }

        private void setValue(ref GPRSparam gprs, ref GPRSparam mms, string title, string text)
        {
            switch (title)
            {
                case "Tecno / itel":
                    gprs.Country = text;
                    mms.Country = text;
                    break;
                case "GPRS/ EDGE/Internet": 
                    gprs.Type = GPRSTYPE.GPRS;
                    gprs.Typestr = "gprs";
                    break;
                case "mvno_type":
                    gprs.Mvno_type = text;
                    mms.Mvno_type = text;
                    break;
                case "mvno_match_data":
                    gprs.Mvno_match_data = text;
                    mms.Mvno_match_data = text;
                    break;
                case "insert SIM  idle display":
                    gprs.Idledisplay = text;
                    mms.Idledisplay = text;
                    break;
                case "Operator Name":
                    gprs.OperatorName = text;
                    mms.OperatorName = text;
                    break;
                case "NAME":
                    gprs.Name = text;
                    break;
                case "Homepage":
                    gprs.Homepage = text;    
                    break;
                case "APN":
                    gprs.Apn = text;
                    break;
                case "Proxy Enable":
                    gprs.ProxyEnable = text;
                    break;
                case "PROXY":
                    if (text == "")
                    {
                        text = "0.0.0.0";
                    }
                    gprs.Proxy = text;
                    break;
                case "PORT":
                    text = CheckPort(text);
                    gprs.Port = text;
                    break;
                case "USERNAME":
                    gprs.Username = text;
                    break;
                case "PASSWORD":
                    gprs.Password = text;
                    break;
                case "SERVER":
                    gprs.Server = text;
                    break;
                case "MMSC":
                    gprs.Mmsc = text;
                    break;
                case "MMSPROXY":
                    if (text == "")
                    {
                        text = "0.0.0.0";
                    }
                    gprs.Mmsproxy = text;
                    break;
                case "MMS PORT":
                    text = CheckPort(text);
                    gprs.Mmsport = text;
                    break;
                case "MCC":
                    gprs.Mcc = text;
                    break;
                case "MNC":
                    if (text.Length < 2)
                    {
                        text = "0" + text;
                    }
                    gprs.Mnc = text;
                    break;
                case "AUTHENTICATION TYPE":
                    text = CheckAuthType(text);
                    gprs.Authtype = text;
                    break;
                case "APN TYPE":
                    gprs.Apntype = text;
                    break;
                case "mms_MMS":
                    mms.Type = GPRSTYPE.MMS;
                    mms.Typestr = "mms";
                    break;
                case "mms_Operator Name":
                    mms.OperatorName = text;
                    break;
                case "mms_NAME":
                    mms.Name = text;
                    break;
                case "mms_APN":
                    mms.Apn = text;
                    break;
                case "mms_Proxy Enable":
                    mms.ProxyEnable = text;
                    break;
                case "mms_PROXY":
                    if (text == "")
                    {
                        text = "0.0.0.0";
                    }
                    mms.Proxy = text;
                    break;
                case "mms_PORT":
                    text = CheckPort(text);
                    mms.Port = text;
                    break;
                case "mms_USERNAME":
                    mms.Username = text;
                    break;
                case "mms_PASSWORD":
                    mms.Password = text;
                    break;
                case "mms_SERVER":
                    mms.Server = text;
                    break;
                case "mms_MMSC":
                    mms.Mmsc = text;
                    break;
                case "mms_MMS PROXY":
                    if (text == "")
                    {
                        text = "0.0.0.0";
                    }
                    mms.Mmsproxy = text;
                    break;
                case "mms_MMS PORT":
                    text = CheckPort(text);
                    mms.Mmsport = text;
                    break;
                case "mms_MCC":
                    mms.Mcc = text;
                    break;
                case "mms_MNC":
                    if (text.Length < 2)
                    {
                        text = "0" + text;
                    }
                    mms.Mnc = text;
                    break;
                case "mms_AUTHENTICATION TYPE":
                    mms.Authtype = text;
                    break;
                case "mms_APN TYPE":
                    mms.Apntype = text;
                    break;

                default:
                    break;
            }
        }

        private string CheckAuthType(string text)
        {
            //throw new NotImplementedException();
            if (text == "Normal")   //讨论最后结果，是随便选哪种都行
            {
                text = "PAP";
            }
            else if (text == "Secured") //讨论最后结果，是随便选哪种都行
            {
                text = "PAP";
            }
            else if (text.Contains("PAP"))
            {
                text = "PAP";
            }
            else if (text =="CHAP")
            {
                text = "CHAP";
            }
            else if (text.Contains("Not"))
            {
                text = "PAP";
            }
            else if (text.Contains("None"))
            {
                text = "PAP";
            }
            else
            {
                text = "PAP";
            }
            return text;
        }

        private string CheckPort(string text)
        {
            if (text.Contains("or"))
            {
                text = System.Text.RegularExpressions.Regex.Match(text, "(?<col>[0-9]*)").Result("$1").ToString();
            }
            if (text == "")
            {
                text = "0";
            }
            return text;
        }
        public static string GetMegerValue(ExcelWorksheet wSheet, int row, int column)
        {
            string range = wSheet.MergedCells[row, column];
            if (range == null)
                if (wSheet.Cells[row, column].Value != null)
                    return wSheet.Cells[row, column].Value.ToString();
                else
                    return "";
            object value =
                wSheet.Cells[(new ExcelAddress(range)).Start.Row, (new ExcelAddress(range)).Start.Column].Value;
            if (value != null)
                return value.ToString();
            else
                return "";
        }
        private void ConvertGprs(Dictionary<int, string> dictHeadervalue, Dictionary<string, int> dictHeader)
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


        private void ConvertAPNGprs(Dictionary<int, string> dictHeadervalue, Dictionary<int, string> dictHeader)
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
            mms.Mcc = GetValue(dictHeadervalue, 31);
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
        /// 修正值 对各种数据的验证，如果不对提示异常退出
        /// </summary>
        /// <param name="text"></param>
        /// <param name="p"></param>
        /// <returns></returns>
        private string revisedValue(string text, string title)
        {

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
                        if (item.Mcc == "" || item.Mnc == "" )
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
                        string text = OutExcelSetProxyValue(item);
                        itemValues.Add(text);        //短信的和gprs的不一样
                        text = OutExcelSetPortValue(item);
                        itemValues.Add(text);         //短信的和gprs的不一样
                        itemValues.Add(item.Username);
                        itemValues.Add(item.Password);
                        itemValues.Add("0.0.0.0");
                        itemValues.Add("0.0.0.0");
                        text = OutExcelSetHomePageValue(item,url);
                        itemValues.Add(text);
                        text = OutExcelSetAuthTypeValue(item);
                        itemValues.Add(text);    //#PAP 0 CHAP 1  这个更据表上来
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

        private string OutExcelSetAuthTypeValue(GPRSparam item)
        {
            //throw new NotImplementedException();
            //#PAP 0 CHAP 1  这个更据表上来
            string text = "0";
            switch (item.Authtype)
            {
                case "CHAP":
                    text = "1";
                    break;
                default:
                    break;
            }
            return text;
        }

        private string OutExcelSetHomePageValue(GPRSparam item,string url)
        {
            //throw new NotImplementedException();
            string text = "";
            if (item.Type == GPRSTYPE.GPRS)
            {
                if (item.Homepage != "")
                {
                    text = item.Homepage;
                }
                else if (item.Server != "")
                {
                    text = item.Server;      //homepage
                }
                else
                {
                    text = url;
                }
            }
            else if (item.Type == GPRSTYPE.MMS)
            {
                if (item.Mmsc != "")
                {
                    text = item.Mmsc;      //homepage
                }
                else
                {
                     text = url;
                }
            }
            return text;
        }

        private string OutExcelSetPortValue(GPRSparam item)
        {
            string text = "";
            if (item.Type == GPRSTYPE.GPRS)
            {
                text = item.Port;
                if (text == "")
                {
                    text = item.Mmsport;
                }
            }
            else if (item.Type == GPRSTYPE.MMS)
            {
                text = item.Mmsport;
                if (text == "")
                {
                    text = item.Port;
                }
            }
            else
            {

            }
            return text;
        }

        private string OutExcelSetProxyValue(GPRSparam item)
        {
            string text = "";
            //throw new NotImplementedException();
            if (item.Type == GPRSTYPE.GPRS)
            {
                text = item.Proxy;
                if (text == "")
                {
                    text = item.Mmsproxy;
                }
            }
            else if (item.Type == GPRSTYPE.MMS)
            {
                text = item.Mmsproxy; 
                if (text == "")
                {
                    text = item.Proxy;
                }
            }
            else
            { 

            }
            //itemValues.Add(item.Type == GPRSTYPE.GPRS ? item.Proxy : item.Mmsproxy);        //短信的和gprs的不一样
            return text;
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
            string path = @"G:\GPRSTool\GPRSTOOL\智能机网络参数文件\apns-conf-transsion.xml";
            if (System.IO.File.Exists(path) == false)
            {
                return;
            }
            int count = 0;
           
            foreach (GPRSparam item in listGprs)
            {
                StringBuilder sb = new StringBuilder();
                GPRSparam tmpitem = item;
                if (tmpitem.Mcc == "" || tmpitem.Mnc == "" || tmpitem.Name == "")
                {
                    continue;
                }
               // if (checkApnMncMccSmart(path, tmpitem) == false)
                {
                 //   continue;
                }
                count++;
                checkApnName(ref tmpitem);
                
                sb.AppendLine("    <apn carrier=\"" + tmpitem.Name + "\"");
                sb.AppendLine("        mcc=\"" + tmpitem.Mcc + "\"");
                sb.AppendLine("        mnc=\"" + tmpitem.Mnc + "\"");

                sb.AppendLine("        apn=\"" + tmpitem.Apn + "\"");
                sb.AppendLine("        proxy=\"" + tmpitem.Proxy + "\"");
                sb.AppendLine("        port=\"" + tmpitem.Port + "\"");
                sb.AppendLine("        server=\"" + tmpitem.Server + "\"");
                sb.AppendLine("        user=\"" + tmpitem.Username + "\"");
                sb.AppendLine("        password=\"" + tmpitem.Password + "\"");
                sb.AppendLine("        type=\"" + tmpitem.Apntype + "\"");
                sb.AppendLine("        authtype=\"" + tmpitem.Authtype + "\"");
                sb.AppendLine("        preload=\"1\"");
                sb.AppendLine("    />");

                if (sb.ToString() != "")
                {
                    LoadApnXML(path);
                    string NewApnContent = ApnContentXML;
                    contentMostPrimitive = ApnContentXML;
                    string newPath = "apns-conf-transsionNEW.xml";
                    //这里判断是全新追加，还是更新
                    if (IsCheckApdOrMod( tmpitem, newPath))
                    {
                       


                        ////append
                        //sb.AppendLine("</apns>");
                        ////替换原始数据到另一个文件中
                        //NewApnContent = NewApnContent.Replace("</apns>", sb.ToString());
                        //System.IO.File.AppendAllText(newPath, NewApnContent);
                    }
                    else
                    {
                        //modifly

                    }

                }

            }
            System.IO.File.WriteAllText(path.Replace(".xml","1.xml"), contentMostPrimitive);
            MessageBox.Show("程序导出结束 共导入" + count.ToString() + "条");
           
        }
        string contentMostPrimitive = "";
        private bool IsCheckApdOrMod( GPRSparam tmpitem, string newPath)
        {

            //throw new NotImplementedException();
            string parttn =string.Format("carrier=\"%0\"[\\s]*?mcc=\"%1\"[\\s]*?mnc=\"%2\"",tmpitem.Name,tmpitem.Mcc,tmpitem.Mnc);
            Match ma = Regex.Match(contentMostPrimitive, parttn, RegexOptions.IgnoreCase);
            if (ma.Success)
            {
                /**
                * 临时创建初始文件用的方法
                * **/
                parttn = "<apn carrier=\"[\\s\\S]*?\"[\\s]*?mcc=\"[\\s\\S]*?\"[\\s]*?mnc=\"[\\s\\S]*?\"[^>]*?>";
                Match mapn = Regex.Match(contentMostPrimitive, parttn, RegexOptions.IgnoreCase);
                string value = mapn.Result("$1");
                contentMostPrimitive = contentMostPrimitive.Replace(value, "");
                return true;
            }
            else
            {
                return false;
            }
            
        }



        string SpnConf = "";
        private void checkApnName(ref GPRSparam item)
        {
            //throw new NotImplementedException();
            if (System.IO.File.Exists("spn-conf.xml") == false)
            {
                MessageBox.Show("程序文件spn-conf.xml不全，无法正常操作。");
                return;
            }
            if (SpnConf == "" )
            {
                SpnConf = System.IO.File.ReadAllText("spn-conf.xml");
            }
            string tmpmccmnc = "\"" + item.Mcc + item.Mnc + "\"";
            System.Text.RegularExpressions.Match mc = System.Text.RegularExpressions.Regex.Match(SpnConf, "numeric="+tmpmccmnc+" spn=\"(?<spn>[\\s\\S]*?)\"/>");
            string strName = mc.Groups["spn"].Value;
            if (strName != "")
            {
                item.Name = strName;
            }

        }
        /// <summary>
        /// xml文件内容
        /// </summary>
        string ApnContentXML = "";
        private bool LoadApnXML(string path)
        {
            if (ApnContentXML == "")
            {
                ApnContentXML = System.IO.File.ReadAllText(path);
                if (ApnContentXML.Contains("/apns>") == false)
                {

                    if (System.IO.File.Exists("head.txt") == true)
                    {
                        MessageBox.Show("文件内容有误，不是apn指定的文件！将会以head.txt文件为打开文件");
                        ApnContentXML = System.IO.File.ReadAllText("head.txt");
                    }
                    else
                    {
                        MessageBox.Show("非指定文件，程序文件head.txt不全，无法正常操作。");
                        System.Environment.Exit(0);
                    }
                    
                }
                return true;
            }
            return false;
        }
        private bool checkApnMncMccSmart(string path,GPRSparam gprs )
        {
            LoadApnXML(path);
            if (ApnContentXML.Contains("mcc=\"" + gprs.Mcc + "\"") == true && ApnContentXML.Contains("mnc=\"" + gprs.Mnc + "\"") == true)
            {
                return false;
            }

            return true;
        }

        private void btnAPNImport_Click(object sender, EventArgs e)
        {

        }

        private void importAPNExcel(object obj)
        {
            string filename = (string)obj;
            EpplusAPNExcel2007Read(filename);
        }

        private void EpplusAPNExcel2007Read(string filename)
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
                FileInfo newFile = new FileInfo(filename);
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
                        for (int i = rowStart; i <= rowEnd; i++)
                        {
                            if (worksheet.Cells[i, colStart].Value == null)
                            {
                                continue;
                            }
                            string titlestr = worksheet.Cells[i, colStart].Value.ToString();
                            if (titlestr == null || titlestr == "")
                            {
                                continue;
                            }
                            dictHeader[i] = titlestr.Replace(" ", "");
                        }
                        off += 1;
                        int count = 0;
                        //遍历每一列
                        for (int row111 = rowStart + off; row111 <= rowEnd; row111++)
                        {
                            Dictionary<int, string> dictHeadervalue = new Dictionary<int, string>();
                            count++;
                            //遍历每一列的单元格
                            for (int col = colStart; col <= colEnd; col++)
                            {
                                string text = "";
                                //得到单元格信息
                                ExcelRange cell = null;
                                try
                                {
                                    cell = worksheet.Cells[col, row111];
                                }
                                catch (Exception err)
                                {
                                    text = "";
                                    Console.WriteLine("" + err.Message);
                                    Lv.Log.Write("提取单元数据出错　row" + row111.ToString() + " col" + col.ToString() + err.Message, Lv.Log.MessageType.Error);
                                }
                                if (cell.Value == null)
                                {
                                    text = "";
                                }
                                else
                                {
                                    text = cell.RichText.Text;
                                }
                                // dictHeadervalue[col] = text;         //标题，值 不用标题了可以做判断
                                //对每一个网络参数进行修正操作
                                if (col >= 34)
                                {
                                    break;
                                }
                                text = revisedValue(text, dictHeader[col]);
                                dictHeadervalue.Add(col, text);
                            }

                            //保存 当前列的gprs mms
                            ConvertAPNGprs(dictHeadervalue, dictHeader);
                        }
                        Console.WriteLine("总处理列数:" + count);
                        Lv.Log.Write("总处理行数: " + count, Lv.Log.MessageType.Info);
                    }
                }
            }
            catch (Exception err)
            {
                Console.WriteLine("加载excel出错了　" + err.Message);
                MessageBox.Show("加载excel出错了　" + err.Message + " 如找不到原因，可致邮452113521@qq.com");
                Lv.Log.Write("加载excel出错了　 " + err.Message, Lv.Log.MessageType.Error);
                return;
            }

            MessageBox.Show("已正常打开网络参数表");
        }

        private void btnImportExcelPath_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "*.*|*.*";
            if (ofd.ShowDialog() == DialogResult.OK)
            {

                txtAddress.Text = ofd.FileName;
            }
        }


    }
}
