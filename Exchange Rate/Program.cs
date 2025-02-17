﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static System.Net.Mime.MediaTypeNames;
using Excel = Microsoft.Office.Interop.Excel;
using System.Diagnostics;
using Microsoft.Office.Interop.Excel;
using System.Data.SqlClient;
using System.Net.Mail;




namespace Exchange_Rate
{
    internal class Program
    {
        static String strSQLConnection = "Data Source =192.168.10.22; Initial Catalog = Price; Persist Security Info=false; User ID = sa; Password = yzf; Max Pool Size=30000;Connection Timeout=1200";//資料庫連接正式區
        //static String strSQLConnection = "Data Source =192.168.10.22; Initial Catalog = Test; Persist Security Info=false; User ID = sa; Password = yzf; Max Pool Size=30000;Connection Timeout=1200";//資料庫連接測試區區
        static void Main(string[] args)
        {
            try
            {
                string fileName = @"\\dcc\二岸共用資料\ID\ID匯率.XLS";
                object missing = System.Reflection.Missing.Value;
                Excel.Application excel = new Excel.Application();//lauch excel application
                if (excel == null)
                {
                    return;
                }
                else
                {
                    excel.Visible = false; excel.UserControl = true;
                    //以只讀方式打開EXCEL文件
                    Workbook wb = excel.Application.Workbooks.Open(fileName, missing, true, missing, missing, missing,
                     missing, missing, missing, true, missing, missing, missing, missing, missing);
                    //取得工作薄
                    Worksheet ws = (Worksheet)wb.Worksheets.get_Item("公司匯率");

                    //取得資料行數
                    int rowsint = ws.UsedRange.Cells.Rows.Count;

                    //取得資料範圍
                    Range range = ws.Cells.get_Range("B2", "B11");

                    //讀取資料
                    object[,] arryItem = (object[,])range.Value2;

                    string HK = arryItem[1, 1].ToString();
                    Do_cur_update("HK$", HK);

                    string HKD = arryItem[2, 1].ToString();
                    Do_cur_update("港幣", HKD);

                    string RM = arryItem[3, 1].ToString();
                    Do_cur_update("RMB", RM);

                    string RMB = arryItem[4, 1].ToString();
                    Do_cur_update("人民幣", RMB);

                    string US = arryItem[5, 1].ToString();
                    Do_cur_update("US$", US);

                    string USD = arryItem[6, 1].ToString();
                    Do_cur_update("美金", USD);

                    string VND = arryItem[7, 1].ToString();
                    Do_cur_update("越南盾", VND);

                    string VNDUS = arryItem[8, 1].ToString();
                    Do_cur_update("越南盾(US$)", VNDUS);

                    string VNDNT = arryItem[9, 1].ToString();
                    Do_cur_update("越南盾(NT$)", VNDNT);

                    string EUR = arryItem[10, 1].ToString();
                    Do_cur_update("歐元", EUR);

                }
                excel.Quit();
                excel = null;
                Process[] procs = Process.GetProcessesByName("excel");


                foreach (Process pro in procs)
                {
                    pro.Kill();//没有更好的方法,只有杀掉进程
                }
                GC.Collect();
                string strResult = "匯率自動輸入成功 ";
                Mail(strResult);
            }
            catch (Exception ex)
            {
                string strResult = "匯率自動輸入失敗 " + ex;
                Mail(strResult);
            }
        }

        private static void Do_cur_update(string codename, string convert)     //更新cur資料
        {
            SqlConnection conn = new SqlConnection(strSQLConnection);
            conn.Open(); //開啟資料庫
            SqlCommand cmd;
            string strSQL = $@"exec cur_update '{codename.Trim()}',{convert} ,'系統' ";
            cmd = new SqlCommand(strSQL, conn);
            //cmd.ResetCommandTimeout();
            //CommandTimeout 重設為30秒
            //怕下列指令執行較長,將他延長設為1200秒
            cmd.CommandTimeout = 800000;
            cmd.ExecuteNonQuery();
            conn.Close(); //關閉資料庫連接
        }

        private static void Mail(string strResult)   //發送MAIL
        {
            MailMessage MyMail = new MailMessage();
            MyMail.From = new MailAddress("sqluser@msl.com.tw");
            //MyMail.To.Add("收件者Email");加入收件者Email
            MyMail.To.Add("peggy@msl.com.tw"); //加入收件者Email
            //MyMail.CC.Add("副本的Mail"); //加入副本的Mail
            //MyMail.Bcc.Add("密件副本的收件者Mail"); //加入密件副本的Mail          
            MyMail.Subject = "匯率自動輸入";
            MyMail.Body = strResult + " 結束時間:" + DateTime.Now.ToString(); //設定信件內容
            MyMail.IsBodyHtml = false; //是否使用html格式
            //Attachment attdata = new Attachment(@"D:\" + sDate + @"生產日報.xlsx", MediaTypeNames.Application.Octet);
            //MyMail.Attachments.Add(attdata);
            SmtpClient MySMTP = new SmtpClient("webmail.msl.com.tw", 25);
            MySMTP.Credentials = new System.Net.NetworkCredential("sqluser@msl.com.tw", "msl22995234");
            MySMTP.Send(MyMail);
            MyMail.Dispose(); //釋放資源
        }
    }
}
