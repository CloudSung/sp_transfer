using System;
using System.Collections.Generic;
using System.Collections;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Reflection;
using System.IO;
using System.Runtime.InteropServices;
using System.Diagnostics;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Excel = Microsoft.Office.Interop.Excel;



namespace sp_transfer
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        public class FilePath
        {
            public static string OEMFilePath = "";
            public static string nightOwlFilePath = "";
            //Declare arraylistt to a globed arraylist.
            //Data stratuct is use Arraylist add a new list as a array object.
            public static ArrayList updateList = new ArrayList();
           
        }

        private void Button1_Click(object sender, EventArgs e)
        {
            string ASCode = GetASC("TSMC");
            Console.WriteLine(ASCode);

            if (FilePath.OEMFilePath == "" || FilePath.nightOwlFilePath == "")
            {
                MessageBox.Show("Please Choose file path", "Alert");
                return;
            } else {
                getExcelFile();
            }
        }
       

        public static void getExcelFile()
        {
            //Create COM Objects. Create a COM object for everything that is referenced
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(FilePath.OEMFilePath);
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;

            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;

            List<string> dateList = new List<string>();


            //iterate over the rows and columns and print to the console as it appears in the file
            //excel is not zero based!!

            bool isPrint = false;
            for (int i = 1; i <= rowCount; i++)
            {
                for (int j = 1; j <= colCount; j++)
                {
                    //new line
                    if (j == 1) {
                        //Console.Write("\r\n");
                        isPrint = false;
                    }

                    //write the value to the console
                    if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                    {
                        if (xlRange.Cells[i, j].Value2.ToString() == "Date")
                        {
                            isPrint = true;
                        }
                        if (isPrint)
                        {
                            if ((xlRange.Cells[i, j].Value2.ToString() != "Date"))
                            {
                                //Console.Write(xlRange.Cells[i, j].Value2.ToString() + "\t")
                                foreach (string myStringList in dateList)
                                {
                                    if (myStringList != xlRange.Cells[i, j].Value2.ToString())
                                    {
                                        dateList.Add(xlRange.Cells[i, j].Value2.ToString());
                                    }   
                                }
                            }  
                        }


                        //RootObject root = new RootObject();
                        //root.Date = xlRange.Cells[i, j].Value2.ToString();
                        //root.APN = "00001";
                        //root.AppleVendorCode = "111111M";
                        //root.Deamon = "12345";
                        ////Console.Write(xlRange.Cells[i, j].Value2.ToString() + "\t");

                        ////物件序列化
                        //string strJson = JsonConvert.SerializeObject(root, Formatting.Indented);
                        ////輸出結果
                        //Console.Write(strJson);

                    }


                    // Get Date of SP date to confirm is correct withlatest SP templelte or not.

                    //Start to collect value after Demand to rowCount.
                    //Key : supplierName/Date
                    //USse list as array.



                    //test to print out which value i need
                }
            }
            foreach (string myStringList in dateList)
            {
                Console.WriteLine(myStringList);
            }
            Console.Write("Loading Finish.");

            //cleanup
            GC.Collect();
            GC.WaitForPendingFinalizers();

            //rule of thumb for releasing com objects:
            //  never use two dots, all COM objects must be referenced and released individually
            //  ex: [somthing].[something].[something] is bad

            //release com objects to fully kill excel process from running in the background
            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(xlWorksheet);

            //close and release
            xlWorkbook.Close();
            Marshal.ReleaseComObject(xlWorkbook);

            //quit and release
            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);
            
        }


        private void SaveFile() {
            //Write part

            Excel.Application App = new Excel.Application();

            //取得欲寫入的檔案路徑
            Excel.Workbook Wbook = App.Workbooks.Open(FilePath.OEMFilePath);

            //將欲修改的檔案屬性設為非唯讀(Normal)，若寫入檔案為唯讀，則會無法寫入
            System.IO.FileInfo xlsAttribute = new FileInfo(FilePath.OEMFilePath);
            xlsAttribute.Attributes = FileAttributes.Normal;

            //取得batchItem的工作表
            Excel.Worksheet Wsheet = (Excel.Worksheet)Wbook.Sheets["SheetA"];

            //取得工作表的單元格
            //列(左至右)ABCDE、行(上至下)12345
            Excel.Range aRangeChange = Wsheet.get_Range("B1");

            //在工作表的特定儲存格，設定內容
            aRangeChange.Value2 = "加入訊息";

            //設置禁止彈出保存和覆蓋的詢問提示框
            Wsheet.Application.DisplayAlerts = false;
            Wsheet.Application.AlertBeforeOverwriting = false;
        }


        //Set Gatter AAnd Setter for ASC file
        public class RootObject
        {
            //ASC Transfer
            public string VendorName { get; set; }
            public string VendorNo { get; set; }
            public string AppleVendorCode { get; set; }
            public string UpdateTimestamp { get; set; }
        }

        public class transferToJson {

            //Transfer HHLH Excel Table to JSON, Setting Object Key.
            //Use JSON.net refer : https://dotblogs.com.tw/shadow/2012/08/16/74099
            public string SupplierName { get; set; }
            public string APN { get; set; }
            public string Date { get; set; }
            public string Deamon { get; set; }
            public string Site { get; set; }
        }

        //User Link and Lambda to search Json content.
        private string GetASC(string supplierName)
        {
            string strTsmc = supplierName;
            string ASCStr = System.IO.File.ReadAllText(@"../../ASC.txt");
            string LastValue = "";

            var MyClassList =JsonConvert.DeserializeObject<List<RootObject>>(ASCStr);

            var MyClass = MyClassList.Where(p => p.VendorName == strTsmc).FirstOrDefault();
            if (MyClass!=null)
            {
                LastValue = MyClass.AppleVendorCode;
            }
            return LastValue;

        }



        private void Button2_Click(object sender, EventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Multiselect = false;
            dialog.Title = "Choose Excel File...";
            dialog.Filter = "Excel File(*.xls, *.xlsx)|*.xls*";
            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                string file = dialog.FileName;
                Console.WriteLine(file);
                textBox1.Text = file;
            }
            FilePath.OEMFilePath = textBox1.Text;
        }

        private void Button3_Click(object sender, EventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Multiselect = false;
            dialog.Title = "Choose Excel File...";
            dialog.Filter = "Excel File(*.xls, *.xlsx)|*.xls*";
            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                string file = dialog.FileName;
                Console.WriteLine(file);
                textBox2.Text = file;
            }
            FilePath.nightOwlFilePath = textBox2.Text;
        }
    }
}
