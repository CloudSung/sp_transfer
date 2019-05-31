using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;
using System.IO;
using System.Runtime.InteropServices;
using System.Diagnostics;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;



namespace sp_transfer
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

 

        private void Button1_Click(object sender, EventArgs e)
        {
            //get file
            //getExcelFile();

            Debug.WriteLine("Test2321213124");
            GetASC();
        }

        public static void getExcelFile()
        {

            //Create COM Objects. Create a COM object for everything that is referenced
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(@"C:\APO\HHCD_BLDD2\Forecast.xlsx");
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;

            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;


            Console.Write(@"test");
            //iterate over the rows and columns and print to the console as it appears in the file
            //excel is not zero based!!
            for (int i = 1; i <= rowCount; i++)
            {
                for (int j = 1; j <= colCount; j++)
                {
                    //new line
                    if (j == 1)
                        Console.Write("\r\n");

                    //write the value to the console
                    if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                        Console.Write(xlRange.Cells[i, j].Value2.ToString() + "\t");
                }
            }

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
            string strPath = "C:\\APO\\HHCD_BLDD2\\Forecasr.xlsx";
            Excel.Workbook Wbook = App.Workbooks.Open(strPath);

            //將欲修改的檔案屬性設為非唯讀(Normal)，若寫入檔案為唯讀，則會無法寫入
            System.IO.FileInfo xlsAttribute = new FileInfo(strPath);
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

        public class RootObject
        {
            public string VendorName { get; set; }
            public object VendorNo { get; set; }
            public string AppleVendorCode { get; set; }
            public string UpdateTimestamp { get; set; }
        }


        private void GetASC(){
            string ASCStr = System.IO.File.ReadAllText(@"../../ASC.txt");
            //Console.WriteLine(ASCStr);

            //var result = JsonConvert.DeserializeObject<List<RootObject>>(ASCStr);            
            //Console.WriteLine(result);

            //JArray ASC_Ary = JArray.Parse(ASCStr);
            //Console.WriteLine(ASC_Ary.ToString());
            JObject ASC_Dict = JObject.Parse(ASCStr);
            Console.WriteLine(ASC_Dict.ToString());
        }
    }
}
