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


namespace 口算题卡
{
    public partial class Form1 : Form
    {
        Excel.Application excelApp;
        ExcelClass myExcel;
        public Form1()
        {
            InitializeComponent();
            excelApp = new Excel.Application();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            Console.WriteLine(Application.StartupPath);
            string strResult = " string.Empty";
            //writeExcel(strResult, "d:\\口算题卡.xlsx");
            Console.WriteLine(strResult);
            myExcel = new ExcelClass();
            myExcel.Open("d:\\口算题卡.xlsx");

        }

        private void kousuan()
        {
            Random random = new Random();
            int intLineTemp = 0;    //每页有8行，到8行也就是一页后，行数+11，开始下一页。
            bool bolPageStart = true;


            for (int i = 2; i < 1052; i += 6)//行
            {
                if (bolPageStart)
                {
                    bolPageStart = false;
                    myExcel.Write(i - 1, 1, "日期")  ;
                    myExcel.Write(i - 1, 4, "用时");
                    myExcel.Write(i - 1, 7, "错误");
                }
                for (int j = 1; j < 9; j += 2) //列

                {
                    string strTemp = string.Empty;
                    int int1, int2, intResult;
                    int1 = random.Next(11, 99);
                    int2 = random.Next(11, 99);
                    intResult = int1 * int2;
                    strTemp = string.Format("{0} * {1} = {2}", int1, int2, intResult);
                    myExcel.Write(i, j, strTemp);
                    Console.WriteLine(string.Format("行{0}，列{1}，值{2}", i, j, strTemp));
                }
                intLineTemp++;
                if (intLineTemp>=8)
               {
                    intLineTemp = 0;
                    bolPageStart = true;
                    i+=5;
                }
            }
            myExcel.Close();
        }

        /// <summary>
        /// 打开excel文件
        /// </summary>
        /// <param name="path"></param>
        private void OpenExcel(string path)
        {
            //1.创建Applicaton对象
            Microsoft.Office.Interop.Excel.Application xApp = new

            Microsoft.Office.Interop.Excel.Application();

            //2.得到workbook对象，打开已有的文件
            Excel.Workbook xBook = xApp.Workbooks.Open(path);
        }

        /// <summary>
        /// 关闭并保存Excel
        /// </summary>
        /// <param name="xBook"></param>
        private void CloseExcel(Excel.Application xApp, Excel.Workbook xBook, Excel.Worksheet xSheet)
        {
            //5.保存保存WorkBook
            xBook.Save();
            //6.从内存中关闭Excel对象

            xSheet = null;
            xBook.Close();
            xBook = null;
            //关闭EXCEL的提示框
            xApp.DisplayAlerts = false;
            //Excel从内存中退出
            xApp.Quit();
            xApp = null;
        }

        //将数据写入已存在Excel
        public static void writeExcel(string result, string filepath)
        {
            Random random = new Random();
            //1.创建Applicaton对象
            Microsoft.Office.Interop.Excel.Application xApp = new

            Microsoft.Office.Interop.Excel.Application();

            //2.得到workbook对象，打开已有的文件
            Excel.Workbook xBook = xApp.Workbooks.Open(filepath);


            //3.指定要操作的Sheet
            Microsoft.Office.Interop.Excel.Worksheet xSheet = (Microsoft.Office.Interop.Excel.Worksheet)xBook.Sheets[1];

            //在第一列的左边插入一列  1:第一列
            //xlShiftToRight:向右移动单元格   xlShiftDown:向下移动单元格
            //Range Columns = (Range)xSheet.Columns[1, System.Type.Missing];
            //Columns.Insert(XlInsertShiftDirection.xlShiftToRight, Type.Missing);
            for (int j = 1; j < 45; j += 6)//行            
            {
                for (int i = 0; i < 7; i++) //列
                {
                    string strTemp = string.Empty;
                    int int1, int2, intResult;
                    int1 = random.Next(11, 99);
                    int2 = random.Next(11, 99);
                    intResult = int1 * int2;
                    strTemp = string.Format("{0} * {1} = {2}", int1, int2, intResult);

                    xSheet.Cells[i + 1][j + 1] = strTemp;
                }
            }
            //4.向相应对位置写入相应的数据
            //xSheet.Cells[2,1] = result;

            //5.保存保存WorkBook
            xBook.Save();
            //xBook .print
            //6.从内存中关闭Excel对象

            xSheet = null;
            xBook.Close();
            xBook = null;
            //关闭EXCEL的提示框
            xApp.DisplayAlerts = false;
            //Excel从内存中退出
            xApp.Quit();
            xApp = null;

        }

        private void button1_Click(object sender, EventArgs e)
        {
            kousuan ();
        }

        private void button2_Click(object sender, EventArgs e)
        {

            for (int j = 2; j < 45; j = j + 6)
            {
                for (int i = 1; i < 8; i = i + 2)
                {
                    myExcel.Write(j, i, "嘟嘟" + i.ToString());
                }
            }


            myExcel.Close();

            MessageBox.Show("Clicked Test Button.");
        }
        private void letter(int number)

        {
            Console.Write(number.ToString());

        }
        private void hh()
        {
            Console.WriteLine();
        }


        private void CalZMLM()
        {
            float houdu = 0.02f;

            

            for (int i = 1; i < 10000; i++)
            {
                houdu = houdu * 2;
                string strTemp = "折叠" + i.ToString() + " 次，厚度是" + (houdu/1000).ToString() +"M";
                Console.WriteLine(strTemp);
                if (houdu/1000 >8848)
                {
                    Console.WriteLine("OK" + i.ToString());
                    break;
                }
            }
        }

        private void qiuhe()
        {
            int total=0;
            for (int i = 0; i < 10000; i++)
            {
                total = total + i + 1;
            }
            Console.WriteLine(total.ToString());
            

        }

        private void button3_Click(object sender, EventArgs e)
        {
            qiuhe();
            //letter();
            //letter();
            //letter();
            //letter();

            //letter(2);
            //letter(3);




            //for (int i = 1; i < 11; i = i + 2)
            //{
            //    letter(i);
            //}
            //for (int i = 0; i <5; i++)
            //{
            //    int j = 0;
            //    letter(j);
            //    j = j + 2;
            //}

            ////for (int i = 0; i < 10; i++)
            ////{
            ////    for (int j = 0; j <10; j=j+3 )
            ////    {
            ////        //letter();
            ////    }
            ////    hh();
            ////}


            //for (int i = 0; i < 4; i++)
            //{
            //    hh();
            //    for (int j = i+1; j <10; j=j+2)
            //    {
            //        letter(j);
            //    }
            //}

            ////for (int i = 0; i < 10; i++)
            ////{
            ////    for (int j = 0; j <10; j++)
            ////    {
            ////        letter();
            ////    }
            ////    hh();
            ////}





        }
    }
}
         