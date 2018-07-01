using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;


namespace 口算题卡
{
    class ExcelClass
    {

        private Excel.Application xApp;
        private Excel.Workbook xBook;
        private Excel.Worksheet xSheet;

        public ExcelClass()
        {
            //1.创建Applicaton对象
            xApp = new Excel.Application();
            xApp.DisplayAlerts = false;
        }

        /// <summary>
        /// 打开excel文件
        /// </summary>
        /// <param name="path"></param>
        public void Open(string path)
        {
            //2.得到workbook对象，打开已有的文件
            xBook = xApp.Workbooks.Open(path);
            //3.指定要操作的Sheet
            xSheet = (Microsoft.Office.Interop.Excel.Worksheet)xBook.Sheets[1];
            
        }

        /// <summary>
        /// 关闭并保存Excel
        /// </summary>
        /// <param name="xBook"></param>
        public void Close()
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

        /// <summary>
        /// 将数据写入已存在Excel
        /// </summary>
        /// <param name="intRow">行</param>
        /// <param name="intCol">列</param>
        /// <param name="strValue"></param>
        public  void Write(int intRow,int intCol ,string strValue)
        {
            //在第一列的左边插入一列  1:第一列
            //xlShiftToRight:向右移动单元格   xlShiftDown:向下移动单元格
            //Range Columns = (Range)xSheet.Columns[1, System.Type.Missing];
            //Columns.Insert(XlInsertShiftDirection.xlShiftToRight, Type.Missing);
            xSheet.Cells[intRow, intCol] = strValue;
        }
    }
}
