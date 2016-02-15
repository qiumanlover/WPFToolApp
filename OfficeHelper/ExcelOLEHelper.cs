using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeHelper
{
    public class ExcelOLEHelper
    {

        public void ReadExcel()
        {
            //Application excel = new Application();
            //var sheet = excel.Worksheets[1];
            //sheet.Shapes.AddOLEObject

            object oMissing = System.Reflection.Missing.Value;
            Microsoft.Office.Interop.Word.Application word = new Microsoft.Office.Interop.Word.Application();//创建word对象
            word.Visible = true;//显示出来
            Microsoft.Office.Interop.Word.Document dcu = word.Documents.Add(ref oMissing, ref oMissing, ref oMissing, ref oMissing);//创建一个新的空文档，格式为默认的
            dcu.Activate();//激活当前文档
            object type = @"Excel.Sheet.12";//插入的excel 格式，这里我用的是excel 2010，所以是.12
            object filename = @"D:\Life.xlsx";//插入的excel的位置
            word.Selection.InlineShapes.AddOLEObject(ref type, ref filename, ref oMissing, ref oMissing);//执行插入操作


            //Console.ReadKey();
        }
    }
}
