using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using Tool;

namespace ChangeName
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.Write("enter file path:");
            string filePath = Console.ReadLine();
            Console.WriteLine();
            ExcelHelper eh = new ExcelHelper(filePath);
            Random rowRan = new Random(eh.FirstRowNum);
            Random columnRan = new Random(eh.FirstColumnNum);
            Console.Write("enter run times: ");

            do
            {
                int count = int.Parse(Console.ReadLine());
                // 1 //
                DateTime start = DateTime.Now;
                for (int i = 0; i < count; i++)
                {
                    //string value = eh.GetValue(rowRan.Next(eh.LastRowNum), columnRan.Next(eh.LastColumnNum));
                }
                Console.WriteLine($"get value from workbook\t\t for {count} times : {(DateTime.Now - start).TotalSeconds}");
                //start = DateTime.Now;
                //for (int i = eh.FirstRowNum; i < eh.LastRowNum; i++)
                //{
                //    for (int j = eh.FirstColumnNum; j < eh.LastColumnNum; j++)
                //    {
                //        eh.Update(rowRan.Next(), i, j);
                //    }
                //}
                //Console.WriteLine($"update value :\t {(DateTime.Now - start).TotalSeconds}");

                // 4 //
                start = DateTime.Now;
                var dic = eh.ToDictionary();
                Console.WriteLine($"init value to dictionary :\t {(DateTime.Now - start).TotalSeconds}");
                start = DateTime.Now;
                for (int i = 0; i < count; i++)
                {
                    string dicValue = dic[rowRan.Next(eh.LastRowNum)][columnRan.Next(eh.LastColumnNum)];
                }
                Console.WriteLine($"get value from dictionary\t for {count} times : {(DateTime.Now - start).TotalSeconds}");
                start = DateTime.Now;
                eh.Update(dic);
                Console.WriteLine($"update value from dictionary :\t\t {(DateTime.Now - start).TotalSeconds}");

                // 2 //
                start = DateTime.Now;
                var arr = eh.ToArray();
                Console.WriteLine($"init value to array :\t\t {(DateTime.Now - start).TotalSeconds}");
                start = DateTime.Now;
                for (int i = 0; i < count; i++)
                {
                    string arrValue = arr[rowRan.Next(eh.LastRowNum) - eh.FirstRowNum][columnRan.Next(eh.LastColumnNum) - eh.FirstColumnNum];
                }
                Console.WriteLine($"get value from array\t\t for {count} times : {(DateTime.Now - start).TotalSeconds}");
                start = DateTime.Now;
                eh.Update(arr);
                Console.WriteLine($"update value from array :\t\t {(DateTime.Now - start).TotalSeconds}");

                // 3 //
                start = DateTime.Now;
                var dt = eh.ToDataTable();
                Console.WriteLine($"init value to datatable :\t {(DateTime.Now - start).TotalSeconds}");
                start = DateTime.Now;
                for (int i = 0; i < count; i++)
                {
                    string dtValue = dt.Rows[rowRan.Next(eh.LastRowNum)][columnRan.Next(eh.LastColumnNum)].ToString();
                }
                Console.WriteLine($"get value from datatable\t for {count} times : {(DateTime.Now - start).TotalSeconds}");
                start = DateTime.Now;
                eh.Update(dt);
                Console.WriteLine($"update value from datatable :\t\t {(DateTime.Now - start).TotalSeconds}");

                // 5 //
                start = DateTime.Now;
                eh.Save("D:\\1.xlsx", true);
                Console.WriteLine($"save to disk :\t {(DateTime.Now - start).TotalSeconds}");

                Console.WriteLine("\r\n");
                Console.Write("enter run times: ");
            }
            while (true);
        }
    }
}
