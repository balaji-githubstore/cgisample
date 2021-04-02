using ClosedXML.Excel;
using System;
using System.Data;

namespace ConsoleApp1
{
    class Program
    {
        static void Main(string[] args)
        {
		//bala
            //DataTable dt = new DataTable();
            //dt.Columns.Add("username");
            //dt.Columns.Add("password");

            //DataRow row1 = dt.NewRow();
            //row1[0] = "admin";
            //row1[1] = "pass123";
            //dt.Rows.Add(row1);

            //DataRow row2 = dt.NewRow();
            //row2["username"] = "john";
            //row2["password"] = "john123";
            //dt.Rows.Add(row2);


            //Console.WriteLine(dt.Rows[0]["username"]);

            XLWorkbook book = new XLWorkbook(@"D:\Report\OpenEMRData.xlsx");
            IXLWorksheet sheet = book.Worksheet("Invalid Credential");
            int rowCount = sheet.RangeUsed().RowCount();
            int colCount = sheet.RangeUsed().ColumnCount();

            DataTable dt = new DataTable();

            for(int c=1;c<= colCount; c++)
            {
                Console.WriteLine(sheet.Row(1).Cell(c).GetString());
                dt.Columns.Add(sheet.Row(1).Cell(c).GetString());              
            }

            Console.WriteLine("-----------------------------");
            for (int r = 2; r <= rowCount; r++)
            {
                DataRow row = dt.NewRow();
                for (int c = 1; c <= colCount; c++)
                {
                    Console.WriteLine(sheet.Row(r).Cell(c).GetString());
                    row[c - 1] = sheet.Row(r).Cell(c).GetString();
                }
                dt.Rows.Add(row);
            }

        }
    }
}
