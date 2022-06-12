using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using OfficeOpenXml;

namespace ExcelProject
{
    public class ExcelWorker
    {
        public static List<GaylordSpreadsheet> data = new List<GaylordSpreadsheet>();


        public static void ExctractData()
        {
            string file = @"C:\Users\meowm\Github\ExcelProject\Book.xlsx";
            using (ExcelPackage package = new ExcelPackage(new FileInfo(file)))
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                var sheet = package.Workbook.Worksheets["Sheet1"];
                var gaylordInfo = new ExcelWorker().GetList<GaylordSpreadsheet>(sheet);
                data = gaylordInfo;
            }
        }

        private List<T> GetList<T>(ExcelWorksheet sheet)
        {
            List<T> list = new List<T>();
            var columnInfo = Enumerable.Range(1, sheet.Dimension.Columns).ToList().Select(n => 
                new {Index = n, ColumnName = sheet.Cells[1, n].Value.ToString()}
            );

            for(int row = 2; row < sheet.Dimension.Rows; row++)
            {
                T obj = (T)Activator.CreateInstance(typeof(T));
                foreach(var prop in typeof(T).GetProperties())
                {
                    int col = columnInfo.SingleOrDefault(c => c.ColumnName == prop.Name).Index;
                    var val = sheet.Cells[row, col].Value;
                    var propType = prop.PropertyType;
                    prop.SetValue(obj, Convert.ChangeType(val, propType));
                }
                obj.ToString();
                list.Add(obj);
            }

            return list;
        }
    }
}