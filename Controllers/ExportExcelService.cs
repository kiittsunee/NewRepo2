using ClosedXML.Excel;
using DotNetApi.Models;
using Microsoft.EntityFrameworkCore;
using System.Data;
using System.Reflection;

namespace TodoApi.Controllers
{
    public class ExportExcelService : IExportExcelService
    {
        private string GetExcelColumnName(int columnNumber)
        {
            string columnName = "";

            while (columnNumber > 0)
            {
                int modulo = (columnNumber - 1) % 26;
                columnName = Convert.ToChar('A' + modulo) + columnName;
                columnNumber = (columnNumber - modulo) / 26;
            }

            return columnName;
        }
        public DataTable TodoItemTable(DbSet<TodoItem> TodoItems)
        {
            DataTable dataTable = new DataTable(typeof(TodoItem).Name);
            PropertyInfo[] Properties = typeof(TodoItem).GetProperties(bindingAttr: BindingFlags.Public | BindingFlags.Instance) ;
            foreach (PropertyInfo property in Properties)
            {
                dataTable.Columns.Add(property.Name);
            }
            foreach (TodoItem item in TodoItems)
            {
                var values = new object[Properties.Length];
                for(int i=0; i<Properties.Length; i++)
                {
#pragma warning disable CS8601 // Возможно, назначение-ссылка, допускающее значение NULL.
                    values[i] = Properties[i].GetValue(item, index: null);
#pragma warning restore CS8601 // Возможно, назначение-ссылка, допускающее значение NULL.
                }
                dataTable.Rows.Add(values);
            }
            return dataTable;
        }

        public XLWorkbook ExportExcel(DataTable dt, string title = "", string fileName = "")
        {
            int columsNumbers = dt.Columns.Count;
            XLWorkbook wb = new XLWorkbook();
            IXLWorksheet ws = wb.Worksheets.Add(dt, sheetName: "Лист 1");

            if (title.Length > 0) 
            {
                ws.Row(1).InsertRowsAbove(1);
                ws.Cell("A1").Value = title;
                string columnRange = "A" + GetExcelColumnName(columsNumbers) + "1";
                ws.Range(columnRange).Merge();
            }

            ws.Columns().AdjustToContents();
            return wb;
        }
    }
}
