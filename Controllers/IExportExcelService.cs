using ClosedXML.Excel;
using ClosedXML.Extensions;
using Aspose.Cells;
using DotNetApi.Models;
using System.Data;
using System.Reflection;
using Microsoft.EntityFrameworkCore;

namespace TodoApi.Controllers
{
    public interface IExportExcelService
    {
        XLWorkbook ExportExcel(DataTable dt, string title = "", string fileName = "");
        DataTable TodoItemTable(List<TodoItem> items);
    }
}
