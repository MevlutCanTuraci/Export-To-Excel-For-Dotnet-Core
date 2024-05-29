#region Imports
using App.ExcelExport;
namespace App.ExelExport;
#endregion


class Program
{
    public static void Main()
    {
        string outPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory);
        ExcelHelper.ExportToTable(excelOutPath: outPath);
        ExcelHelper.ExportToDefalt(excelOutPath: outPath);

        Console.Read();
    }
}