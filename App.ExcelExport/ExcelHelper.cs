#region Imports
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using GenFu;
using System;
#endregion


namespace App.ExcelExport
{
    public class ExcelHelper
    {
        private Dictionary<string, int>? _headerMapping;

        //Only for generic funcs
        public void ConfigureHeaders(Dictionary<string, int> headerMapping)
        {
            //_headerMapping = new Dictionary<string, int>();
            _headerMapping = headerMapping;
        }


        private static IEnumerable<EmployeeModel> GenerateData()
        {
            GenFu.GenFu.Configure<EmployeeModel>()
            .Fill(p => p.Name).AsFirstName()
            .Fill(p => p.Surname).AsLastName()
            .Fill(p => p.Age).WithinRange(19, 45)
            .Fill(p => p.Email).AsEmailAddress()
            .Fill(p => p.Addess).AsAddress()
            .Fill(p => p.Bio).AsLoremIpsumSentences();

            return A.ListOf<EmployeeModel>();
        }

        public static void ExportToTable(string excelOutPath)
        {
            var users = GenerateData();

            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Employee");

                // Table headers
                worksheet.Cell(1, 1).Value = "Id";
                worksheet.Cell(1, 2).Value = "Name";
                worksheet.Cell(1, 3).Value = "Surname";
                worksheet.Cell(1, 4).Value = "Age";
                worksheet.Cell(1, 5).Value = "Email";
                worksheet.Cell(1, 6).Value = "Address";
                worksheet.Cell(1, 7).Value = "Bio";

                // Adding data to table
                int currentRow = 2;
                foreach (var user in users)
                {
                    worksheet.Cell(currentRow, 1).Value = user.Id;
                    worksheet.Cell(currentRow, 2).Value = user.Name;
                    worksheet.Cell(currentRow, 3).Value = user.Surname;
                    worksheet.Cell(currentRow, 4).Value = user.Age;
                    worksheet.Cell(currentRow, 5).Value = user.Email;
                    worksheet.Cell(currentRow, 6).Value = user.Addess;
                    worksheet.Cell(currentRow, 7).Value = user.Bio;
                    
                    currentRow++;
                }

                // Create table
                var range = worksheet.Range(1, 1, currentRow - 1, /* Last cell number  -> */ 7 /* <- */);
                var table = range.CreateTable();

                // Change table theme
                table.Theme = XLTableTheme.TableStyleMedium2; // Örnek bir tema

                // opsionel
                //table.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                //table.Style.Border.OutsideBorder = XLBorderStyleValues.Thick;
                //table.Style.Border.InsideBorder = XLBorderStyleValues.Thin;

                // Expand by the contents of each column
                worksheet.Columns().AdjustToContents().AdjustToContents();

                // Save the file
                using (var stream = new MemoryStream())
                {
                    workbook.SaveAs(stream);
                    var content = stream.ToArray();

                    var fileName = string.Concat("Table", "_", Guid.NewGuid().ToString("N"), ".xlsx");
                    var appPath = Path.Combine(excelOutPath, fileName);
                    File.WriteAllBytes(appPath, content);

                    Console.WriteLine("Saved the table theme excel file.");
                }
            }
        }

        public static void ExportToDefalt(string excelOutPath)
        {
            var users = GenerateData();

            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Employee");
                var currentRow = 1;

                // Table headers
                worksheet.Cell(currentRow, 1).Value = "Id";          //A1
                worksheet.Cell(currentRow, 2).Value = "Name";        //B1
                worksheet.Cell(currentRow, 3).Value = "Surname";     //C1
                worksheet.Cell(currentRow, 4).Value = "Age";         //D1
                worksheet.Cell(currentRow, 5).Value = "Email";       //E1
                worksheet.Cell(currentRow, 6).Value = "Address";     //F1
                worksheet.Cell(currentRow, 7).Value = "Bio";         //G1

                // Make column headings bold
                worksheet.Range("A1:G1").Style.Font.Bold = true;
                // Color the background of column headings
                worksheet.Range("A1:G1").Style.Fill.BackgroundColor = XLColor.LightGray;

                foreach (var user in users)
                {
                    currentRow++;
                    worksheet.Cell(currentRow, 1).Value = user.Id;
                    worksheet.Cell(currentRow, 2).Value = user.Name;
                    worksheet.Cell(currentRow, 3).Value = user.Surname;
                    worksheet.Cell(currentRow, 4).Value = user.Age;
                    worksheet.Cell(currentRow, 5).Value = user.Email;
                    worksheet.Cell(currentRow, 6).Value = user.Addess;
                    worksheet.Cell(currentRow, 7).Value = user.Bio;
                }


                // Expand by the contents of each column
                worksheet.Columns().AdjustToContents();

                using (var stream = new MemoryStream())
                {
                    workbook.SaveAs(stream);
                    var content = stream.ToArray();

                    var fileName = string.Concat("Default", "_", Guid.NewGuid().ToString("N"), ".xlsx");
                    var appPath = Path.Combine(excelOutPath, fileName);
                    File.WriteAllBytes(appPath, content);

                    Console.WriteLine("Saved the default theme excel file.");
                }
            }
        }

        public static void ExportWithUseGeneric(string excelOutPath)
        {
            var users = GenerateData();

            ExcelHelper ex = new ExcelHelper();
            ex.ConfigureHeaders(new Dictionary<string, int>
            {
                { /* Header name */ nameof(EmployeeModel.Id),      /* Index */ 1 },
                { /* Header name */ nameof(EmployeeModel.Name),    /* Index */ 2 },
                { /* Header name */ nameof(EmployeeModel.Age),     /* Index */ 3 },
                { /* Header name */ nameof(EmployeeModel.Addess),  /* Index */ 4 },
                { /* Header name */ nameof(EmployeeModel.Bio),     /* Index */ 5 },
                
            });

            ex.ExportToDefaltGeneric<EmployeeModel>("Employee With Generic", users.ToList(), (model, worksheet, rowId) =>
            {
                ex.AddCellValue(worksheet, rowId, /* Header name */ nameof(EmployeeModel.Id),         /* Header value */ model.Id);
                ex.AddCellValue(worksheet, rowId, /* Header name */ nameof(EmployeeModel.Name),       /* Header value */ model.Name);
                ex.AddCellValue(worksheet, rowId, /* Header name */ nameof(EmployeeModel.Age),        /* Header value */ model.Age);
                ex.AddCellValue(worksheet, rowId, /* Header name */ nameof(EmployeeModel.Addess),     /* Header value */ model.Addess);
                ex.AddCellValue(worksheet, rowId, /* Header name */ nameof(EmployeeModel.Bio),        /* Header value */ model.Bio);
            });
        }

        /* Generic funcs */

        public void ExportToDefaltGeneric<T>(string v, List<T> employeeModels, Action<T, IXLWorksheet, int> value)
        {
            if (_headerMapping == null)
            {
                throw new NullReferenceException(nameof(ConfigureHeaders) + " For " + nameof(_headerMapping));
            }

            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add(v);

                // Add headers
                foreach (var header in _headerMapping!)
                {
                    worksheet.Cell(1, header.Value).Value = header.Key;
                }

                int currentRow = 2;
                foreach (var item in employeeModels)
                {
                    value.Invoke(item, worksheet, currentRow);
                    currentRow++;
                }

                worksheet.Columns().AdjustToContents().AdjustToContents();

                // Save the file
                using (var stream = new MemoryStream())
                {
                    workbook.SaveAs(stream);
                    var content = stream.ToArray();

                    var fileName = string.Concat("Default-Generic", "_", Guid.NewGuid().ToString("N"), ".xlsx");
                    var appPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, fileName);
                    File.WriteAllBytes(appPath, content);

                    Console.WriteLine("Saved the default theme (generic) excel file.");
                }
            }
        }


        //Only for generic funcs
        public void AddCellValue(IXLWorksheet worksheet, int rowId, string headerName, XLCellValue value)
        {
            var index = _headerMapping
                .Where(p => p.Key.Equals(headerName, StringComparison.OrdinalIgnoreCase))
                .Select(s => s.Value)
                .FirstOrDefault();

            worksheet.Cell(rowId, index).Value = value;
        }



    }
}