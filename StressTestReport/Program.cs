// See https://aka.ms/new-console-template for more information
using OfficeOpenXml;
using StressTestReport.Models;

Console.WriteLine("Hello, World!");

static void run()
{



    List<CodeDetail> codeDetails = PopulateCodeDetails();

    FileInfo fileInfo = new FileInfo(@"F:\Temp\file.xlsx");

    ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

    using (ExcelPackage excelPackage = new ExcelPackage(fileInfo))
    {

        var worksheet = GetWorkSheet(excelPackage, 0);
        var worksheet1 = GetWorkSheet(excelPackage, 1);

        worksheet.Cells["B2"].LoadFromCollection(codeDetails, false, OfficeOpenXml.Table.TableStyles.Medium1);
        worksheet1.Cells["B2"].LoadFromCollection(codeDetails, false, OfficeOpenXml.Table.TableStyles.Medium1);
        excelPackage.Save();
    }

    static ExcelWorksheet GetWorkSheet(ExcelPackage excelPackage, int v)
    {
        var worksheet = excelPackage.Workbook.Worksheets.Add($"Content - {v}");
        worksheet.View.ShowGridLines = false;
        worksheet.Cells["B1"].Value = "Code";
        worksheet.Cells["C1"].Value = "Time";
        worksheet.Cells["D1"].Value = "Date";
        worksheet.Cells["B1:D1"].Style.Font.Bold = true;
        return worksheet;
    }

    static List<CodeDetail> PopulateCodeDetails()
    {
        Console.WriteLine("Populating..!");
        List<CodeDetail> codeDetails = new();
        Random random = new Random();
        for (int i = 0; i < 1_000_000; i++)
        {
            Console.WriteLine($"Populating count: {i}!");
            CodeDetail codeDetail = new CodeDetail
            {
                Code = random.Next(1232443).ToString(),
                Time = DateTime.Now.ToShortTimeString(),
                Date = DateTime.Now.ToShortDateString()
            };
            codeDetails.Add(codeDetail);

        }
        return codeDetails;
    }

}

run();