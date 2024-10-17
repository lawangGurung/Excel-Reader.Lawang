
using Excel_Reader.Lawang.Model;
using OfficeOpenXml;
using Spectre.Console;

namespace Excel_Reader.Lawang;

public class ExcelOperator
{

    public async Task<List<Person>> ReadExcel(FileInfo excelFile)
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        List<Person> people = new();

        using var package = new ExcelPackage(excelFile);
        await package.LoadAsync(excelFile);

        ExcelWorksheet worksheet = package.Workbook.Worksheets[0];

        int rowCount = worksheet.Dimension.Rows;
        
        //Assumes the data row starts at row 2
        for(int i = 2; i <= rowCount; i++)
        {
            var person = new Person()
            {
                Id = int.Parse(worksheet.Cells[i, 1].Text),
                FirstName = worksheet.Cells[i, 2].Text,
                LastName = worksheet.Cells[i, 3].Text,
                Gender = worksheet.Cells[i, 4].Text,
                Country = worksheet.Cells[i, 5].Text,
                Age = int.Parse(worksheet.Cells[i, 6].Text),
                Date = worksheet.Cells[i, 7].Text
            };

            people.Add(person);
        }
        AnsiConsole.MarkupLine("[green bold]READING FROM EXCEL COMPLETE  :bookmark:[/]");
        return people;

    }
}
