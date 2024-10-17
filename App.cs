using Excel_Reader.Lawang;
using Excel_Reader.Lawang.Data;
using Spectre.Console;

internal class App
{

    private readonly Database _db;
    private readonly ExcelOperator _excelOperator;

    public App(Database db, ExcelOperator excelOperator)
    {
        _db = db;
        _excelOperator = excelOperator;
    }


    public async Task Run()
    {
        Console.Clear();
        string dbName = "Excel-Reader.db";

        //Delete the existing database if it exists
        DeleteDatabaseIfExists(dbName);

        //create Database
        AnsiConsole.Status().Start("Creating Database...", ctx => {
            Thread.Sleep(1000);
            _db.CreateDatabase();
        });

        try
        {
            //Checks wether the presence of excel file
            FileInfo excelFile = GetFileInfo();
            if (!File.Exists(excelFile.FullName))
            {
                AnsiConsole.Markup("[red bold] EXCEL FILE DOES NOT EXIST[/]");
                return;
            }

            // Read from excel file, which returns list of person
            var peopleList = await AnsiConsole.Status().StartAsync("Reading from Excel...", async ctx => {
                // To show what app is doing at current moment
                Thread.Sleep(1000);
                return await _excelOperator.ReadExcel(excelFile);
            });

            await AnsiConsole.Status().StartAsync("Seeding Data into database...", async ctx => {
                // To show what app is doing at current moment
                Thread.Sleep(1000);
                await _db.InsertData(peopleList);
            });
            
        }
        catch (Exception ex)
        {
            Console.WriteLine(ex.Message);
        }

    }

    private FileInfo GetFileInfo()
    {
        string dir = Directory.GetCurrentDirectory();
        string path = Path.Combine(dir, "sample.xlsx");

        FileInfo fileInfo = new FileInfo(path);
        if (fileInfo.Extension == ".xls")
        {
            // checks whether the Excel file is in the older Excel 97-2003 format (.xls), which is an OLE compound document. Can't read by epplus
            throw new Exception("The EPPlus library only supports the newer Excel formats (.xlsx, .xlsm, .xltx, .xltm) that are based on the Open XML standard.");
        }
        return fileInfo;
    }
    private void DeleteDatabaseIfExists(string dbName)
    {

        if (File.Exists(dbName))
        {
            File.Delete(dbName);
        }
    }


}