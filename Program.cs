// See https://aka.ms/new-console-template for more information
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;

Console.WriteLine("Hello, World!");

string filePath = "C:\\Users\\William\\Desktop\\Stats\\Global-Stats\\GlobalStats\\Global RR Stats Community Document.xlsx";
List<String> rounds = new List<String>();
List<List<String>> victimslist = new List<List<String>>();
List<String> victims = new List<String>();
List<int> numberSeasons = new List<int>();
List<int> numberPlayers = new List<int>();

if (!File.Exists(filePath))
{
    Console.WriteLine($"Error: File not found at {filePath}");
    return;
}

using var workbook = new XLWorkbook(filePath);

for (int sheet = 8; sheet <= workbook.Worksheets.Count; sheet++)
{
    var worksheet = workbook.Worksheet(sheet);
    rounds.Add(worksheet.Name);
    numberSeasons.Add((worksheet.Columns().Count() - 1) / 3);
    Console.WriteLine("Working on " + worksheet.Name);
    Console.WriteLine(((worksheet.Columns().Count() - 1) / 3).ToString( ) + " Seasons!");
    IXLCell victimsStart = worksheet.Search("Kill List (include alive)").First();
    var rangeUsed = worksheet.RangeUsed();
    for (int season = 1; season <= (worksheet.Columns().Count() - 1); season+=3)
    {
        int firstDataColumn = season+1;
        int firstDataRow = victimsStart.WorksheetRow().RowNumber() + 1;
        int lastDataColumn = season+3;
        int lastDataRow = rangeUsed.LastRowUsed().RowNumber();
        IXLRange range = worksheet.Range(firstDataRow,firstDataColumn,lastDataRow,firstDataColumn); 
        foreach (IXLCell cell in range.CellsUsed())
        {
        string value = cell.GetString(); 
        victims.Add(value);
        }
        victims.Sort();
        victimslist.Add(new List<String>(victims));
        victims = new List<String>();
    }
}

var statscompiled = new XLWorkbook();
var allrosters = statscompiled.AddWorksheet("All Rosters");
allrosters.Cell("A1").InsertData(rounds);
allrosters.Cell("B1").InsertData(numberSeasons);
allrosters.Cell("C1").InsertData(victimslist);
statscompiled.SaveAs("C:\\Users\\William\\Desktop\\Stats\\Global-Stats\\GlobalStats\\StatsCompiled.xlsx");
Console.WriteLine("Stats are now compiled!");