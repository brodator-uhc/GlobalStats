// See https://aka.ms/new-console-template for more information
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;

Console.WriteLine("Hello, World!");

string filePath = "C:\\Users\\William\\Desktop\\Stats\\Global-Stats\\GlobalStats\\Global RR Stats Community Document.xlsx";
List<String> rounds = new List<String>();

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
    Console.WriteLine("Working on " + worksheet.Name);
    for (int season = 1; season <= (worksheet.ColumnCount() - 1) / 3; season++)
    {
        
    }
}

var statscompiled = new XLWorkbook();
var allrosters = statscompiled.AddWorksheet("All Rosters");
allrosters.Cell("A1").InsertData(rounds);
statscompiled.SaveAs("C:\\Users\\William\\Desktop\\Stats\\Global-Stats\\GlobalStats\\StatsCompiled.xlsx");
Console.WriteLine("Stats are now compiled!");


//Console.WriteLine($"Reading data from '{worksheet.Name}'...");

// Iterate over the rows and columns
// ClosedXML uses 1-based indexing for rows and columns
// Read the cell value and print it to the console
//var cellValue = worksheet.Cell(row, col).Value.ToString();
//Console.Write(cellValue + "\t");

//Console.WriteLine();
