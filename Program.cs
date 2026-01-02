// See https://aka.ms/new-console-template for more information
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;

string filePath = "C:\\Users\\William\\Desktop\\Stats\\Global-Stats\\GlobalStats\\Global RR Stats Community Document.xlsx";

if (!File.Exists(filePath))
{
    Console.WriteLine($"Error: File not found at {filePath}");
    return;
}

//Statistics to collect
List<String> global_rounds = new List<String>();
List<String> rounds = new List<String>();
List<int> total_seasons = new List<int>();
List<String> seasons = new List<String>();
List<DateTime> dates = new List<DateTime>();
List<List<String>> rosterslist = new List<List<String>>();
List<String> rosters = new List<String>();
List<List<String>> victimslist = new List<List<String>>();
List<List<String>> killerslist = new List<List<String>>();
List<String> killers = new List<String>();
List<int> numberPlayers = new List<int>();

using var workbook = new XLWorkbook(filePath);
for (int sheet = 8; sheet <= workbook.Worksheets.Count; sheet++)
{
    var worksheet = workbook.Worksheet(sheet);

    //Name of round + seasons
    global_rounds.Add(worksheet.Name);
    total_seasons.Add((worksheet.Columns().Count() - 1) / 3);
    Console.WriteLine("Working on " + worksheet.Name);
    Console.WriteLine(((worksheet.Columns().Count() - 1) / 3).ToString( ) + " Seasons!");

    for (int season = 1; season <= (worksheet.Columns().Count() - 1); season+=3)
    {
        IXLCell victimsStart = worksheet.Search("Kill List (include alive)").First();
        var rangeUsed = worksheet.RangeUsed();
        int firstDataColumn = season+1;
        int firstDataRow = victimsStart.WorksheetRow().RowNumber() + 1;
        int lastDataColumn = season+3;
        int lastDataRow = rangeUsed.LastRowUsed().RowNumber();

        //Get values for the round list
        rounds.Add(worksheet.Name);
        seasons.Add(worksheet.Cell(1,firstDataColumn).GetString());
        dates.Add(worksheet.Cell(2,firstDataColumn).GetDateTime());

        //Get the roster/victims for the season
        IXLRange victimRange = worksheet.Range(firstDataRow,firstDataColumn,lastDataRow,firstDataColumn); 
        foreach (IXLCell cell in victimRange.CellsUsed())
        {
        string value = cell.GetString(); 
        rosters.Add(value);
        }
        victimslist.Add(new List<String>(rosters));
        rosters.Sort();
        rosterslist.Add(new List<String>(rosters));
        rosters = new List<String>();

        //Get the killers for the season
        IXLRange killerRange = worksheet.Range(firstDataRow,lastDataColumn,lastDataRow,lastDataColumn);
        foreach (IXLCell cell in killerRange.Cells())
        {
        string value = cell.GetString(); 
        killers.Add(value);
        }
        killerslist.Add(new List<String>(killers));
        killers = new List<String>();
    }
}

var statscompiled = new XLWorkbook();
//Making Round List Page
var roundlist = statscompiled.AddWorksheet("Round List");
roundlist.Cell("A1").InsertData(global_rounds);
roundlist.Cell("B1").InsertData(total_seasons);

//Making All Rosters Page
var allrosters = statscompiled.AddWorksheet("All Rosters");
allrosters.Column("C").Style.NumberFormat.Format = "mm/dd/yyyy";
allrosters.Column("A").Width = 34; 
allrosters.Column("B").Width = 6; 
allrosters.Column("C").Width = 20; 
allrosters.Cell("A1").InsertData(rounds);
allrosters.Cell("B1").InsertData(seasons);
allrosters.Cell("C1").InsertData(dates);
allrosters.Cell("D1").InsertData(rosterslist);
allrosters.Sort(3);

//Making Victims Page
var allvictims = statscompiled.AddWorksheet("All Victims");
allvictims.Column("C").Style.NumberFormat.Format = "mm/dd/yyyy";
allvictims.Column("A").Width = 34; 
allvictims.Column("B").Width = 6; 
allvictims.Column("C").Width = 20; 
allvictims.Cell("A1").InsertData(rounds);
allvictims.Cell("B1").InsertData(seasons);
allvictims.Cell("C1").InsertData(dates);
allvictims.Cell("D1").InsertData(victimslist);
allvictims.Sort(3);

//Making Killers Page
var allkillers = statscompiled.AddWorksheet("All Killers");
allkillers.Column("C").Style.NumberFormat.Format = "mm/dd/yyyy";
allkillers.Column("A").Width = 34; 
allkillers.Column("B").Width = 6; 
allkillers.Column("C").Width = 20; 
allkillers.Cell("A1").InsertData(rounds);
allkillers.Cell("B1").InsertData(seasons);
allkillers.Cell("C1").InsertData(dates);
allkillers.Cell("D1").InsertData(killerslist);
allkillers.Sort(3);

statscompiled.SaveAs("C:\\Users\\William\\Desktop\\Stats\\Global-Stats\\GlobalStats\\StatsCompiled.xlsx");
Console.WriteLine("Stats are now compiled!");