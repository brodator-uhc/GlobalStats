// See https://aka.ms/new-console-template for more information
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;

string filePath = "..\\..\\..\\Global RR Stats Community Document.xlsx";

if (!File.Exists(filePath))
{
    Console.WriteLine($"Error: File not found at {filePath}");
    return;
}

//TO ADD
//Seasons analysed, rounds analysed, unique players
//Largets roster,Smallest roster, Smallest roster (1+ season)
//Global stats
//Kill Records

//Statistics to collect
char separator = ',';
List<String> global_rounds = new List<String>();
List<String> rounds = new List<String>();
List<String> NRrounds = new List<String>();
List<int> total_seasons = new List<int>();
List<DateTime> round_debuts = new List<DateTime>();
List<String> seasons = new List<String>();
List<String> NRseasons = new List<String>();
List<DateTime> dates = new List<DateTime>();
List<DateTime> NRdates = new List<DateTime>();
List<List<String>> rosterslist = new List<List<String>>();
List<String> rosters = new List<String>();
List<String> rs_players = new List<String>();
List<int> roster_sizes = new List<int>();
List<String> victims = new List<String>();
List<String> methods = new List<String>();
List<String> teams = new List<String>();
List<String> teamcolors = new List<String>();
List<String> killers = new List<String>();
List<int> numberPlayers = new List<int>();

List<String> kl_rounds = new List<String>();
List<String> kl_seasons = new List<String>();
List<DateTime> kl_dates = new List<DateTime>();
List<String> tl_rounds = new List<String>();
List<String> tl_seasons = new List<String>();
List<DateTime> tl_dates = new List<DateTime>();

Dictionary<String, String> playerDebutsRound = new Dictionary<String, String>();
Dictionary<String, String> playerDebutsSeason = new Dictionary<String, String>();
Dictionary<String, DateTime> playerDebutsDate = new Dictionary<String, DateTime>();
Dictionary<String, String> FRplayerDebutsRound = new Dictionary<String, String>();
Dictionary<String, String> FRplayerDebutsSeason = new Dictionary<String, String>();
Dictionary<String, DateTime> FRplayerDebutsDate = new Dictionary<String, DateTime>();
Dictionary<String, int> gs_seasonsplayed = new Dictionary<String, int>();
Dictionary<String, int> gs_wins = new Dictionary<String, int>();
Dictionary<String, int> gs_alive = new Dictionary<String, int>();
Dictionary<String, int> gs_runnerup = new Dictionary<String, int>();
Dictionary<String, int> gs_kills = new Dictionary<String, int>();
Dictionary<String, int> gs_topfrags = new Dictionary<String, int>();
Dictionary<String, int> gs_pve = new Dictionary<String, int>();
Dictionary<String, int> gs_firstblood = new Dictionary<String, int>();
Dictionary<String, int> gs_firstdeath = new Dictionary<String, int>();
Dictionary<String, int> gs_ironman = new Dictionary<String, int>();
Dictionary<String, int> gs_firstdamage = new Dictionary<String, int>();
Dictionary<String, int> gs_deaths = new Dictionary<String, int>();
Dictionary<String, int> gs_totaluniques = new Dictionary<String, int>();
Dictionary<String, Double> gs_kdr = new Dictionary<String, Double>();
Dictionary<String, Double> gs_kpr = new Dictionary<String, Double>();

using var workbook = new XLWorkbook(filePath);
for (int sheet = 8; sheet <= workbook.Worksheets.Count; sheet++)
{
    var worksheet = workbook.Worksheet(sheet);

    //Name of round + seasons
    global_rounds.Add(worksheet.Name);
    total_seasons.Add((worksheet.Columns().Count() - 1) / 3);
    round_debuts.Add(worksheet.Cell(2,2).GetDateTime());
    Console.WriteLine("Working on " + worksheet.Name);
    Console.WriteLine(((worksheet.Columns().Count() - 1) / 3).ToString( ) + " Seasons!");

    for (int season = 1; season <= (worksheet.Columns().Count() - 1); season+=3)
    {
        IXLCell victimsStart = worksheet.Search("Kill List (include alive)").First();
        IXLCell winnerCell = worksheet.Cell(1,1);
        Dictionary<String, int> killcounts = new Dictionary<String, int>();
        String winningTeam = "";
        var rangeUsed = worksheet.RangeUsed();
        int firstDataColumn = season+1;
        int firstDataRow = victimsStart.WorksheetRow().RowNumber() + 1;
        int lastDataColumn = season+3;
        int lastDataRow = rangeUsed.LastRowUsed().RowNumber();
        int middleDataColumn = season+2;
        int seasonSize = 0;
        int teamSize = 0;
        int fbfind = 0;

        //Get values for the round list
        rounds.Add(worksheet.Name);
        seasons.Add(worksheet.Cell(1,firstDataColumn).GetString());
        dates.Add(worksheet.Cell(2,firstDataColumn).GetDateTime());

        if (!worksheet.Cell(1,lastDataColumn).GetString().Equals("NR"))
        {
            NRrounds.Add(worksheet.Name);
            NRseasons.Add(worksheet.Cell(1,firstDataColumn).GetString());
            NRdates.Add(worksheet.Cell(2,firstDataColumn).GetDateTime()); 
        }

        //Get the roster/victims for the season
        IXLRange victimRange = worksheet.Range(firstDataRow,firstDataColumn,lastDataRow,firstDataColumn);
        seasonSize = victimRange.RowsUsed().Count();
        foreach (IXLCell cell in victimRange.CellsUsed())
        {
        string value = cell.GetString(); 
        victims.Add(value);

        if (playerDebutsRound.ContainsKey(value)){
                if (worksheet.Cell(2,firstDataColumn).GetDateTime() < playerDebutsDate[value])
                {
                    playerDebutsRound[value] = worksheet.Name;
                    playerDebutsSeason[value] = worksheet.Cell(1,firstDataColumn).GetString();
                    playerDebutsDate[value] = worksheet.Cell(2,firstDataColumn).GetDateTime();
                }
            } else
            {
                playerDebutsRound.Add(value,worksheet.Name);
                playerDebutsSeason.Add(value,worksheet.Cell(1,firstDataColumn).GetString());
                playerDebutsDate.Add(value,worksheet.Cell(2,firstDataColumn).GetDateTime());
                gs_seasonsplayed.Add(value,0);
                gs_wins.Add(value,0);
                gs_alive.Add(value,0);
                gs_runnerup.Add(value,0);
                gs_kills.Add(value,0);
                gs_topfrags.Add(value,0);
                gs_pve.Add(value,0);
                gs_firstblood.Add(value,0);
                gs_firstdeath.Add(value,0);
                gs_ironman.Add(value,0);
                gs_firstdamage.Add(value,0);
                gs_deaths.Add(value,0);
                gs_totaluniques.Add(value,0);
                gs_kdr.Add(value,0);
                gs_kpr.Add(value,0);
            }

        if (FRplayerDebutsRound.ContainsKey(value)){
                if (!worksheet.Cell(1,lastDataColumn).GetString().Equals("NR"))
                {
                    if (worksheet.Cell(2,firstDataColumn).GetDateTime() < FRplayerDebutsDate[value])
                    {
                        FRplayerDebutsRound[value] = worksheet.Name;
                        FRplayerDebutsSeason[value] = worksheet.Cell(1,firstDataColumn).GetString();
                        FRplayerDebutsDate[value] = worksheet.Cell(2,firstDataColumn).GetDateTime();
                    }
                }
            } else
            {
                if (!worksheet.Cell(1,lastDataColumn).GetString().Equals("NR"))
                {
                    FRplayerDebutsRound.Add(value,worksheet.Name);
                    FRplayerDebutsSeason.Add(value,worksheet.Cell(1,firstDataColumn).GetString());
                    FRplayerDebutsDate.Add(value,worksheet.Cell(2,firstDataColumn).GetDateTime());
                }
            }    

        if (gs_deaths.ContainsKey(value))
            {
                if (!worksheet.Cell(cell.WorksheetRow().RowNumber(),lastDataColumn).GetString().Equals("Nothing"))
                {
                    gs_deaths[value] += 1;
                } else
                {
                    gs_alive[value] += 1;
                }
            }

        if (!rosters.Contains(value))
            {
                rosters.Add(value);

                if (gs_seasonsplayed.ContainsKey(value))
                {
                    gs_seasonsplayed[value] += 1;
                }

            }

        if (!rs_players.Contains(value))
            {
                rs_players.Add(value); 

                if (gs_totaluniques.ContainsKey(value))
                {
                    gs_totaluniques[value] += 1;
                }
            }
        }
        rosters.Sort();
        rosterslist.Add(rosters);
        rosters = new List<String>();

        gs_firstdeath[worksheet.Cell(firstDataRow,firstDataColumn).GetString()] += 1;

        IXLRange ironmanRange = worksheet.Range(5,firstDataColumn,5,lastDataColumn);
        IXLRange POOironmanRange = worksheet.Range(5,firstDataColumn,9,lastDataColumn);
        if (worksheet.Name.Equals("Party of One"))
        {
            foreach (IXLCell cell in POOironmanRange.CellsUsed())
            {
                string value = cell.GetString(); 
                gs_ironman[value] += 1;
            }
        } else
        {
            foreach (IXLCell cell in ironmanRange.CellsUsed())
            {
                string value = cell.GetString(); 
                gs_ironman[value] += 1;
            }
        }

        IXLRange fdRange = worksheet.Range(7,firstDataColumn,7,lastDataColumn);
        IXLRange POOfdRange = worksheet.Range(7,firstDataColumn,7,lastDataColumn);
        if (worksheet.Name.Equals("Party of One"))
        {
            foreach (IXLCell cell in POOfdRange.CellsUsed())
            {
                string value = cell.GetString(); 
                gs_firstdamage[value] += 1;
            }
        } else
        {
            foreach (IXLCell cell in fdRange.CellsUsed())
            {
                string value = cell.GetString(); 
                gs_firstdamage[value] += 1;
            }
        }

        //Get the killers for the season
        IXLRange killerRange = worksheet.Range(firstDataRow,lastDataColumn,firstDataRow+(seasonSize-1),lastDataColumn);
        foreach (IXLCell cell in killerRange.Cells())
        {
        string value = cell.GetString(); 
        killers.Add(value);
        kl_rounds.Add(worksheet.Name);
        kl_seasons.Add(worksheet.Cell(1,firstDataColumn).GetString());
        kl_dates.Add(worksheet.Cell(2,firstDataColumn).GetDateTime());

        if (gs_kills.ContainsKey(value))
            {
                gs_kills[value] += 1;
                if (gs_deaths[value].Equals(0))
                {
                    gs_kdr[value] = Convert.ToDouble(gs_kills[value]);
                } else
                {
                    gs_kdr[value] = Convert.ToDouble(gs_kills[value])/Convert.ToDouble(gs_deaths[value]);
                }
                gs_kpr[value] = Convert.ToDouble(gs_kills[value])/Convert.ToDouble(gs_seasonsplayed[value]); 

                if (killcounts.ContainsKey(value))
                {
                    killcounts[value] += 1;
                }
                else
                {
                    killcounts.Add(value,1);
                }

                if (fbfind == 0)
                {
                    gs_firstblood[value] += 1;
                    fbfind += 1;
                }
            } else
            {
                if (!value.Equals("Nothing"))
                {
                    gs_pve[worksheet.Cell(cell.WorksheetRow().RowNumber(),firstDataColumn).GetString()] += 1;
                }
            }
        }

        if (!worksheet.Name.Equals("PolyCraft Egg Hunt"))
        {
            int topFragAmount = killcounts.Values.Max();
            foreach (String killer in killcounts.Keys)
            {
                if (killcounts[killer] == topFragAmount)
                {
                    gs_topfrags[killer] += 1;
                }
            } 
        }

        //Get the methods for the season
        IXLRange methodRange = worksheet.Range(firstDataRow,middleDataColumn,firstDataRow+(seasonSize-1),middleDataColumn);
        foreach (IXLCell cell in methodRange.Cells())
        {
        string value = cell.GetString(); 
        methods.Add(value);
        }

        //Get the teams for the season
        IXLRange teamRange = worksheet.Range(9,firstDataColumn,firstDataRow-2,firstDataColumn);
        teamSize = teamRange.RowsUsed().Count();
        List<String> seasonTeams = new List<String>();

        if (!worksheet.Cell(3,middleDataColumn).GetString().Equals("FFA")){
            foreach (IXLCell cell in teamRange.CellsUsed())
            {
                string value = cell.GetString(); 
                teams.Add(value);
                seasonTeams.Add(value);
                tl_rounds.Add(worksheet.Name);
                tl_seasons.Add(worksheet.Cell(1,firstDataColumn).GetString());
                tl_dates.Add(worksheet.Cell(2,firstDataColumn).GetDateTime());
            }

            IXLRange teamColorsRange = worksheet.Range(9,lastDataColumn,9+(teamSize-1),lastDataColumn);
            foreach (IXLCell cell in teamColorsRange.Cells())
            {
                string value = cell.GetString(); 
                teamcolors.Add(value);
            }
        }

        //Get the winners

        if (worksheet.Cell(firstDataRow + (seasonSize - 1), lastDataColumn).GetString().Equals("Nothing"))
        {
            String seasonWinner = worksheet.Cell(firstDataRow + (seasonSize - 1), firstDataColumn).GetString();
            winnerCell = worksheet.Cell(firstDataRow + (seasonSize - 1), firstDataColumn);
            if (worksheet.Cell(3, middleDataColumn).GetString().Equals("FFA"))
            {
                if (gs_wins.ContainsKey(seasonWinner))
                {
                    gs_wins[seasonWinner] += 1;
                }
            } else
            {
                foreach (String wteam in seasonTeams)
                {
                    if (wteam.Contains(seasonWinner))
                    {
                        winningTeam = wteam;
                        String[] winners = wteam.Split(separator);
                        foreach (String winner in winners)
                        {
                            if (gs_wins.ContainsKey(winner))
                            {
                                gs_wins[winner] += 1;
                            }
                        }
                    }
                }
            }
        } else
        {
            // TO FIX LATER
            String seasonWinner = worksheet.Cell(firstDataRow + (seasonSize - 1), firstDataColumn).GetString();
            winnerCell = worksheet.Cell(firstDataRow + (seasonSize - 1), firstDataColumn);
            if (worksheet.Cell(3, middleDataColumn).GetString().Equals("FFA"))
            {
                if (gs_wins.ContainsKey(seasonWinner))
                {
                    gs_wins[seasonWinner] += 1;
                }
            } else
            {
                foreach (String wteam in seasonTeams)
                {
                    if (wteam.Contains(seasonWinner))
                    {
                        winningTeam = wteam;
                        String[] winners = wteam.Split(separator);
                        foreach (String winner in winners)
                        {
                            if (gs_wins.ContainsKey(winner))
                            {
                                gs_wins[winner] += 1;
                            }
                        }
                    }
                }
            }
        }

        //Get the runner ups

        if (worksheet.Cell(3, middleDataColumn).GetString().Equals("FFA"))
        {
            if (gs_runnerup.ContainsKey(winnerCell.CellAbove().GetString()))
            {
                gs_runnerup[winnerCell.CellAbove().GetString()] += 1;
            }
        } else
        {
            while (winningTeam.Contains(winnerCell.CellAbove().GetString()))
            {
                winnerCell = winnerCell.CellAbove();
            }

            String seasonRunnerUp = winnerCell.CellAbove().GetString();
            foreach (String ruteam in seasonTeams)
            {
                if (ruteam.Contains(seasonRunnerUp))
                {
                    String[] runnerups = ruteam.Split(separator);
                    foreach (String runnerup in runnerups)
                    {
                        if (gs_runnerup.ContainsKey(runnerup))
                        {
                            gs_runnerup[runnerup] += 1;
                        }
                    }
                }
            }
        }
    }

    roster_sizes.Add(rs_players.Count);
    rs_players = new List<String>();
}

var roundstats = new XLWorkbook();
var statscompiled = new XLWorkbook();
var rrdebut = new XLWorkbook();
var globalrrstats = new XLWorkbook();

//Making Round List Page
var roundlist = roundstats.AddWorksheet("Round List");
roundlist.Column("D").Style.NumberFormat.Format = "dd mmm, yyyy";
roundlist.Column("A").Width = 34; 
roundlist.Column("B").Width = 6; 
roundlist.Column("C").Width = 6; 
roundlist.Column("D").Width = 20; 
roundlist.Cell("A1").InsertData(global_rounds);
roundlist.Cell("B1").InsertData(total_seasons);
roundlist.Cell("C1").InsertData(roster_sizes);
roundlist.Cell("D1").InsertData(round_debuts);
roundlist.Sort(4);

//Making All Rosters Page
var allrosters = roundstats.AddWorksheet("All Rosters");
allrosters.Column("C").Style.NumberFormat.Format = "mm/dd/yyyy";
allrosters.Column("A").Width = 34; 
allrosters.Column("B").Width = 6; 
allrosters.Column("C").Width = 20; 
allrosters.Cell("A1").InsertData(rounds);
allrosters.Cell("B1").InsertData(seasons);
allrosters.Cell("C1").InsertData(dates);
allrosters.Cell("D1").InsertData(rosterslist);
allrosters.Sort(3);

//Making NR All Rosters Page
var allrostersnr = roundstats.AddWorksheet("All Rosters (NR)");
allrostersnr.Column("C").Style.NumberFormat.Format = "mm/dd/yyyy";
allrostersnr.Column("A").Width = 34; 
allrostersnr.Column("B").Width = 6; 
allrostersnr.Column("C").Width = 20; 
allrostersnr.Cell("A1").InsertData(NRrounds);
allrostersnr.Cell("B1").InsertData(NRseasons);
allrostersnr.Cell("C1").InsertData(NRdates);
allrostersnr.Sort(3);

//Making Kills Page
var allkills = statscompiled.AddWorksheet("All Kills");
allkills.Column("C").Style.NumberFormat.Format = "mm/dd/yyyy";
allkills.Column("A").Width = 34; 
allkills.Column("B").Width = 6; 
allkills.Column("C").Width = 20; 
allkills.Cell("A1").InsertData(kl_rounds);
allkills.Cell("B1").InsertData(kl_seasons);
allkills.Cell("C1").InsertData(kl_dates);
allkills.Cell("D1").InsertData(victims);
allkills.Cell("E1").InsertData(methods);
allkills.Cell("F1").InsertData(killers);
allkills.Sort(3);

//Making Teams Page
var allteams = statscompiled.AddWorksheet("All Teams");
allteams.Column("C").Style.NumberFormat.Format = "mm/dd/yyyy";
allteams.Column("A").Width = 34; 
allteams.Column("B").Width = 6; 
allteams.Column("C").Width = 20; 
allteams.Cell("A1").InsertData(tl_rounds);
allteams.Cell("B1").InsertData(tl_seasons);
allteams.Cell("C1").InsertData(tl_dates);
allteams.Cell("D1").InsertData(teams);
allteams.Cell("E1").InsertData(teamcolors);
//allteams.Sort(3);

//RR Debuts
var rrdebuts = rrdebut.AddWorksheet("RR Debuts");
rrdebuts.Column("C").Style.NumberFormat.Format = "mm/dd/yyyy";
rrdebuts.Column("A").Width = 34; 
rrdebuts.Column("B").Width = 6; 
rrdebuts.Column("C").Width = 20; 
rrdebuts.Cell("A1").InsertData(playerDebutsRound.Values);
rrdebuts.Cell("B1").InsertData(playerDebutsSeason.Values);
rrdebuts.Cell("C1").InsertData(playerDebutsDate.Values);
rrdebuts.Cell("D1").InsertData(playerDebutsRound.Keys);
rrdebuts.Sort(3);

//RR Debuts (No NR)
var rrdebutsnonr = rrdebut.AddWorksheet("RR Debuts (No NR)");
rrdebutsnonr.Column("C").Style.NumberFormat.Format = "mm/dd/yyyy";
rrdebutsnonr.Column("A").Width = 34; 
rrdebutsnonr.Column("B").Width = 6; 
rrdebutsnonr.Column("C").Width = 20; 
rrdebutsnonr.Cell("A1").InsertData(FRplayerDebutsRound.Values);
rrdebutsnonr.Cell("B1").InsertData(FRplayerDebutsSeason.Values);
rrdebutsnonr.Cell("C1").InsertData(FRplayerDebutsDate.Values);
rrdebutsnonr.Cell("D1").InsertData(FRplayerDebutsRound.Keys);
rrdebutsnonr.Sort(3);

//Global Stats
var globalstats = globalrrstats.AddWorksheet("Global Stats");
globalstats.Cell("A1").InsertData(gs_seasonsplayed.Keys);
globalstats.Cell("B1").InsertData(gs_seasonsplayed.Values);
globalstats.Cell("C1").InsertData(gs_wins.Values);
globalstats.Cell("D1").InsertData(gs_alive.Values);
globalstats.Cell("E1").InsertData(gs_runnerup.Values);
globalstats.Cell("F1").InsertData(gs_kills.Values);
globalstats.Cell("G1").InsertData(gs_topfrags.Values);
globalstats.Cell("H1").InsertData(gs_pve.Values);
globalstats.Cell("I1").InsertData(gs_firstblood.Values);
globalstats.Cell("J1").InsertData(gs_firstdeath.Values);
globalstats.Cell("K1").InsertData(gs_ironman.Values);
globalstats.Cell("L1").InsertData(gs_firstdamage.Values);
globalstats.Cell("M1").InsertData(gs_deaths.Values);
globalstats.Cell("N1").InsertData(gs_totaluniques.Values);
globalstats.Cell("O1").InsertData(gs_kdr.Values);
globalstats.Cell("P1").InsertData(gs_kpr.Values);
globalstats.Sort(1);

rrdebut.SaveAs("..\\..\\..\\RR_Debuts.xlsx");
roundstats.SaveAs("..\\..\\..\\Round_Stats.xlsx");
statscompiled.SaveAs("..\\..\\..\\Stats_Compiled.xlsx");
globalrrstats.SaveAs("..\\..\\..\\Global_Stats.xlsx");
Console.WriteLine("Stats are now compiled!");