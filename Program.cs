using ClosedXML.Excel;

string filePath = "..\\..\\..\\Global RR Stats Community Document.xlsx";

if (!File.Exists(filePath))
{
    Console.WriteLine($"Error: File not found at {filePath}");
    return;
}

//TO ADD
//Stats Verification
//Add Pve/Alive list
//Add deaths list with kill list
//Exceptions

//Big list of variables that are saved at the end
//Variables for the rounds list
List<String> rl_rounds = new List<String>();
List<int> rl_seasons = new List<int>();
List<int> rl_rostersizes = new List<int>();
List<DateTime> rl_rounddebuts = new List<DateTime>();
//Variables for the all rosters doc
List<String> ar_rounds = new List<String>();
List<String> ar_seasons = new List<String>();
List<DateTime> ar_dates = new List<DateTime>();
List<List<String>> ar_rosters = new List<List<String>>();
//Variables for the all rosters without non-reddits
List<String> nr_rounds = new List<String>();
List<String> nr_seasons = new List<String>();
List<DateTime> nr_dates = new List<DateTime>();
List<List<String>> nr_rosters = new List<List<String>>();
//Variables for the kill list
List<String> kl_rounds = new List<String>();
List<String> kl_seasons = new List<String>();
List<DateTime> kl_dates = new List<DateTime>();
List<String> kl_victims = new List<String>();
List<String> kl_methods = new List<String>();
List<String> kl_killers = new List<String>();
//Variables for the team list
List<String> tl_rounds = new List<String>();
List<String> tl_seasons = new List<String>();
List<DateTime> tl_dates = new List<DateTime>();
List<String> tl_teams = new List<String>();
List<String> tl_teamcolors = new List<String>();
//Variables for the first damage list
List<String> dl_rounds = new List<String>();
List<String> dl_seasons = new List<String>();
List<DateTime> dl_dates = new List<DateTime>();
List<String> dl_players = new List<String>();
//Variables for the ironman list
List<String> il_rounds = new List<String>();
List<String> il_seasons = new List<String>();
List<DateTime> il_dates = new List<DateTime>();
List<String> il_players = new List<String>();
//Variables for the first death list
List<String> fdl_rounds = new List<String>();
List<String> fdl_seasons = new List<String>();
List<DateTime> fdl_dates = new List<DateTime>();
List<String> fdl_players = new List<String>();
//Variables for the first blood list
List<String> fbl_rounds = new List<String>();
List<String> fbl_seasons = new List<String>();
List<DateTime> fbl_dates = new List<DateTime>();
List<String> fbl_players = new List<String>();
//Variables for the top kills list
List<String> tfl_rounds = new List<String>();
List<String> tfl_seasons = new List<String>();
List<DateTime> tfl_dates = new List<DateTime>();
List<String> tfl_players = new List<String>();
//Variables for the runner up list
List<String> rul_rounds = new List<String>();
List<String> rul_seasons = new List<String>();
List<DateTime> rul_dates = new List<DateTime>();
List<String> rul_players = new List<String>();
//Variables for the wins list
List<String> wl_rounds = new List<String>();
List<String> wl_seasons = new List<String>();
List<DateTime> wl_dates = new List<DateTime>();
List<String> wl_players = new List<String>();
//Variables for the round debut list
Dictionary<String, String> dl_round = new Dictionary<String, String>();
Dictionary<String, String> dl_season = new Dictionary<String, String>();
Dictionary<String, DateTime> dl_date = new Dictionary<String, DateTime>();
//Variables for the round debut list without non-reddits
Dictionary<String, String> nrdl_round = new Dictionary<String, String>();
Dictionary<String, String> nrdl_season = new Dictionary<String, String>();
Dictionary<String, DateTime> nrdl_date = new Dictionary<String, DateTime>();
//Variables for the global stats
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
//Variables for the kill records
Dictionary<String, int> kr_killrecord = new Dictionary<String, int>();
Dictionary<String, String> kr_round = new Dictionary<String, String>();
Dictionary<String, String> kr_season = new Dictionary<String, String>();
Dictionary<String, DateTime> kr_date = new Dictionary<String, DateTime>();

//Goes through every stats tabs on the doc
using var workbook = new XLWorkbook(filePath);
for (int sheet = 8; sheet <= workbook.Worksheets.Count; sheet++)
{
    List<String> roundRoster = new List<String>();
    //Collecting the Name, the total amount of seasons & the date of S1 for the round
    var worksheet = workbook.Worksheet(sheet);
    rl_rounds.Add(worksheet.Name);
    rl_seasons.Add((worksheet.Columns().Count() - 1) / 3);
    rl_rounddebuts.Add(worksheet.Cell(2, 2).GetDateTime());
    Console.WriteLine("Working on " + worksheet.Name);
    Console.WriteLine(((worksheet.Columns().Count() - 1) / 3).ToString() + " Seasons!");

    //Goes through every season on the round sheet
    for (int season = 1; season <= (worksheet.Columns().Count() - 1); season += 3)
    {
        //Sets variables for stats logic
        List<String> seasonRoster = new List<String>();
        List<String> seasonTeams = new List<String>();
        Dictionary<String, int> killboard = new Dictionary<String, int>();
        IXLCell winnerCell = worksheet.Cell(1, 1);
        String winningTeam = "";
        char separator = ',';
        int seasonSize = 0;
        int teamSize = 0;
        int first_blood = 0;

        //Gets the row for the start of the death log & sets cell locations relative to the season
        IXLCell victimsStart = worksheet.Search("Kill List (include alive)").First();
        var rangeUsed = worksheet.RangeUsed();
        int firstDataRow = victimsStart.WorksheetRow().RowNumber() + 1;
        int lastDataRow = rangeUsed.LastRowUsed().RowNumber();
        int firstDataColumn = season + 1;
        int middleDataColumn = season + 2;
        int lastDataColumn = season + 3;

        //Get seasons data for the all rosters list
        ar_rounds.Add(worksheet.Name);
        ar_seasons.Add(worksheet.Cell(1, firstDataColumn).GetString());
        ar_dates.Add(worksheet.Cell(2, firstDataColumn).GetDateTime());

        //Get seasons data for all rounds except for the non-reddit releases
        if (!worksheet.Cell(1, lastDataColumn).GetString().Equals("NR"))
        {
            nr_rounds.Add(worksheet.Name);
            nr_seasons.Add(worksheet.Cell(1, firstDataColumn).GetString());
            nr_dates.Add(worksheet.Cell(2, firstDataColumn).GetDateTime());
        }

        //Loops through all the victim cells
        IXLRange victimRange = worksheet.Range(firstDataRow, firstDataColumn, lastDataRow, firstDataColumn);
        seasonSize = victimRange.RowsUsed().Count();
        foreach (IXLCell cell in victimRange.CellsUsed())
        {
            string value = cell.GetString();

            //Checks if its the players debut round, sets the date if it is
            //If new players sets all the variables for them
            if (dl_round.ContainsKey(value))
            {
                if (worksheet.Cell(2, firstDataColumn).GetDateTime() < dl_date[value])
                {
                    dl_round[value] = worksheet.Name;
                    dl_season[value] = worksheet.Cell(1, firstDataColumn).GetString();
                    dl_date[value] = worksheet.Cell(2, firstDataColumn).GetDateTime();
                }
            }
            else
            {
                dl_round.Add(value, worksheet.Name);
                dl_season.Add(value, worksheet.Cell(1, firstDataColumn).GetString());
                dl_date.Add(value, worksheet.Cell(2, firstDataColumn).GetDateTime());

                //Adds new player to the Global Stats
                gs_seasonsplayed.Add(value, 0);
                gs_wins.Add(value, 0);
                gs_alive.Add(value, 0);
                gs_runnerup.Add(value, 0);
                gs_kills.Add(value, 0);
                gs_topfrags.Add(value, 0);
                gs_pve.Add(value, 0);
                gs_firstblood.Add(value, 0);
                gs_firstdeath.Add(value, 0);
                gs_ironman.Add(value, 0);
                gs_firstdamage.Add(value, 0);
                gs_deaths.Add(value, 0);
                gs_totaluniques.Add(value, 0);
                gs_kdr.Add(value, 0);
                gs_kpr.Add(value, 0);
            }

            //Checks for players debut round but also excludes non-reddit rounds
            if (nrdl_round.ContainsKey(value))
            {
                if (!worksheet.Cell(1, lastDataColumn).GetString().Equals("NR"))
                {
                    if (worksheet.Cell(2, firstDataColumn).GetDateTime() < nrdl_date[value])
                    {
                        nrdl_round[value] = worksheet.Name;
                        nrdl_season[value] = worksheet.Cell(1, firstDataColumn).GetString();
                        nrdl_date[value] = worksheet.Cell(2, firstDataColumn).GetDateTime();
                    }
                }
            }
            else
            {
                if (!worksheet.Cell(1, lastDataColumn).GetString().Equals("NR"))
                {
                    nrdl_round.Add(value, worksheet.Name);
                    nrdl_season.Add(value, worksheet.Cell(1, firstDataColumn).GetString());
                    nrdl_date.Add(value, worksheet.Cell(2, firstDataColumn).GetDateTime());
                }
            }

            //If the players didn't die gets +1 alive on the global stats, else +1 death
            if (gs_deaths.ContainsKey(value))
            {
                if (worksheet.Cell(cell.WorksheetRow().RowNumber(), lastDataColumn).GetString().Equals("Nothing"))
                {
                    gs_alive[value] += 1;
                }
                else
                {
                    gs_deaths[value] += 1;
                }
            }

            //Makes roster for the season, skips players who show up twice with respawns gamemodes
            //Also adds +1 seasons played for the global stats
            if (!seasonRoster.Contains(value))
            {
                seasonRoster.Add(value);

                if (gs_seasonsplayed.ContainsKey(value))
                {
                    gs_seasonsplayed[value] += 1;
                }

            }

            //Makes roster for the round, adds new players
            //Also adds +1 unique round for the global stats
            if (!roundRoster.Contains(value))
            {
                roundRoster.Add(value);

                if (gs_totaluniques.ContainsKey(value))
                {
                    gs_totaluniques[value] += 1;
                }
            }
        }
        //Adds rosters to a list for the sheet, skips non-reddit for the alternate page
        seasonRoster.Sort();
        ar_rosters.Add(seasonRoster);
        if (!worksheet.Cell(1, lastDataColumn).GetString().Equals("NR"))
        {
            nr_rosters.Add(seasonRoster);
        }

        //Gets first death of the season
        gs_firstdeath[worksheet.Cell(firstDataRow, firstDataColumn).GetString()] += 1;

        //Add to first death list
        fdl_rounds.Add(worksheet.Name);
        fdl_seasons.Add(worksheet.Cell(1, firstDataColumn).GetString());
        fdl_dates.Add(worksheet.Cell(2, firstDataColumn).GetDateTime());
        fdl_players.Add(worksheet.Cell(firstDataRow, firstDataColumn).GetString());

        //Gets ironman for the season
        //Different Range for Party Of One since ironman takes 5 rows for that sheet
        IXLRange ironmanRange = worksheet.Range(5, firstDataColumn, 5, lastDataColumn);
        IXLRange POOironmanRange = worksheet.Range(5, firstDataColumn, 9, lastDataColumn);
        if (worksheet.Name.Equals("Party of One"))
        {
            foreach (IXLCell cell in POOironmanRange.CellsUsed())
            {
                string value = cell.GetString();
                gs_ironman[value] += 1;

                //Add to ironman list
                il_rounds.Add(worksheet.Name);
                il_seasons.Add(worksheet.Cell(1, firstDataColumn).GetString());
                il_dates.Add(worksheet.Cell(2, firstDataColumn).GetDateTime());
                il_players.Add(value);
            }
        }
        else
        {
            foreach (IXLCell cell in ironmanRange.CellsUsed())
            {
                string value = cell.GetString();
                gs_ironman[value] += 1;

                //Add to ironman list
                il_rounds.Add(worksheet.Name);
                il_seasons.Add(worksheet.Cell(1, firstDataColumn).GetString());
                il_dates.Add(worksheet.Cell(2, firstDataColumn).GetDateTime());
                il_players.Add(value);
            }
        }

        //Gets first damage for the season
        //Different Range for Party Of One since ironman takes 5 rows for that sheet
        IXLRange fdRange = worksheet.Range(7, firstDataColumn, 7, lastDataColumn);
        IXLRange POOfdRange = worksheet.Range(7, firstDataColumn, 7, lastDataColumn);
        if (worksheet.Name.Equals("Party of One"))
        {
            foreach (IXLCell cell in POOfdRange.CellsUsed())
            {
                string value = cell.GetString();
                gs_firstdamage[value] += 1;

                //Add to first damage list
                dl_rounds.Add(worksheet.Name);
                dl_seasons.Add(worksheet.Cell(1, firstDataColumn).GetString());
                dl_dates.Add(worksheet.Cell(2, firstDataColumn).GetDateTime());
                dl_players.Add(value);
            }
        }
        else
        {
            foreach (IXLCell cell in fdRange.CellsUsed())
            {
                string value = cell.GetString();
                gs_firstdamage[value] += 1;

                //Add to first damage list
                dl_rounds.Add(worksheet.Name);
                dl_seasons.Add(worksheet.Cell(1, firstDataColumn).GetString());
                dl_dates.Add(worksheet.Cell(2, firstDataColumn).GetDateTime());
                dl_players.Add(value);
            }
        }

        //Loops through all the kiler cells
        IXLRange killerRange = worksheet.Range(firstDataRow, lastDataColumn, firstDataRow + (seasonSize - 1), lastDataColumn);
        foreach (IXLCell cell in killerRange.Cells())
        {
            string value = cell.GetString();

            //Checks if killer is PvE or Player
            if (gs_kills.ContainsKey(value))
            {
                //Sets values for the kill list
                kl_rounds.Add(worksheet.Name);
                kl_seasons.Add(worksheet.Cell(1, firstDataColumn).GetString());
                kl_dates.Add(worksheet.Cell(2, firstDataColumn).GetDateTime());
                kl_victims.Add(cell.CellLeft().CellLeft().GetString());
                kl_methods.Add(cell.CellLeft().GetString());
                kl_killers.Add(value);
                
                //Adds +1 kill on global stats
                gs_kills[value] += 1;

                //Updates KDR & KPR of player
                //If infinite sets it to the amount of kills they have
                if (gs_deaths[value].Equals(0))
                {
                    gs_kdr[value] = Convert.ToDouble(gs_kills[value]);
                }
                else
                {
                    gs_kdr[value] = Convert.ToDouble(gs_kills[value]) / Convert.ToDouble(gs_deaths[value]);
                }
                gs_kpr[value] = Convert.ToDouble(gs_kills[value]) / Convert.ToDouble(gs_seasonsplayed[value]);

                //Figures out the killboard of the season
                if (killboard.ContainsKey(value))
                {
                    killboard[value] += 1;
                }
                else
                {
                    killboard.Add(value, 1);
                }

                //Gets First Blood for the season
                if (first_blood == 0)
                {
                    gs_firstblood[value] += 1;
                    first_blood += 1;

                    //Add to first blood list
                    fbl_rounds.Add(worksheet.Name);
                    fbl_seasons.Add(worksheet.Cell(1, firstDataColumn).GetString());
                    fbl_dates.Add(worksheet.Cell(2, firstDataColumn).GetDateTime());
                    fbl_players.Add(value);
                }
            }
            else
            {
                //Adds +1 PvE Death for the player
                if (!value.Equals("Nothing"))
                {
                    gs_pve[worksheet.Cell(cell.WorksheetRow().RowNumber(), firstDataColumn).GetString()] += 1;
                }
            }
        }

        //Gets top frags for the season
        //Skips PolyCraft Egg Hunt since no one got kills in that
        if (!worksheet.Name.Equals("PolyCraft Egg Hunt"))
        {
            int topFragAmount = killboard.Values.Max();
            foreach (String killer in killboard.Keys)
            {
                if (killboard[killer] == topFragAmount)
                {
                    gs_topfrags[killer] += 1;

                    //Add to top frag list
                    tfl_rounds.Add(worksheet.Name);
                    tfl_seasons.Add(worksheet.Cell(1, firstDataColumn).GetString());
                    tfl_dates.Add(worksheet.Cell(2, firstDataColumn).GetDateTime());
                    tfl_players.Add(killer);
                }

                //Checks if the player that got kills beat their kill record
                if (kr_killrecord.ContainsKey(killer))
                {
                    if (killboard[killer] > kr_killrecord[killer])
                    {
                        kr_killrecord[killer] = killboard[killer];
                        kr_round[killer] = worksheet.Name;
                        kr_season[killer] = worksheet.Cell(1, firstDataColumn).GetString();
                        kr_date[killer] = worksheet.Cell(2, firstDataColumn).GetDateTime();
                    } else if (killboard[killer] == kr_killrecord[killer])
                    {
                        //If kill records are tied picks the first one that happened
                        if (kr_date[killer] > worksheet.Cell(2, firstDataColumn).GetDateTime())
                        {
                            kr_round[killer] = worksheet.Name;
                            kr_season[killer] = worksheet.Cell(1, firstDataColumn).GetString();
                            kr_date[killer] = worksheet.Cell(2, firstDataColumn).GetDateTime();
                        }
                    }
                } else
                {
                    //If first round with kills sets the kill record
                    kr_killrecord.Add(killer,killboard[killer]);
                    kr_round.Add(killer,worksheet.Name);
                    kr_season.Add(killer,worksheet.Cell(1, firstDataColumn).GetString());
                    kr_date.Add(killer,worksheet.Cell(2, firstDataColumn).GetDateTime());
                }
            }
        }

        //Get the teams for the season
        //Skips FFA seasons since no teams
        IXLRange teamRange = worksheet.Range(9, firstDataColumn, firstDataRow - 2, firstDataColumn);
        teamSize = teamRange.RowsUsed().Count();
        if (!worksheet.Cell(3, middleDataColumn).GetString().Equals("FFA"))
        {
            //Loops the Cells in the team list
            foreach (IXLCell cell in teamRange.CellsUsed())
            {
                string value = cell.GetString();

                //Adds the team to the list of the season
                seasonTeams.Add(value);

                //Adds the team info the to team list, skips if player is a solo
                if (value.Contains(","))
                {
                    tl_teams.Add(value);
                    tl_rounds.Add(worksheet.Name);
                    tl_seasons.Add(worksheet.Cell(1, firstDataColumn).GetString());
                    tl_dates.Add(worksheet.Cell(2, firstDataColumn).GetDateTime());
                }
            }

            //Get the team color and adds it to the team list, skips if player is a solo
            IXLRange teamColorsRange = worksheet.Range(9, lastDataColumn, 9 + (teamSize - 1), lastDataColumn);
            foreach (IXLCell cell in teamColorsRange.Cells())
            {
                string value = cell.GetString();

                if (cell.CellLeft().CellLeft().GetString().Contains(","))
                {
                tl_teamcolors.Add(value);
                }
            }
        }

        //Get the winners of the season
        //If Nothing is a regular season ending and gives the win to the last player on the list
        //Else is either a double kill win or no wins and is figured out to give the wins needed
        if (worksheet.Cell(firstDataRow + (seasonSize - 1), lastDataColumn).GetString().Equals("Nothing"))
        {
            //Gets the last season winner
            String seasonWinner = worksheet.Cell(firstDataRow + (seasonSize - 1), firstDataColumn).GetString();
            winnerCell = worksheet.Cell(firstDataRow + (seasonSize - 1), firstDataColumn);

            //If FFA no need to look for teams, else looks for the team
            if (worksheet.Cell(3, middleDataColumn).GetString().Equals("FFA"))
            {
                if (gs_wins.ContainsKey(seasonWinner))
                {
                    gs_wins[seasonWinner] += 1;

                    //Add to winner list
                    wl_rounds.Add(worksheet.Name);
                    wl_seasons.Add(worksheet.Cell(1, firstDataColumn).GetString());
                    wl_dates.Add(worksheet.Cell(2, firstDataColumn).GetDateTime());
                    wl_players.Add(seasonWinner);
                }
            }
            else
            {
                //Figures out the full team that won the season
                foreach (String team in seasonTeams)
                {
                    if (team.Contains(seasonWinner))
                    {
                        winningTeam = team;

                        //Splits the team string to get each player and gives them a win
                        String[] winners = team.Split(separator);
                        foreach (String winner in winners)
                        {
                            if (gs_wins.ContainsKey(winner))
                            {
                                gs_wins[winner] += 1;

                                //Add to winner list
                                wl_rounds.Add(worksheet.Name);
                                wl_seasons.Add(worksheet.Cell(1, firstDataColumn).GetString());
                                wl_dates.Add(worksheet.Cell(2, firstDataColumn).GetDateTime());
                                wl_players.Add(winner);
                            }
                        }
                    }
                }
            }
        }
        else
        {
            // TO FIX LATER
            String seasonWinner = worksheet.Cell(firstDataRow + (seasonSize - 1), firstDataColumn).GetString();
            winnerCell = worksheet.Cell(firstDataRow + (seasonSize - 1), firstDataColumn);
            if (worksheet.Cell(3, middleDataColumn).GetString().Equals("FFA"))
            {
                if (gs_wins.ContainsKey(seasonWinner))
                {
                    gs_wins[seasonWinner] += 1;

                    //Add to winner list
                    wl_rounds.Add(worksheet.Name);
                    wl_seasons.Add(worksheet.Cell(1, firstDataColumn).GetString());
                    wl_dates.Add(worksheet.Cell(2, firstDataColumn).GetDateTime());
                    wl_players.Add(seasonWinner);
                }
            }
            else
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

                                //Add to winner list
                                wl_rounds.Add(worksheet.Name);
                                wl_seasons.Add(worksheet.Cell(1, firstDataColumn).GetString());
                                wl_dates.Add(worksheet.Cell(2, firstDataColumn).GetDateTime());
                                wl_players.Add(winner);
                            }
                        }
                    }
                }
            }
        }

        //Get the runner ups of the season
        //If FFA it has to be the player above
        //Else figures out the next team after the winners
        if (worksheet.Cell(3, middleDataColumn).GetString().Equals("FFA"))
        {
            if (gs_runnerup.ContainsKey(winnerCell.CellAbove().GetString()))
            {
                gs_runnerup[winnerCell.CellAbove().GetString()] += 1;

                //Add to runner up list
                rul_rounds.Add(worksheet.Name);
                rul_seasons.Add(worksheet.Cell(1, firstDataColumn).GetString());
                rul_dates.Add(worksheet.Cell(2, firstDataColumn).GetDateTime());
                rul_players.Add(winnerCell.CellAbove().GetString());
            }
        }
        else
        {
            //Looks for the next cell that contains someone not on the winning team
            while (winningTeam.Contains(winnerCell.CellAbove().GetString()))
            {
                winnerCell = winnerCell.CellAbove();
            }

            //Figures out the full team of runner ups
            String seasonRunnerUp = winnerCell.CellAbove().GetString();
            foreach (String team in seasonTeams)
            {
                if (team.Contains(seasonRunnerUp))
                {
                    //Splits the team string to get each player and gives them a runner up
                    String[] runnerups = team.Split(separator);
                    foreach (String runner_up in runnerups)
                    {
                        if (gs_runnerup.ContainsKey(runner_up))
                        {
                            gs_runnerup[runner_up] += 1;

                            //Add to runner up list
                            rul_rounds.Add(worksheet.Name);
                            rul_seasons.Add(worksheet.Cell(1, firstDataColumn).GetString());
                            rul_dates.Add(worksheet.Cell(2, firstDataColumn).GetDateTime());
                            rul_players.Add(runner_up);
                        }
                    }
                }
            }
        }
    }

    rl_rostersizes.Add(roundRoster.Count);
}

//Takes all the lists and adds them to new docs to save
var roundlist = new XLWorkbook();
var statscompiled = new XLWorkbook();
var rrdebut = new XLWorkbook();
var globalstats = new XLWorkbook();

//Making Round List Page
var round_list = roundlist.AddWorksheet("Round List");
round_list.Column("D").Style.NumberFormat.Format = "dd mmm, yyyy";
round_list.Column("A").Width = 34;
round_list.Column("B").Width = 6;
round_list.Column("C").Width = 6;
round_list.Column("D").Width = 20;
round_list.Cell("A1").InsertData(rl_rounds);
round_list.Cell("B1").InsertData(rl_seasons);
round_list.Cell("C1").InsertData(rl_rostersizes);
round_list.Cell("D1").InsertData(rl_rounddebuts);
round_list.Sort(4);

//Making All Rosters Page
var allrosters = roundlist.AddWorksheet("All Rosters");
allrosters.Column("C").Style.NumberFormat.Format = "mm/dd/yyyy";
allrosters.Column("A").Width = 34;
allrosters.Column("B").Width = 6;
allrosters.Column("C").Width = 20;
allrosters.Cell("A1").InsertData(ar_rounds);
allrosters.Cell("B1").InsertData(ar_seasons);
allrosters.Cell("C1").InsertData(ar_dates);
allrosters.Cell("D1").InsertData(ar_rosters);
allrosters.Sort(3);

//Making NR All Rosters Page
var allrosters_nr = roundlist.AddWorksheet("All Rosters (NR)");
allrosters_nr.Column("C").Style.NumberFormat.Format = "mm/dd/yyyy";
allrosters_nr.Column("A").Width = 34;
allrosters_nr.Column("B").Width = 6;
allrosters_nr.Column("C").Width = 20;
allrosters_nr.Cell("A1").InsertData(nr_rounds);
allrosters_nr.Cell("B1").InsertData(nr_seasons);
allrosters_nr.Cell("C1").InsertData(nr_dates);
allrosters_nr.Cell("D1").InsertData(nr_rosters);
allrosters_nr.Sort(3);

//Making Kills Page
var allkills = statscompiled.AddWorksheet("All Kills");
allkills.Column("C").Style.NumberFormat.Format = "mm/dd/yyyy";
allkills.Column("A").Width = 34;
allkills.Column("B").Width = 6;
allkills.Column("C").Width = 20;
allkills.Cell("A1").InsertData(kl_rounds);
allkills.Cell("B1").InsertData(kl_seasons);
allkills.Cell("C1").InsertData(kl_dates);
allkills.Cell("D1").InsertData(kl_victims);
allkills.Cell("E1").InsertData(kl_methods);
allkills.Cell("F1").InsertData(kl_killers);
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
allteams.Cell("D1").InsertData(tl_teams);
allteams.Cell("E1").InsertData(tl_teamcolors);
allteams.Sort(3);

//First Damage list
var firstdamage = statscompiled.AddWorksheet("First Damage");
firstdamage.Column("C").Style.NumberFormat.Format = "mm/dd/yyyy";
firstdamage.Column("A").Width = 34;
firstdamage.Column("B").Width = 6;
firstdamage.Column("C").Width = 20;
firstdamage.Cell("A1").InsertData(dl_rounds);
firstdamage.Cell("B1").InsertData(dl_seasons);
firstdamage.Cell("C1").InsertData(dl_dates);
firstdamage.Cell("D1").InsertData(dl_players);
firstdamage.Sort(3);

//Ironman list
var ironman = statscompiled.AddWorksheet("Ironman");
ironman.Column("C").Style.NumberFormat.Format = "mm/dd/yyyy";
ironman.Column("A").Width = 34;
ironman.Column("B").Width = 6;
ironman.Column("C").Width = 20;
ironman.Cell("A1").InsertData(il_rounds);
ironman.Cell("B1").InsertData(il_seasons);
ironman.Cell("C1").InsertData(il_dates);
ironman.Cell("D1").InsertData(il_players);
ironman.Sort(3);

//First Death list
var firstdeath = statscompiled.AddWorksheet("First Death");
firstdeath.Column("C").Style.NumberFormat.Format = "mm/dd/yyyy";
firstdeath.Column("A").Width = 34;
firstdeath.Column("B").Width = 6;
firstdeath.Column("C").Width = 20;
firstdeath.Cell("A1").InsertData(fdl_rounds);
firstdeath.Cell("B1").InsertData(fdl_seasons);
firstdeath.Cell("C1").InsertData(fdl_dates);
firstdeath.Cell("D1").InsertData(fdl_players);
firstdeath.Sort(3);

//First Blood list
var firstblood = statscompiled.AddWorksheet("First Blood");
firstblood.Column("C").Style.NumberFormat.Format = "mm/dd/yyyy";
firstblood.Column("A").Width = 34;
firstblood.Column("B").Width = 6;
firstblood.Column("C").Width = 20;
firstblood.Cell("A1").InsertData(fbl_rounds);
firstblood.Cell("B1").InsertData(fbl_seasons);
firstblood.Cell("C1").InsertData(fbl_dates);
firstblood.Cell("D1").InsertData(fbl_players);
firstblood.Sort(3);

//Most Kills list
var topfrags = statscompiled.AddWorksheet("Top Frags");
topfrags.Column("C").Style.NumberFormat.Format = "mm/dd/yyyy";
topfrags.Column("A").Width = 34;
topfrags.Column("B").Width = 6;
topfrags.Column("C").Width = 20;
topfrags.Cell("A1").InsertData(tfl_rounds);
topfrags.Cell("B1").InsertData(tfl_seasons);
topfrags.Cell("C1").InsertData(tfl_dates);
topfrags.Cell("D1").InsertData(tfl_players);
topfrags.Sort(3);

//Runner Up list
var runnerup = statscompiled.AddWorksheet("Runner Ups");
runnerup.Column("C").Style.NumberFormat.Format = "mm/dd/yyyy";
runnerup.Column("A").Width = 34;
runnerup.Column("B").Width = 6;
runnerup.Column("C").Width = 20;
runnerup.Cell("A1").InsertData(rul_rounds);
runnerup.Cell("B1").InsertData(rul_seasons);
runnerup.Cell("C1").InsertData(rul_dates);
runnerup.Cell("D1").InsertData(rul_players);
runnerup.Sort(3);

//Win list
var win = statscompiled.AddWorksheet("Wins");
win.Column("C").Style.NumberFormat.Format = "mm/dd/yyyy";
win.Column("A").Width = 34;
win.Column("B").Width = 6;
win.Column("C").Width = 20;
win.Cell("A1").InsertData(wl_rounds);
win.Cell("B1").InsertData(wl_seasons);
win.Cell("C1").InsertData(wl_dates);
win.Cell("D1").InsertData(wl_players);
win.Sort(3);

//RR Debuts
var rr_debut = rrdebut.AddWorksheet("RR Debuts");
rr_debut.Column("C").Style.NumberFormat.Format = "mm/dd/yyyy";
rr_debut.Column("A").Width = 34;
rr_debut.Column("B").Width = 6;
rr_debut.Column("C").Width = 20;
rr_debut.Cell("A1").InsertData(dl_round.Values);
rr_debut.Cell("B1").InsertData(dl_season.Values);
rr_debut.Cell("C1").InsertData(dl_date.Values);
rr_debut.Cell("D1").InsertData(dl_round.Keys);
rr_debut.Sort(3);

//RR Debuts (No NR)
var rr_debut_nr = rrdebut.AddWorksheet("RR Debuts (No NR)");
rr_debut_nr.Column("C").Style.NumberFormat.Format = "mm/dd/yyyy";
rr_debut_nr.Column("A").Width = 34;
rr_debut_nr.Column("B").Width = 6;
rr_debut_nr.Column("C").Width = 20;
rr_debut_nr.Cell("A1").InsertData(nrdl_round.Values);
rr_debut_nr.Cell("B1").InsertData(nrdl_season.Values);
rr_debut_nr.Cell("C1").InsertData(nrdl_date.Values);
rr_debut_nr.Cell("D1").InsertData(nrdl_round.Keys);
rr_debut_nr.Sort(3);

//Global Stats
var global_stats = globalstats.AddWorksheet("Global Stats");
global_stats.Column("A").Width = 20;
global_stats.Cell("A1").InsertData(gs_seasonsplayed.Keys);
global_stats.Cell("B1").InsertData(gs_seasonsplayed.Values);
global_stats.Cell("C1").InsertData(gs_wins.Values);
global_stats.Cell("D1").InsertData(gs_alive.Values);
global_stats.Cell("E1").InsertData(gs_runnerup.Values);
global_stats.Cell("F1").InsertData(gs_kills.Values);
global_stats.Cell("G1").InsertData(gs_topfrags.Values);
global_stats.Cell("H1").InsertData(gs_pve.Values);
global_stats.Cell("I1").InsertData(gs_firstblood.Values);
global_stats.Cell("J1").InsertData(gs_firstdeath.Values);
global_stats.Cell("K1").InsertData(gs_ironman.Values);
global_stats.Cell("L1").InsertData(gs_firstdamage.Values);
global_stats.Cell("M1").InsertData(gs_deaths.Values);
global_stats.Cell("N1").InsertData(gs_totaluniques.Values);
global_stats.Cell("O1").InsertData(gs_kdr.Values);
global_stats.Cell("P1").InsertData(gs_kpr.Values);
global_stats.Sort(1);

//Kill Records
var kill_records = globalstats.AddWorksheet("Kill Records");
kill_records.Column("A").Width = 20;
kill_records.Column("B").Width = 6;
kill_records.Column("C").Width = 34;
kill_records.Column("D").Width = 6;
kill_records.Column("E").Width = 20;
kill_records.Cell("A1").InsertData(kr_killrecord.Keys);
kill_records.Cell("B1").InsertData(kr_killrecord.Values);
kill_records.Cell("C1").InsertData(kr_round.Values);
kill_records.Cell("D1").InsertData(kr_season.Values);
kill_records.Cell("E1").InsertData(kr_date.Values);
kill_records.Sort(1);

//Saves the new docs
roundlist.SaveAs("..\\..\\..\\Round_List.xlsx");
statscompiled.SaveAs("..\\..\\..\\Stats_Compiled.xlsx");
globalstats.SaveAs("..\\..\\..\\Global_Stats.xlsx");
rrdebut.SaveAs("..\\..\\..\\RR_Debuts.xlsx");
Console.WriteLine("Stats are now compiled!");