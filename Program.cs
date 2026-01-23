using System.Reflection.PortableExecutable;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;

string filePath = "..\\..\\..\\Global RR Stats Community Document.xlsx";
//string filePath = "..\\..\\..\\Non-Reddit Stats Community Document.xlsx";
//string filePath = "..\\..\\..\\Global Live Round Stats Community Document.xlsx";

if (!File.Exists(filePath))
{
    Console.ForegroundColor = ConsoleColor.Red;
    Console.WriteLine($"Error: File not found at {filePath}");
    return;
}

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
//Variables for the pve death list
List<String> pvel_rounds = new List<String>();
List<String> pvel_seasons = new List<String>();
List<DateTime> pvel_dates = new List<DateTime>();
List<String> pvel_players = new List<String>();
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
//Variables for the alive list
List<String> al_rounds = new List<String>();
List<String> al_seasons = new List<String>();
List<DateTime> al_dates = new List<DateTime>();
List<String> al_players = new List<String>();
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
//Unique PvE Deaths
Dictionary<String, int> unique_pve_deaths = new Dictionary<String, int>();
//Variables for the reddit posts
List<String> rp_winners = new List<String>();
List<String> rp_runnerups = new List<String>();
List<String> rp_mostkills = new List<String>();
List<String> rp_mostkillsteam = new List<String>();
List<String> rp_firstdamage = new List<String>();
List<String> rp_ironman = new List<String>();
List<String> rp_firstblood = new List<String>();
List<String> rp_firstdeath = new List<String>();
Dictionary<String, int> rp_kills = new Dictionary<String, int>();
Dictionary<String, String> rp_kills_list = new Dictionary<String, String>();
Dictionary<String, int> rp_pvedeaths = new Dictionary<String, int>();
Dictionary<String, String> rp_pvedeaths_list = new Dictionary<String, String>();
List<String> rp_participations = new List<String>();
List<String> rp_debutants = new List<String>();

//Goes through every stats tabs on the doc
using var workbook = new XLWorkbook(filePath);
for (int sheet = 8; sheet <= workbook.Worksheets.Count; sheet++)
{
    List<String> roundRoster = new List<String>();
    //Collecting the Name, the total amount of seasons & the date of S1 for the round
    var worksheet = workbook.Worksheet(sheet);
    String round_name = worksheet.Name;
    if (round_name.Contains("Sheet"))
    {
        round_name = "???";
    }
    rl_rounds.Add(round_name);
    rl_seasons.Add((worksheet.Columns().Count() - 1) / 3);
    rl_rounddebuts.Add(worksheet.Cell(2, 2).GetDateTime());
    Console.ForegroundColor = ConsoleColor.Green;
    Console.WriteLine("Working on " + round_name);
    Console.WriteLine(((worksheet.Columns().Count() - 1) / 3).ToString() + " Seasons!");

    //Goes through every season on the round sheet
    for (int season = 1; season <= (worksheet.Columns().Count() - 1); season += 3)
    {
        //Gets the row for the start of the death log & sets cell locations relative to the season
        IXLCell victimsStart = worksheet.Search("Kill List (include alive)").First();
        var rangeUsed = worksheet.RangeUsed();
        int firstDataRow = victimsStart.WorksheetRow().RowNumber() + 1;
        int lastDataRow = rangeUsed.LastRowUsed().RowNumber();
        int firstDataColumn = season + 1;
        int middleDataColumn = season + 2;
        int lastDataColumn = season + 3;
        String season_debutant = "";

        //Sets round named to be changed for crossovers and ??? to not be called Sheet
        round_name = worksheet.Name;
        if (round_name.Contains("Sheet"))
        {
            round_name = "???";
        }
        String season_number = worksheet.Cell(1, firstDataColumn).GetString();
        DateTime season_date = worksheet.Cell(2, firstDataColumn).GetDateTime();

        //Sets name of round to contain both names for crossovers
        if (round_name.Equals("Phobia") && season_number.Equals("20"))
        {
            round_name = "WMC x Phobia";
            season_number = "30/20";
        }
        if (round_name.Equals("Scattershot") && season_number.Equals("6"))
        {
            round_name = "The Melon Blooded x Scattershot";
            season_number = "40/6";
        }
        if (round_name.Equals("Cinema") && season_number.Equals("16b"))
        {
            round_name = "Phobia x Cinema";
            season_number = "28/16b";
        }

        //Only count one of the crossover round towards itself only
        if (round_name.Equals("WMC") && season_number.Equals("30") ||
            round_name.Equals("The Melon Blooded") && season_number.Equals("40") ||
            round_name.Equals("Phobia") && season_number.Equals("28"))
        {
            //Adds the unique people from the crossover round into the round
            IXLRange victimRange = worksheet.Range(firstDataRow, firstDataColumn, lastDataRow, firstDataColumn);
            foreach (IXLCell cell in victimRange.CellsUsed())
            {
                string value = cell.GetString();

                if (!roundRoster.Contains(value))
                {
                    roundRoster.Add(value);

                    gs_totaluniques[value] += 1;

                    season_debutant += value + ", ";
                }
            }

            //Formats the debutants for reddit posts
            if (season_debutant.Length > 0)
            {
                season_debutant = season_debutant.Remove(season_debutant.Length - 2);
                rp_debutants.Add("**S" + season_number + " (" + (season_debutant.Count(c => c == ',') + 1) + "):** " + season_debutant + Environment.NewLine);
            }
            else
            {
                rp_debutants.Add("**S" + season_number + " (" + season_debutant.Count(c => c == ',') + "):** " + Environment.NewLine);
            }
        }
        else
        {
            //Sets variables for stats logic
            List<String> seasonRoster = new List<String>();
            List<String> seasonDebutant = new List<String>();
            List<String> seasonTeams = new List<String>();
            Dictionary<String, int> killboard = new Dictionary<String, int>();
            IXLCell winnerCell = worksheet.Cell(1, 1);
            IXLCell winnerCell2 = worksheet.Cell(1, 1);
            IXLCell lastAliveCell = worksheet.Cell(1, 1);
            String winningTeam = "";
            String winningTeam2 = "";
            char separator = ',';
            int seasonSize = 0;
            int teamSize = 0;
            int first_blood = 0;
            int double_kill_ending = 0;
            int dragon_win = 0;
            int dragon_rush_ru = 0;
            int double_kill_runnerup = 0;

            //Get seasons data for the all rosters list
            ar_rounds.Add(round_name);
            ar_seasons.Add(season_number);
            ar_dates.Add(season_date);

            //Get seasons data for all rounds except for the non-reddit releases
            if (!worksheet.Cell(1, lastDataColumn).GetString().Equals("NR"))
            {
                nr_rounds.Add(round_name);
                nr_seasons.Add(season_number);
                nr_dates.Add(season_date);
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
                        tl_rounds.Add(round_name);
                        tl_seasons.Add(season_number);
                        tl_dates.Add(season_date);
                    }
                }

                //Get the team color and adds it to the team list, skips if player is a solo
                IXLRange teamColorsRange = worksheet.Range(9, lastDataColumn, 9 + (teamSize - 1), lastDataColumn);
                foreach (IXLCell cell in teamColorsRange.Cells())
                {
                    string value = cell.GetString();

                    if (cell.CellLeft().CellLeft().GetString().Contains(","))
                    {
                        if (value.Equals(""))
                        {
                            tl_teamcolors.Add("N/A");
                        }
                        else
                        {
                            tl_teamcolors.Add(value);
                        }
                    }
                }
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
                    if (season_date < dl_date[value])
                    {
                        dl_round[value] = round_name;
                        dl_season[value] = season_number;
                        dl_date[value] = season_date;
                    }
                }
                else
                {
                    dl_round.Add(value, round_name);
                    dl_season.Add(value, season_number);
                    dl_date.Add(value, season_date);

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
                        if (season_date < nrdl_date[value])
                        {
                            nrdl_round[value] = round_name;
                            nrdl_season[value] = season_number;
                            nrdl_date[value] = season_date;
                        }
                    }
                }
                else
                {
                    if (!worksheet.Cell(1, lastDataColumn).GetString().Equals("NR"))
                    {
                        nrdl_round.Add(value, round_name);
                        nrdl_season.Add(value, season_number);
                        nrdl_date.Add(value, season_date);
                    }
                }

                //Add Error Messages for suicides.
                if (worksheet.Cell(cell.WorksheetRow().RowNumber(), lastDataColumn).GetString().Equals(value))
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine("ERROR: " + value + " suicided! " + round_name + " " + season_number);
                }

                //If the players didn't die gets +1 alive on the global stats, else +1 death
                if (worksheet.Cell(cell.WorksheetRow().RowNumber(), lastDataColumn).GetString().Equals("Nothing"))
                {
                    gs_alive[value] += 1;

                    //Add to ironman list
                    al_rounds.Add(round_name);
                    al_seasons.Add(season_number);
                    al_dates.Add(season_date);
                    al_players.Add(value);
                }
                else
                {
                    gs_deaths[value] += 1;
                }

                //Makes roster for the season, skips players who show up twice with respawns gamemodes
                //Also adds +1 seasons played for the global stats
                if (!seasonRoster.Contains(value))
                {
                    seasonRoster.Add(value);
                    gs_seasonsplayed[value] += 1;
                }

                //Makes roster for the round, adds new players
                //Also adds +1 unique round for the global stats
                if (!roundRoster.Contains(value))
                {
                    roundRoster.Add(value);
                    seasonDebutant.Add(value);
                    gs_totaluniques[value] += 1;
                }
            }
            //Formats the debutants for reddit posts
            seasonDebutant.Sort();
            foreach (String debutant in seasonDebutant)
            {
                season_debutant += debutant + ", ";
            }
            if (season_debutant.Length > 0)
            {
                season_debutant = season_debutant.Remove(season_debutant.Length - 2);
                rp_debutants.Add("**S" + worksheet.Cell(1, firstDataColumn).GetString() + " (" + (season_debutant.Count(c => c == ',') + 1) + "):** " + season_debutant + Environment.NewLine);
            }
            else
            {
                rp_debutants.Add("**S" + worksheet.Cell(1, firstDataColumn).GetString() + " (" + season_debutant.Count(c => c == ',') + "):** " + Environment.NewLine);
            }

            //Adds rosters to a list for the sheet, skips non-reddit for the alternate page
            seasonRoster.Sort();
            ar_rosters.Add(seasonRoster);
            if (!worksheet.Cell(1, lastDataColumn).GetString().Equals("NR"))
            {
                nr_rosters.Add(seasonRoster);
            }

            if (!worksheet.Cell(3, middleDataColumn).GetString().Equals("FFA"))
            {
                //Verify if player is misspelled or missing in teams
                int playercheck = 0;
                foreach (String player in seasonRoster)
                {
                    foreach (String team in seasonTeams)
                    {
                        if (team.Contains(player))
                        {
                            playercheck += 1;
                        }
                    }

                    if (playercheck == 0)
                    {
                        Console.ForegroundColor = ConsoleColor.Red;
                        Console.WriteLine("ERROR: Player " + player + " missing in teams! " + round_name + " " + season_number);
                    }

                    if (playercheck > 1)
                    {
                        Console.ForegroundColor = ConsoleColor.Red;
                        Console.WriteLine("ERROR: Player " + player + " duplicate in teams! " + round_name + " " + season_number);
                    }

                    playercheck = 0;
                }

                //Verify if player is misspelled or missing in victims
                int teamcheck = 0;
                foreach (String team in seasonTeams)
                {
                    String[] teamplayers = team.Split(separator);

                    foreach (String teamplayer in teamplayers)
                    {
                        if (seasonRoster.Contains(teamplayer))
                        {
                            teamcheck += 1;
                        }

                        if (teamcheck == 0)
                        {
                            Console.ForegroundColor = ConsoleColor.Red;
                            Console.WriteLine("ERROR: Player " + teamplayer + " missing in victims! " + round_name + " " + season_number);
                        }

                        teamcheck = 0;
                    }
                }
            }

            //Figures out if there is a double kill for first death, otherwise add +1 to the first death
            if (worksheet.Cell(firstDataRow, firstDataColumn).GetString().Equals(worksheet.Cell(firstDataRow + 1, lastDataColumn).GetString())
                && worksheet.Cell(firstDataRow + 1, firstDataColumn).GetString().Equals(worksheet.Cell(firstDataRow, lastDataColumn).GetString()))
            {
                gs_firstdeath[worksheet.Cell(firstDataRow, firstDataColumn).GetString()] += 1;
                gs_firstdeath[worksheet.Cell(firstDataRow + 1, firstDataColumn).GetString()] += 1;

                rp_firstdeath.Add("**S" + worksheet.Cell(1, firstDataColumn).GetString() + ":** " + worksheet.Cell(firstDataRow, firstDataColumn).GetString() + " & " + worksheet.Cell(firstDataRow + 1, firstDataColumn).GetString() + " (Double Kill)" + Environment.NewLine);

                //Add to first death list
                fdl_rounds.Add(round_name);
                fdl_rounds.Add(round_name);
                fdl_seasons.Add(season_number);
                fdl_seasons.Add(season_number);
                fdl_dates.Add(season_date);
                fdl_dates.Add(season_date);
                fdl_players.Add(worksheet.Cell(firstDataRow, firstDataColumn).GetString());
                fdl_players.Add(worksheet.Cell(firstDataRow + 1, firstDataColumn).GetString());
            }
            else
            {
                gs_firstdeath[worksheet.Cell(firstDataRow, firstDataColumn).GetString()] += 1;

                //Add to first death list
                fdl_rounds.Add(round_name);
                fdl_seasons.Add(season_number);
                fdl_dates.Add(season_date);
                fdl_players.Add(worksheet.Cell(firstDataRow, firstDataColumn).GetString());

                //Double the stats for round exception
                if (round_name.Equals("Game Changer")
                    && season_number.Equals("5"))
                {
                    gs_firstdeath[worksheet.Cell(firstDataRow, firstDataColumn).CellBelow().GetString()] += 1;
                    rp_firstdeath.Add("**S" + worksheet.Cell(1, firstDataColumn).GetString() + ":** " + worksheet.Cell(firstDataRow, firstDataColumn).GetString() + " & " + worksheet.Cell(firstDataRow, firstDataColumn).CellBelow().GetString() + " (" + worksheet.Cell(firstDataRow, firstDataColumn).CellRight().CellRight().GetString() + ")" + Environment.NewLine);

                    //Add to first death list
                    fdl_rounds.Add(round_name);
                    fdl_seasons.Add(season_number);
                    fdl_dates.Add(season_date);
                    fdl_players.Add(worksheet.Cell(firstDataRow, firstDataColumn).CellBelow().GetString());
                }
                else
                {
                    if (worksheet.Cell(firstDataRow, firstDataColumn).CellRight().CellRight().GetString().Equals(""))
                    {
                        String pvedeath = getPvEDeath(worksheet.Cell(firstDataRow, firstDataColumn).CellRight().CellRight());
                        rp_firstdeath.Add("**S" + worksheet.Cell(1, firstDataColumn).GetString() + ":** " + worksheet.Cell(firstDataRow, firstDataColumn).GetString() + " (" + pvedeath + ")" + Environment.NewLine);

                    }
                    else
                    {
                        rp_firstdeath.Add("**S" + worksheet.Cell(1, firstDataColumn).GetString() + ":** " + worksheet.Cell(firstDataRow, firstDataColumn).GetString() + " (" + worksheet.Cell(firstDataRow, firstDataColumn).CellRight().CellRight().GetString() + ")" + Environment.NewLine);
                    }
                }
            }

            //Gets ironman for the season
            //Different Range for Party Of One since ironman takes 5 rows for that sheet
            IXLRange ironmanRange = worksheet.Range(5, firstDataColumn, 5, lastDataColumn);
            IXLRange POOironmanRange = worksheet.Range(5, firstDataColumn, 9, lastDataColumn);
            if (round_name.Equals("Party of One"))
            {
                foreach (IXLCell cell in POOironmanRange.CellsUsed())
                {
                    string value = cell.GetString();
                    gs_ironman[value] += 1;

                    //Add to ironman list
                    il_rounds.Add(round_name);
                    il_seasons.Add(season_number);
                    il_dates.Add(season_date);
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
                    il_rounds.Add(round_name);
                    il_seasons.Add(season_number);
                    il_dates.Add(season_date);
                    il_players.Add(value);
                }
            }

            //Gets first damage for the season
            //Different Range for Party Of One since ironman takes 5 rows for that sheet
            IXLRange fdRange = worksheet.Range(7, firstDataColumn, 7, lastDataColumn);
            IXLRange POOfdRange = worksheet.Range(11, firstDataColumn, 11, lastDataColumn);
            if (round_name.Equals("Party of One"))
            {
                foreach (IXLCell cell in POOfdRange.CellsUsed())
                {
                    string value = cell.GetString();
                    gs_firstdamage[value] += 1;

                    //Add to first damage list
                    dl_rounds.Add(round_name);
                    dl_seasons.Add(season_number);
                    dl_dates.Add(season_date);
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
                    dl_rounds.Add(round_name);
                    dl_seasons.Add(season_number);
                    dl_dates.Add(season_date);
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
                    kl_rounds.Add(round_name);
                    kl_seasons.Add(season_number);
                    kl_dates.Add(season_date);
                    kl_victims.Add(cell.CellLeft().CellLeft().GetString());
                    kl_methods.Add(cell.CellLeft().GetString());
                    kl_killers.Add(value);

                    //Adds +1 kill on global stats
                    gs_kills[value] += 1;

                    if (rp_kills.ContainsKey(value))
                    {
                        rp_kills[value] += 1;
                        rp_kills_list[value] = rp_kills_list[value] + cell.CellLeft().CellLeft().GetString() + " (S" + worksheet.Cell(1, firstDataColumn).GetString() + "), ";
                    }
                    else
                    {
                        rp_kills.Add(value, 1);
                        rp_kills_list.Add(value, cell.CellLeft().CellLeft().GetString() + " (S" + worksheet.Cell(1, firstDataColumn).GetString() + "), ");
                    }

                    //Figures out the killboard of the season
                    if (killboard.ContainsKey(value))
                    {
                        killboard[value] += 1;
                    }
                    else
                    {
                        killboard.Add(value, 1);
                    }

                    //Check if there was a double kill for first blood, otherwise gives it to the first player found
                    if (first_blood == 0)
                    {
                        if (value.Equals(cell.CellBelow().CellLeft().CellLeft().GetString())
                            && cell.CellBelow().GetString().Equals(cell.CellLeft().CellLeft().GetString()))
                        {
                            gs_firstblood[value] += 1;
                            gs_firstblood[cell.CellBelow().GetString()] += 1;
                            first_blood += 2;

                            rp_firstblood.Add("**S" + worksheet.Cell(1, firstDataColumn).GetString() + ":** " + value + " & " + cell.CellBelow().GetString() + " (Double Kill)" + Environment.NewLine);

                            //Add to first blood list
                            fbl_rounds.Add(round_name);
                            fbl_rounds.Add(round_name);
                            fbl_seasons.Add(season_number);
                            fbl_seasons.Add(season_number);
                            fbl_dates.Add(season_date);
                            fbl_dates.Add(season_date);
                            fbl_players.Add(value);
                            fbl_players.Add(cell.CellBelow().GetString());
                        }
                        else
                        {
                            gs_firstblood[value] += 1;
                            first_blood += 1;

                            //Add to first blood list
                            fbl_rounds.Add(round_name);
                            fbl_seasons.Add(season_number);
                            fbl_dates.Add(season_date);
                            fbl_players.Add(value);

                            //Double the stats for round exception
                            if (round_name.Equals("Game Changer")
                                && season_number.Equals("5"))
                            {
                                gs_firstblood[cell.CellBelow().GetString()] += 1;
                                rp_firstblood.Add("**S" + worksheet.Cell(1, firstDataColumn).GetString() + ":** " + value + " & " + cell.CellBelow().GetString() + " (" + cell.CellLeft().CellLeft().GetString() + " & " + cell.CellBelow().CellLeft().CellLeft().GetString() + ")" + Environment.NewLine);

                                //Add to first blood list
                                fbl_rounds.Add(round_name);
                                fbl_seasons.Add(season_number);
                                fbl_dates.Add(season_date);
                                fbl_players.Add(cell.CellBelow().GetString());
                            }
                            else
                            {
                                rp_firstblood.Add("**S" + worksheet.Cell(1, firstDataColumn).GetString() + ":** " + value + " (" + cell.CellLeft().CellLeft().GetString() + ")" + Environment.NewLine);
                            }
                        }
                    }
                }
                else
                {
                    //Adds +1 PvE Death for the player
                    if (!value.Equals("Nothing"))
                    {
                        gs_pve[worksheet.Cell(cell.WorksheetRow().RowNumber(), firstDataColumn).GetString()] += 1;

                        //Add to pve list
                        pvel_rounds.Add(round_name);
                        pvel_seasons.Add(season_number);
                        pvel_dates.Add(season_date);
                        pvel_players.Add(worksheet.Cell(cell.WorksheetRow().RowNumber(), firstDataColumn).GetString());

                        //Filters all the unique pve deaths
                        if (value.Equals(""))
                        {
                            String pvedeath = getPvEDeath(cell);

                            //Sets values for the kill list
                            kl_rounds.Add(round_name);
                            kl_seasons.Add(season_number);
                            kl_dates.Add(season_date);
                            kl_victims.Add(cell.CellLeft().CellLeft().GetString());
                            kl_methods.Add(cell.CellLeft().GetString());
                            kl_killers.Add(pvedeath);

                            if (unique_pve_deaths.ContainsKey(pvedeath))
                            {
                                unique_pve_deaths[pvedeath] += 1;
                            }
                            else
                            {
                                unique_pve_deaths.Add(pvedeath, 1);
                            }

                            if (rp_pvedeaths.ContainsKey(pvedeath))
                            {
                                rp_pvedeaths[pvedeath] += 1;
                                rp_pvedeaths_list[pvedeath] = rp_pvedeaths_list[pvedeath] + cell.CellLeft().CellLeft().GetString() + " (S" + worksheet.Cell(1, firstDataColumn).GetString() + "), ";
                            }
                            else
                            {
                                rp_pvedeaths.Add(pvedeath, 1);
                                rp_pvedeaths_list.Add(pvedeath, cell.CellLeft().CellLeft().GetString() + " (S" + worksheet.Cell(1, firstDataColumn).GetString() + "), ");
                            }
                        }
                        else
                        {
                            //Sets values for the kill list
                            kl_rounds.Add(round_name);
                            kl_seasons.Add(season_number);
                            kl_dates.Add(season_date);
                            kl_victims.Add(cell.CellLeft().CellLeft().GetString());
                            kl_methods.Add(cell.CellLeft().GetString());
                            kl_killers.Add(value);

                            if (unique_pve_deaths.ContainsKey(value))
                            {
                                unique_pve_deaths[value] += 1;
                            }
                            else
                            {
                                unique_pve_deaths.Add(value, 1);
                            }

                            if (rp_pvedeaths.ContainsKey(value))
                            {
                                rp_pvedeaths[value] += 1;
                                rp_pvedeaths_list[value] = rp_pvedeaths_list[value] + cell.CellLeft().CellLeft().GetString() + " (S" + worksheet.Cell(1, firstDataColumn).GetString() + "), ";
                            }
                            else
                            {
                                rp_pvedeaths.Add(value, 1);
                                rp_pvedeaths_list.Add(value, cell.CellLeft().CellLeft().GetString() + " (S" + worksheet.Cell(1, firstDataColumn).GetString() + "), ");
                            }
                        }

                    }
                }
            }

            //Gets top frags for the season
            //Skips PolyCraft Egg Hunt since no one got kills in that
            if (killboard.Count > 0)
            {
                int topFragAmount = killboard.Values.Max();
                foreach (String killer in killboard.Keys)
                {
                    if (killboard[killer] == topFragAmount)
                    {
                        gs_topfrags[killer] += 1;

                        //Add to top frag list
                        tfl_rounds.Add(round_name);
                        tfl_seasons.Add(season_number);
                        tfl_dates.Add(season_date);
                        tfl_players.Add(killer);
                    }

                    //Checks if the player that got kills beat their kill record
                    if (kr_killrecord.ContainsKey(killer))
                    {
                        if (killboard[killer] > kr_killrecord[killer])
                        {
                            kr_killrecord[killer] = killboard[killer];
                            kr_round[killer] = round_name;
                            kr_season[killer] = season_number;
                            kr_date[killer] = season_date;
                        }
                        else if (killboard[killer] == kr_killrecord[killer])
                        {
                            //If kill records are tied picks the first one that happened
                            if (kr_date[killer] > season_date)
                            {
                                kr_round[killer] = round_name;
                                kr_season[killer] = season_number;
                                kr_date[killer] = season_date;
                            }
                        }
                    }
                    else
                    {
                        //If first round with kills sets the kill record
                        kr_killrecord.Add(killer, killboard[killer]);
                        kr_round.Add(killer, round_name);
                        kr_season.Add(killer, season_number);
                        kr_date.Add(killer, season_date);
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

                if (worksheet.Cell(4, firstDataColumn).GetString().Contains("Dragon Rush") ||
                    worksheet.Cell(4, firstDataColumn).GetString().Contains("Wither Rush") ||
                    worksheet.Cell(4, firstDataColumn).GetString().Contains("Realm Rush") ||
                    worksheet.Cell(4, firstDataColumn).GetString().Contains("Bolas Rush") ||
                    worksheet.Cell(4, firstDataColumn).GetString().Contains("Escape From Gaia") ||
                    worksheet.Cell(4, firstDataColumn).GetString().Contains("Trouble In Paradise") ||
                    worksheet.Cell(4, firstDataColumn).GetString().Contains("Dragon Rush Deviation Version") ||
                    worksheet.Cell(4, firstDataColumn).GetString().Contains("Hydra Rush"))
                {
                    IXLCell dragonRushCell = worksheet.Cell(firstDataRow + (seasonSize - 1), lastDataColumn);
                    if (!dragonRushCell.CellLeft().GetString().Equals("Winner"))
                    {
                        Console.ForegroundColor = ConsoleColor.Red;
                        Console.WriteLine("ERROR: Winner is not the last line of Dragon Rush");
                    }

                    while (dragonRushCell.GetString().Equals("Nothing"))
                    {
                        if (!dragonRushCell.CellLeft().GetString().Equals("Winner"))
                        {
                            dragon_rush_ru = 1;
                        }
                        dragonRushCell = dragonRushCell.CellAbove();
                    }
                }

                lastAliveCell = worksheet.Cell(firstDataRow + (seasonSize - 1), lastDataColumn);
                while (lastAliveCell.CellAbove().GetString().Equals("Nothing"))
                {
                    lastAliveCell = lastAliveCell.CellAbove();
                }

                //If FFA no need to look for teams, else looks for the team
                if (worksheet.Cell(3, middleDataColumn).GetString().Equals("FFA"))
                {
                    gs_wins[seasonWinner] += 1;

                    //Add to winner list
                    wl_rounds.Add(round_name);
                    wl_seasons.Add(season_number);
                    wl_dates.Add(season_date);
                    wl_players.Add(seasonWinner);
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
                                gs_wins[winner] += 1;

                                //Add to winner list
                                wl_rounds.Add(round_name);
                                wl_seasons.Add(season_number);
                                wl_dates.Add(season_date);
                                wl_players.Add(winner);
                            }
                        }
                    }
                }

                //Detects double kill runner ups
                if (lastAliveCell.CellAbove().GetString().Equals(lastAliveCell.CellLeft().CellLeft().CellAbove().CellAbove().GetString())
                    && lastAliveCell.CellAbove().CellAbove().GetString().Equals(lastAliveCell.CellLeft().CellLeft().CellAbove().GetString()))
                {
                    if (worksheet.Cell(3, middleDataColumn).GetString().Equals("FFA"))
                    {
                        double_kill_runnerup = 1;
                    }
                    else
                    {
                        if (!winningTeam.Contains(lastAliveCell.CellAbove().GetString()) && !winningTeam.Contains(lastAliveCell.CellAbove().CellAbove().GetString()))
                        {
                            double_kill_runnerup = 1;
                        }
                    }
                }
            }
            else
            {
                //Check for a double kill ending
                if (worksheet.Cell(firstDataRow + (seasonSize - 1), lastDataColumn).GetString().Equals(worksheet.Cell(firstDataRow + (seasonSize - 2), firstDataColumn).GetString())
                    && worksheet.Cell(firstDataRow + (seasonSize - 2), lastDataColumn).GetString().Equals(worksheet.Cell(firstDataRow + (seasonSize - 1), firstDataColumn).GetString()))
                {
                    //Double kill ending so 2 winners
                    double_kill_ending = 1;
                    String seasonWinner1 = worksheet.Cell(firstDataRow + (seasonSize - 1), firstDataColumn).GetString();
                    String seasonWinner2 = worksheet.Cell(firstDataRow + (seasonSize - 2), firstDataColumn).GetString();
                    winnerCell = worksheet.Cell(firstDataRow + (seasonSize - 1), firstDataColumn);
                    winnerCell2 = worksheet.Cell(firstDataRow + (seasonSize - 2), firstDataColumn);

                    //If FFA no need to look for teams, else looks for the team
                    if (worksheet.Cell(3, middleDataColumn).GetString().Equals("FFA"))
                    {
                        gs_wins[seasonWinner1] += 1;
                        gs_wins[seasonWinner2] += 1;

                        //Add to winner list
                        wl_rounds.Add(round_name);
                        wl_rounds.Add(round_name);
                        wl_seasons.Add(season_number);
                        wl_seasons.Add(season_number);
                        wl_dates.Add(season_date);
                        wl_dates.Add(season_date);
                        wl_players.Add(seasonWinner1);
                        wl_players.Add(seasonWinner2);
                    }
                    else
                    {
                        //Figures out the full team that won the season
                        foreach (String team in seasonTeams)
                        {
                            if (team.Contains(seasonWinner1))
                            {
                                winningTeam = team;

                                //Splits the team string to get each player and gives them a win
                                String[] winners = team.Split(separator);
                                foreach (String winner in winners)
                                {
                                    gs_wins[winner] += 1;

                                    //Add to winner list
                                    wl_rounds.Add(round_name);
                                    wl_seasons.Add(season_number);
                                    wl_dates.Add(season_date);
                                    wl_players.Add(winner);
                                }
                            }
                            if (team.Contains(seasonWinner2))
                            {
                                winningTeam2 = team;

                                //Splits the team string to get each player and gives them a win
                                String[] winners = team.Split(separator);
                                foreach (String winner in winners)
                                {
                                    gs_wins[winner] += 1;

                                    //Add to winner list
                                    wl_rounds.Add(round_name);
                                    wl_seasons.Add(season_number);
                                    wl_dates.Add(season_date);
                                    wl_players.Add(winner);
                                }
                            }
                        }
                    }
                }
                else
                {
                    dragon_win = 1;
                }
            }

            //Get the runner ups of the season
            //If FFA it has to be the player above
            //Else figures out the next team after the winners
            if (dragon_win == 0)
            {
                if (dragon_rush_ru == 0)
                {
                    if (double_kill_runnerup == 0)
                    {
                        if (worksheet.Cell(3, middleDataColumn).GetString().Equals("FFA"))
                        {
                            if (double_kill_ending == 1)
                            {
                                gs_runnerup[winnerCell2.CellAbove().GetString()] += 1;

                                //Add to runner up list
                                rul_rounds.Add(round_name);
                                rul_seasons.Add(season_number);
                                rul_dates.Add(season_date);
                                rul_players.Add(winnerCell2.CellAbove().GetString());
                            }
                            else
                            {
                                gs_runnerup[winnerCell.CellAbove().GetString()] += 1;

                                //Add to runner up list
                                rul_rounds.Add(round_name);
                                rul_seasons.Add(season_number);
                                rul_dates.Add(season_date);
                                rul_players.Add(winnerCell.CellAbove().GetString());
                            }
                        }
                        else
                        {
                            if (double_kill_ending == 1)
                            {

                                while (winningTeam.Contains(winnerCell.CellAbove().GetString()) && winningTeam2.Contains(winnerCell2.CellAbove().GetString()))
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
                                            gs_runnerup[runner_up] += 1;

                                            //Add to runner up list
                                            rul_rounds.Add(round_name);
                                            rul_seasons.Add(season_number);
                                            rul_dates.Add(season_date);
                                            rul_players.Add(runner_up);
                                        }
                                    }
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
                                            gs_runnerup[runner_up] += 1;

                                            //Add to runner up list
                                            rul_rounds.Add(round_name);
                                            rul_seasons.Add(season_number);
                                            rul_dates.Add(season_date);
                                            rul_players.Add(runner_up);
                                        }
                                    }
                                }
                            }
                        }
                    }
                    else
                    {
                        if (worksheet.Cell(3, middleDataColumn).GetString().Equals("FFA"))
                        {
                            gs_runnerup[winnerCell.CellAbove().GetString()] += 1;
                            gs_runnerup[winnerCell.CellAbove().CellAbove().GetString()] += 1;

                            //Add to runner up list
                            rul_rounds.Add(round_name);
                            rul_rounds.Add(round_name);
                            rul_seasons.Add(season_number);
                            rul_seasons.Add(season_number);
                            rul_dates.Add(season_date);
                            rul_dates.Add(season_date);
                            rul_players.Add(winnerCell.CellAbove().GetString());
                            rul_players.Add(winnerCell.CellAbove().CellAbove().GetString());
                        }
                        else
                        {
                            while (winningTeam.Contains(winnerCell.CellAbove().GetString()))
                            {
                                winnerCell = winnerCell.CellAbove();
                            }

                            //Figures out the full team of runner ups
                            String seasonRunnerUp = winnerCell.CellAbove().GetString();
                            String seasonRunnerUp2 = winnerCell.CellAbove().CellAbove().GetString();
                            foreach (String team in seasonTeams)
                            {
                                if (team.Contains(seasonRunnerUp))
                                {
                                    //Splits the team string to get each player and gives them a runner up
                                    String[] runnerups = team.Split(separator);
                                    foreach (String runner_up in runnerups)
                                    {
                                        gs_runnerup[runner_up] += 1;

                                        //Add to runner up list
                                        rul_rounds.Add(round_name);
                                        rul_seasons.Add(season_number);
                                        rul_dates.Add(season_date);
                                        rul_players.Add(runner_up);
                                    }
                                }

                                if (team.Contains(seasonRunnerUp2))
                                {
                                    //Splits the team string to get each player and gives them a runner up
                                    String[] runnerups = team.Split(separator);
                                    foreach (String runner_up in runnerups)
                                    {
                                        gs_runnerup[runner_up] += 1;

                                        //Add to runner up list
                                        rul_rounds.Add(round_name);
                                        rul_seasons.Add(season_number);
                                        rul_dates.Add(season_date);
                                        rul_players.Add(runner_up);
                                    }
                                }
                            }
                        }
                    }
                }
                else
                {
                    IXLCell runnerUpCheck = worksheet.Cell(firstDataRow + (seasonSize - 1), lastDataColumn);
                    IXLCell runnerUpPlayer = worksheet.Cell(firstDataRow + (seasonSize - 1), firstDataColumn);

                    while (runnerUpCheck.GetString().Equals("Nothing"))
                    {
                        if (!runnerUpCheck.CellLeft().GetString().Equals("Winner"))
                        {
                            gs_runnerup[runnerUpPlayer.GetString()] += 1;

                            rul_rounds.Add(round_name);
                            rul_seasons.Add(season_number);
                            rul_dates.Add(season_date);
                            rul_players.Add(runnerUpPlayer.GetString());
                        }
                        runnerUpCheck = runnerUpCheck.CellAbove();
                        runnerUpPlayer = runnerUpPlayer.CellAbove();
                    }
                }
            }
            else
            {
                //Dragon wins the season
                String seasonRunnerUp = worksheet.Cell(firstDataRow + (seasonSize - 1), firstDataColumn).GetString();

                if (worksheet.Cell(3, middleDataColumn).GetString().Equals("FFA"))
                {
                    gs_runnerup[seasonRunnerUp] += 1;

                    rul_rounds.Add(round_name);
                    rul_seasons.Add(season_number);
                    rul_dates.Add(season_date);
                    rul_players.Add(seasonRunnerUp);
                }
                else
                {
                    foreach (String team in seasonTeams)
                    {
                        if (team.Contains(seasonRunnerUp))
                        {
                            //Splits the team string to get each player and gives them a runner up
                            String[] runnerups = team.Split(separator);
                            foreach (String runner_up in runnerups)
                            {
                                gs_runnerup[runner_up] += 1;

                                //Add to runner up list
                                rul_rounds.Add(round_name);
                                rul_seasons.Add(season_number);
                                rul_dates.Add(season_date);
                                rul_players.Add(runner_up);
                            }
                        }
                    }
                }

            }
        }
    }

    rl_rostersizes.Add(roundRoster.Count);
    String rppath = "..\\..\\..\\Reddit Posts\\" + worksheet.Name + ".txt";
    String[] placement = { "1st", "2nd", "3rd", "4th", "5th", "6th", "7th", "8th", "9th", "10th",
                            "11th", "12th", "13th", "14th", "15th", "16th", "17th", "18th", "19th", "20th",
                            "21st", "22nd", "23rd", "24th", "25th", "26th", "27th", "28th", "29th", "30th",
                            "31st", "32nd", "33rd", "34th", "35th", "36th", "37th", "38th", "39th", "40th",
                            "41st", "42nd", "43rd", "44th", "45th", "46th", "47th", "48th", "49th", "50th",
                            "51st", "52nd", "53rd", "54th", "55th", "56th", "57th", "58th", "59th", "60th",
                            "61st", "62nd", "63rd", "64th", "65th", "66th", "67th", "68th", "69th", "70th",
                            "71st", "72nd", "73rd", "74th", "75th", "76th", "77th", "78th", "79th", "80th",
                            "81st", "82nd", "83rd", "84th", "85th", "86th", "87th", "88th", "89th", "90th",
                            "91st", "92nd", "93rd", "94th", "95th", "96th", "97th", "98th", "99th", "100th",
                            "101st", "102nd", "103rd", "104th", "105th", "106th", "107th", "108th", "109th", "110th",
                            "111th", "112th", "113th", "114th", "115th", "116th", "117th", "118th", "119th", "120th",
                            "121st", "122nd", "123rd", "124th", "125th", "126th", "127th", "128th", "129th", "130th",
                            "131st", "132nd", "133rd", "134th", "135th", "136th", "137th", "138th", "139th", "140th",
                            "141st", "142nd", "143rd", "144th", "145th", "146th", "147th", "148th", "149th", "150th",
                            "151st", "152nd", "153rd", "154th", "155th", "156th", "157th", "158th", "159th", "160th",
                            "161st", "162nd", "163rd", "164th", "165th", "166th", "167th", "168th", "169th", "170th",
                            "171st", "172nd", "173rd", "174th", "175th", "176th", "177th", "178th", "179th", "180th",
                            "181st", "182nd", "183rd", "184th", "185th", "186th", "187th", "188th", "189th", "190th",
                            "191st", "192nd", "193rd", "194th", "195th", "196th", "197th", "198th", "199th", "200th",};
    int ranking = 0;
    int ties = 1;
    int currentkill = 0;

    File.WriteAllText(rppath, "## " + worksheet.Name + " Statistics" + Environment.NewLine);
    File.AppendAllText(rppath, Environment.NewLine + "---");
    File.AppendAllText(rppath, Environment.NewLine + "### Winners" + Environment.NewLine);
    File.AppendAllText(rppath, Environment.NewLine + Environment.NewLine + "---");
    File.AppendAllText(rppath, Environment.NewLine + "### Runner Ups" + Environment.NewLine);
    File.AppendAllText(rppath, Environment.NewLine + Environment.NewLine + "---");
    File.AppendAllText(rppath, Environment.NewLine + "### Most Kills" + Environment.NewLine);
    File.AppendAllText(rppath, Environment.NewLine + Environment.NewLine + "---");
    File.AppendAllText(rppath, Environment.NewLine + "### Most Kills (Team)" + Environment.NewLine);
    File.AppendAllText(rppath, Environment.NewLine + Environment.NewLine + "---");
    File.AppendAllText(rppath, Environment.NewLine + "### First Damage" + Environment.NewLine);
    File.AppendAllText(rppath, Environment.NewLine + Environment.NewLine + "---");
    File.AppendAllText(rppath, Environment.NewLine + "### Ironman" + Environment.NewLine);
    File.AppendAllText(rppath, Environment.NewLine + Environment.NewLine + "---");
    File.AppendAllText(rppath, Environment.NewLine + "### First Blood" + Environment.NewLine + Environment.NewLine);
    foreach (String firstbloods in rp_firstblood)
    {
        File.AppendAllText(rppath, firstbloods + Environment.NewLine);
    }
    File.AppendAllText(rppath, "---");
    File.AppendAllText(rppath, Environment.NewLine + "### First Death" + Environment.NewLine + Environment.NewLine);
    foreach (String firstdeaths in rp_firstdeath)
    {
        File.AppendAllText(rppath, firstdeaths + Environment.NewLine);
    }
    File.AppendAllText(rppath, "---");
    File.AppendAllText(rppath, Environment.NewLine + "### Kills" + Environment.NewLine + Environment.NewLine);
    foreach (String kills in rp_kills_list.Keys)
    {
        rp_kills_list[kills] = rp_kills_list[kills].Remove(rp_kills_list[kills].Length - 2);
    }
    rp_kills = rp_kills.OrderBy(kvp => kvp.Key).ToDictionary(kvp => kvp.Key, kvp => kvp.Value);
    rp_kills = rp_kills.OrderByDescending(key => key.Value).ToDictionary(keyValuePair => keyValuePair.Key, keyValuePair => keyValuePair.Value);
    foreach (String kills in rp_kills.Keys)
    {
        if (currentkill > 0)
        {
            if (rp_kills[kills] == currentkill)
            {
                ties += 1;
            }
            else
            {
                ranking += ties;
                ties = 1;
            }
        }
        File.AppendAllText(rppath, "**" + placement[ranking] + " - " + kills + " (" + rp_kills[kills] + "):** " + rp_kills_list[kills] + Environment.NewLine + Environment.NewLine);
        currentkill = rp_kills[kills];
    }
    File.AppendAllText(rppath, "---");
    File.AppendAllText(rppath, Environment.NewLine + "### PvE Deaths" + Environment.NewLine + Environment.NewLine);
    foreach (String pvedeaths in rp_pvedeaths_list.Keys)
    {
        rp_pvedeaths_list[pvedeaths] = rp_pvedeaths_list[pvedeaths].Remove(rp_pvedeaths_list[pvedeaths].Length - 2);
    }
    rp_pvedeaths = rp_pvedeaths.OrderBy(kvp => kvp.Key).ToDictionary(kvp => kvp.Key, kvp => kvp.Value);
    rp_pvedeaths = rp_pvedeaths.OrderByDescending(key => key.Value).ToDictionary(keyValuePair => keyValuePair.Key, keyValuePair => keyValuePair.Value);
    foreach (String pvedeaths in rp_pvedeaths.Keys)
    {
        File.AppendAllText(rppath, "**" + pvedeaths + " (" + rp_pvedeaths[pvedeaths] + "):** " + rp_pvedeaths_list[pvedeaths] + Environment.NewLine + Environment.NewLine);
    }
    File.AppendAllText(rppath, "---");
    File.AppendAllText(rppath, Environment.NewLine + "### Participation" + Environment.NewLine);
    File.AppendAllText(rppath, Environment.NewLine + Environment.NewLine + "---");
    File.AppendAllText(rppath, Environment.NewLine + "### Debutants" + Environment.NewLine + Environment.NewLine);
    foreach (String debutants in rp_debutants)
    {
        File.AppendAllText(rppath, debutants + Environment.NewLine);
    }
    File.AppendAllText(rppath, "---");

    rp_firstblood.Clear();
    rp_firstdeath.Clear();
    rp_kills.Clear();
    rp_kills_list.Clear();
    rp_pvedeaths.Clear();
    rp_pvedeaths_list.Clear();
    rp_debutants.Clear();
}

//Updates KDR & KPR of player
//If infinite sets it to the amount of kills they have
foreach (String player in gs_deaths.Keys)
{
    if (gs_deaths[player].Equals(0))
    {
        gs_kdr[player] = Convert.ToDouble(gs_kills[player]);
    }
    else
    {
        gs_kdr[player] = Convert.ToDouble(gs_kills[player]) / Convert.ToDouble(gs_deaths[player]);
    }
    gs_kpr[player] = Convert.ToDouble(gs_kills[player]) / Convert.ToDouble(gs_seasonsplayed[player]);
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

//Making PvE List Page
var pve_list = roundlist.AddWorksheet("PvE List");
pve_list.Column("A").Width = 34;
pve_list.Column("B").Width = 6;
pve_list.Cell("A1").InsertData(unique_pve_deaths.Keys);
pve_list.Cell("B1").InsertData(unique_pve_deaths.Values);

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

//PvE Death list
var pve_death = statscompiled.AddWorksheet("PvE Deaths");
pve_death.Column("C").Style.NumberFormat.Format = "mm/dd/yyyy";
pve_death.Column("A").Width = 34;
pve_death.Column("B").Width = 6;
pve_death.Column("C").Width = 20;
pve_death.Cell("A1").InsertData(pvel_rounds);
pve_death.Cell("B1").InsertData(pvel_seasons);
pve_death.Cell("C1").InsertData(pvel_dates);
pve_death.Cell("D1").InsertData(pvel_players);
pve_death.Sort(3);

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

//Alive list
var alive = statscompiled.AddWorksheet("Alive");
alive.Column("C").Style.NumberFormat.Format = "mm/dd/yyyy";
alive.Column("A").Width = 34;
alive.Column("B").Width = 6;
alive.Column("C").Width = 20;
alive.Cell("A1").InsertData(al_rounds);
alive.Cell("B1").InsertData(al_seasons);
alive.Cell("C1").InsertData(al_dates);
alive.Cell("D1").InsertData(al_players);
alive.Sort(3);

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
Console.ForegroundColor = ConsoleColor.Green;
Console.WriteLine("Stats are now compiled!");

static String getPvEDeath(IXLCell cell)
{
    String pvedeath = "";

    if (cell.CellLeft().GetString().Contains("lava")
    && !cell.CellLeft().GetString().Contains("discovered"))//Lava
    {
        pvedeath = "Lava";
    }
    else if (cell.CellLeft().GetString().Contains("discovered"))//Magma
    {
        pvedeath = "Magma";
    }
    else if (cell.CellLeft().GetString().Contains("ground")
            || cell.CellLeft().GetString().Contains("doomed")
            || cell.CellLeft().GetString().Contains("fell")
            && !cell.CellLeft().GetString().Contains("world"))//Fall
    {
        pvedeath = "Fall";
    }
    else if (cell.CellLeft().GetString().Contains("world"))//Void
    {
        pvedeath = "Void";
    }
    else if (cell.CellLeft().GetString().Contains("drowned"))//Drowned
    {
        pvedeath = "Drowning";
    }
    else if (cell.CellLeft().GetString().Contains("suffocated"))//Suffocation
    {
        pvedeath = "Suffocation";
    }
    else if (cell.CellLeft().GetString().Contains("burnt")
            || cell.CellLeft().GetString().Contains("burned"))//Burn
    {
        pvedeath = "Burning";
    }
    else if (cell.CellLeft().GetString().Contains("starved"))//Starved
    {
        pvedeath = "Starvation";
    }
    else if (cell.CellLeft().GetString().Contains("fallout"))//Fallout
    {
        pvedeath = "Fallout";
    }
    else if (cell.CellLeft().GetString().Contains("swords"))//Krenzinator
    {
        pvedeath = "Diamond Sword";
    }
    else if (cell.CellLeft().GetString().Contains("water"))//Water
    {
        pvedeath = "Water";
    }
    else if (cell.CellLeft().GetString().Contains("disqualified"))//Disqualified
    {
        pvedeath = "Disqualified";
    }
    else if (cell.CellLeft().GetString().Contains("bats"))//Bats
    {
        pvedeath = "Bats";
    }
    else if (cell.CellLeft().GetString().Contains("extra"))//Extra Damage
    {
        pvedeath = "Extra Damage";
    }
    else if (cell.CellLeft().GetString().Contains("diamonds"))//Blood Diamonds
    {
        pvedeath = "Blood Diamonds";
    }
    else if (cell.CellLeft().GetString().Contains("gambled"))//Gamble
    {
        pvedeath = "Gambling";
    }
    else if (cell.CellLeft().GetString().Contains("button"))//Push The Button
    {
        pvedeath = "Push The Button";
    }
    else if (cell.CellLeft().GetString().Contains("hell"))//Go To Hell
    {
        pvedeath = "Go To Hell";
    }
    else if (cell.CellLeft().GetString().Contains("comply"))//Comply
    {
        pvedeath = "Comply";
    }
    else if (cell.CellLeft().GetString().Contains("learned"))//Newton's Third Law
    {
        pvedeath = "Newtons Third Law";
    }
    else if (cell.CellLeft().GetString().Contains("infiltrator"))//Infiltrators
    {
        pvedeath = "Infiltrator";
    }
    else if (cell.CellLeft().GetString().Contains("love"))//Love
    {
        pvedeath = "Love";
    }
    else if (cell.CellLeft().GetString().Contains("Design"))//Bed
    {
        pvedeath = "Bed";
    }
    else if (cell.CellLeft().GetString().Contains("blew"))//Explosion
    {
        pvedeath = "Explosion";
    }
    else if (cell.CellLeft().GetString().Contains("sneaked"))//Sneaking
    {
        pvedeath = "Sneaking";
    }
    else if (cell.CellLeft().GetString().Contains("withered"))//Withered
    {
        pvedeath = "Withered";
    }
    else if (cell.CellLeft().GetString().Contains("timed")
            || cell.CellLeft().GetString().Contains("disconnected")
            || cell.CellLeft().GetString().Contains("offline"))//Timed Out
    {
        pvedeath = "Left";
    }
    else if (cell.CellLeft().GetString().Contains("stalagmite")
            || cell.CellLeft().GetString().Contains("stalactite"))//Dripstone
    {
        pvedeath = "Dripstone";
    }
    else if (cell.CellLeft().GetString().Contains("anvil"))//Anvil
    {
        pvedeath = "Anvil";
    }
    else if (cell.CellLeft().GetString().Contains("pricked"))//Cactus
    {
        pvedeath = "Cactus";
    }
    else if (cell.CellLeft().GetString().Contains("poked"))//Berry
    {
        pvedeath = "Sweet Berry Bush";
    }
    else if (cell.CellLeft().GetString().Contains("kinetic"))//Elytra
    {
        pvedeath = "Elytra";
    }
    else if (cell.CellLeft().GetString().Contains("bang"))//Firework
    {
        pvedeath = "Firework";
    }
    else if (cell.CellLeft().GetString().Contains("died"))//Death
    {
        pvedeath = "Death";
    }
    else if (cell.CellLeft().GetString().Contains("flames"))//Fire
    {
        pvedeath = "Fire";
    }
    else if (cell.CellLeft().GetString().Contains("pummeled"))//Pummeled
    {
        pvedeath = "Pummeled";
    }
    else if (cell.CellLeft().GetString().Contains("magic"))//Magic
    {
        pvedeath = "Potion";
    }
    else if (cell.CellLeft().GetString().Contains("shot"))//Arrow
    {
        pvedeath = "Arrow";
    }
    else
    {
        pvedeath = "N/A";
    }

    return pvedeath;
}