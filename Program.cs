using System.Reflection.PortableExecutable;
using System.Runtime.InteropServices;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Office2010.PowerPoint;
using DocumentFormat.OpenXml.Spreadsheet;
using StatsAnalyzer;

// If 1 calculates list of round gaps
// If 2 calculates list of global gaps
// If 3 makes a personal stats sheet for a player
// If 4 calculates the global stats
int Statfunction = 4;
// Player to analyze for the stats sheet
String playerStats = "Chasmic";
// Select the stat doc for global stats
// 1 for reddit
// 2 for non-reddit
// 3 for live rounds
int statDoc = 1;

if (Statfunction == 1)
{
    String roundsListDoc = "..\\..\\..\\Stats Sheet\\Reddit\\Round_List.xlsx";
    List<RoundsGaps> roundsGapList = new List<RoundsGaps>();

    if (!File.Exists(roundsListDoc))
    {
        Console.ForegroundColor = ConsoleColor.Red;
        Console.WriteLine($"Error: File not found at {roundsListDoc}");
        return;
    }

    using var roundsDocument = new XLWorkbook(roundsListDoc);
    var roundList = roundsDocument.Worksheet(1);

    Dictionary<String, int> allRounds = new Dictionary<String, int>();
    IXLRange roundRange = roundList.Range(1, 1, roundList.RangeUsed()!.RowCount(), 1);
    foreach (IXLCell round in roundRange.Cells())
    {
        allRounds.Add(round.Value.ToString(), round.CellRight().GetValue<int>());
    }

    var rosterList = roundsDocument.Worksheet(3);

    foreach (String round in allRounds.Keys)
    {
        String lastRoundSeason = "none";
        Dictionary<String, String> lastSeasonPlayed = new Dictionary<String, String>();
        Dictionary<String, DateTime> lastDatePlayed = new Dictionary<String, DateTime>();
        //Skips Rounds with 2 or less seasons.
        if (allRounds[round] > 2)
        {
            //Goes through every season of a round and checks gaps for missed seasons.
            IXLRange seasonRange = rosterList.Range(1, 1, rosterList.RangeUsed()!.RowCount(), 1);
            foreach (IXLCell seasonCell in seasonRange.CellsUsed())
            {
                int seasonRow = seasonCell.WorksheetRow().RowNumber();
                String seasonRound = rosterList.Cell(seasonRow, 1).Value.ToString();
                String seasonNumber = rosterList.Cell(seasonRow, 2).Value.ToString();
                DateTime seasonDate = rosterList.Cell(seasonRow, 3).GetDateTime();
                Dictionary<int, String> _roundPlayer = new Dictionary<int, String>();
                Dictionary<int, String> _roundStartSeason = new Dictionary<int, String>();
                Dictionary<int, String> _roundEndSeason = new Dictionary<int, String>();
                Dictionary<int, DateTime> _roundGapDate = new Dictionary<int, DateTime>();

                //Exceptions for crossover rounds with different names.
                String roundCompare = round;
                if (seasonRound.Contains(roundCompare))
                {
                    if (seasonRound == "WMC x Phobia" ||
                        seasonRound == "The Melon Blooded x Scattershot" ||
                        seasonRound == "Phobia x Cinema")
                    {
                        roundCompare = seasonRound;
                    }
                }

                if (seasonRound == roundCompare)
                {
                    IXLRange rosterRange = rosterList.Range(seasonRow, 4, seasonRow, 129);
                    foreach (IXLCell rosterCell in rosterRange.CellsUsed())
                    {
                        String seasonPlayer = rosterCell.GetString();
                        if (!lastRoundSeason.Equals("none"))
                        {
                            if (lastSeasonPlayed.ContainsKey(seasonPlayer))
                            {
                                if (!lastSeasonPlayed[seasonPlayer].Equals(lastRoundSeason))
                                {
                                    TimeSpan timeDiff = seasonDate - lastDatePlayed[seasonPlayer];
                                    int gapDays = (int)timeDiff.TotalDays;
                                    if (gapDays > 1095)
                                    {
                                        if (_roundPlayer.ContainsKey(gapDays))
                                        {
                                            _roundPlayer[gapDays] = _roundPlayer[gapDays] + ", " + seasonPlayer;
                                        }
                                        else
                                        {
                                            _roundPlayer.Add(gapDays, seasonPlayer);
                                            _roundStartSeason.Add(gapDays, lastSeasonPlayed[seasonPlayer]);
                                            _roundEndSeason.Add(gapDays, seasonNumber);
                                            _roundGapDate.Add(gapDays, seasonDate);
                                        }
                                    }
                                }

                                lastSeasonPlayed[seasonPlayer] = seasonNumber;
                                lastDatePlayed[seasonPlayer] = seasonDate;
                            }
                            else
                            {
                                lastSeasonPlayed.Add(seasonPlayer, seasonNumber);
                                lastDatePlayed.Add(seasonPlayer, seasonDate);
                            }
                        }
                        else
                        {
                            lastSeasonPlayed.Add(seasonPlayer, seasonNumber);
                            lastDatePlayed.Add(seasonPlayer, seasonDate);
                        }
                    }
                    lastRoundSeason = seasonNumber;

                    foreach (int gap in _roundPlayer.Keys)
                    {
                        roundsGapList.Add(new RoundsGaps(_roundPlayer[gap], gap, roundCompare, _roundStartSeason[gap], _roundEndSeason[gap], _roundGapDate[gap]));
                    }
                }
            }
            Console.WriteLine("Checked gaps for " + round + "!");
        }
    }

    //Saves the new docs
    DataExporter.SaveRoundsGaps(roundsGapList);
}
else if (Statfunction == 2)
{
    String playerListDoc = "..\\..\\..\\Stats Sheet\\Reddit\\Global_Stats.xlsx";
    String roundsListDoc = "..\\..\\..\\Stats Sheet\\Reddit\\Round_List.xlsx";
    List<GlobalGaps> globalGapList = new List<GlobalGaps>();

    if (!File.Exists(playerListDoc))
    {
        Console.ForegroundColor = ConsoleColor.Red;
        Console.WriteLine($"Error: File not found at {playerListDoc}");
        return;
    }
    if (!File.Exists(roundsListDoc))
    {
        Console.ForegroundColor = ConsoleColor.Red;
        Console.WriteLine($"Error: File not found at {roundsListDoc}");
        return;
    }

    using var playersDocument = new XLWorkbook(playerListDoc);
    var playerList = playersDocument.Worksheet(1);

    //Gets a list of all players.
    Dictionary<String, int> allPlayers = new Dictionary<String, int>();
    IXLRange playerRange = playerList.Range(1, 1, playerList.RangeUsed()!.RowCount(), 1);
    foreach (IXLCell player in playerRange.Cells())
    {
        allPlayers.Add(player.Value.ToString(), player.CellRight().GetValue<int>());
    }

    using var roundsDocument = new XLWorkbook(roundsListDoc);
    var rosterList = roundsDocument.Worksheet(3);

    foreach (String player in allPlayers.Keys)
    {
        String lastRound = "Fake";
        String lastSeason = "1z";
        DateTime lastPlayed = new DateTime(2012, 1, 1);
        //Skips players with 1 round played.
        if (allPlayers[player] > 1)
        {
            //Goes through every seasons and compares gaps of players.
            IXLRange seasonRange = rosterList.Range(1, 1, rosterList.RangeUsed()!.RowCount(), 1);
            foreach (IXLCell seasonCell in seasonRange.CellsUsed())
            {
                int seasonRow = seasonCell.WorksheetRow().RowNumber();
                IXLRange rosterRange = rosterList.Range(seasonRow, 4, seasonRow, 129);
                foreach (IXLCell rosterCell in rosterRange.CellsUsed())
                {
                    String seasonPlayer = rosterCell.GetString();
                    String seasonRound = rosterList.Cell(seasonRow, 1).Value.ToString();
                    String seasonNumber = rosterList.Cell(seasonRow, 2).Value.ToString();
                    DateTime seasonDate = rosterList.Cell(seasonRow, 3).GetDateTime();

                    if (seasonPlayer == player)
                    {
                        if (lastPlayed != new DateTime(2012, 1, 1))
                        {
                            TimeSpan timeDiff = seasonDate - lastPlayed;
                            int gapDays = (int)timeDiff.TotalDays;
                            if (gapDays > 1095)
                            {
                                String playerCompare = player;
                                foreach (var playerGap in globalGapList.ToList())
                                {
                                    if (playerGap.DayGap == gapDays &&
                                        playerGap.StartRound == lastRound &&
                                        playerGap.StartSeason == lastSeason &&
                                        playerGap.EndRound == seasonRound &&
                                        playerGap.EndSeason == seasonNumber)
                                    {
                                        playerCompare = playerGap.Player + ", " + player;
                                        globalGapList.Remove(playerGap);
                                    }
                                }
                                globalGapList.Add(new GlobalGaps(playerCompare, gapDays, lastRound, lastSeason, lastPlayed, seasonRound, seasonNumber, seasonDate));
                            }
                        }
                        lastRound = seasonRound;
                        lastSeason = seasonNumber;
                        lastPlayed = seasonDate;
                    }
                }
            }
            Console.WriteLine("Checked gaps for " + player + "!");
        }
    }

    //Saves the gap doc
    DataExporter.SaveGlobalGaps(globalGapList);
}
else if (Statfunction == 3)
{
    //IGN of the player to analyze
    String roundsListDoc = "..\\..\\..\\Stats Sheet\\Reddit\\Round_List.xlsx";
    String statsListDoc = "..\\..\\..\\Stats Sheet\\Reddit\\Stats_Compiled.xlsx";
    List<PlayerStats> playerStatsList = new List<PlayerStats>();

    if (!File.Exists(roundsListDoc))
    {
        Console.ForegroundColor = ConsoleColor.Red;
        Console.WriteLine($"Error: File not found at {roundsListDoc}");
        return;
    }
    if (!File.Exists(statsListDoc))
    {
        Console.ForegroundColor = ConsoleColor.Red;
        Console.WriteLine($"Error: File not found at {statsListDoc}");
        return;
    }

    using var roundsDocument = new XLWorkbook(roundsListDoc);
    var rosterList = roundsDocument.Worksheet(3);

    //Makes the list of all rounds played, and sets the other variables
    for (int round = 1; round <= rosterList.Rows().Count(); round++)
    {
        IXLRange rosterRange = rosterList.Range(round, 4, round, 129);
        foreach (IXLCell cell in rosterRange.CellsUsed())
        {
            string value = cell.GetString();

            if (value == playerStats)
            {
                String roundName = rosterList.Cell(round, 1).Value.ToString();
                String seasonNumber = rosterList.Cell(round, 2).Value.ToString();
                DateTime seasonDate = rosterList.Cell(round, 3).Value.GetDateTime();
                playerStatsList.Add(new PlayerStats(roundName, seasonNumber, seasonDate));
            }
        }
    }
    Console.WriteLine("Roster Done");

    //Makes the list of the teams and formats it correctly
    //Also gets team color and makes it match the format used.
    using var statsDocument = new XLWorkbook(statsListDoc);
    var statsList = statsDocument.Worksheet(2);
    IXLRange teamsRange = statsList.Range(1, 4, statsList.RangeUsed()!.RowCount(), 4);

    foreach (IXLCell team in teamsRange.Cells())
    {
        if (team.Value.ToString().Contains(playerStats))
        {
            //Gets team for the season
            String roundName = team.CellLeft(3).Value.ToString();
            String seasonNumber = team.CellLeft(2).Value.ToString();
            String seasonTeam = team.Value.ToString() + ",";
            seasonTeam = seasonTeam.Replace(playerStats + ",", "");
            seasonTeam = seasonTeam.Replace(",", " ");

            //Gets team color for the season
            String seasonTeamColor = team.CellRight().Value.ToString() + "";
            seasonTeamColor = PlayerStats.GetTeamColorChar(seasonTeamColor);

            var seasonStats = playerStatsList.Find(round => round.Round == roundName && round.Season == seasonNumber);
            if (seasonStats != null)
            {
                seasonStats.Team = seasonTeam;
                seasonStats.TeamColor = seasonTeamColor;
            }
        }
    }
    Console.WriteLine("Teams Done");

    //Makes the list for all the kills and deaths of a player, formats it for the doc
    statsList = statsDocument.Worksheet(1);
    IXLRange killsRange = statsList.Range(1, 6, statsList.RangeUsed()!.RowCount(), 6);
    IXLRange deathsRange = statsList.Range(1, 4, statsList.RangeUsed()!.RowCount(), 4);

    foreach (IXLCell playerDeath in killsRange.CellsUsed())
    {
        if (playerDeath.Value.ToString() == playerStats)
        {
            String roundName = playerDeath.CellLeft(5).Value.ToString();
            String seasonNumber = playerDeath.CellLeft(4).Value.ToString();

            var seasonStats = playerStatsList.Find(round => round.Round == roundName && round.Season == seasonNumber);
            if (seasonStats != null)
            {
                String playerKilled = playerDeath.CellLeft(2).Value.ToString();
                if (seasonStats.Kills == "/")
                {
                    seasonStats.Kills = playerKilled;
                }
                else
                {
                    seasonStats.Kills = seasonStats.Kills + " " + playerKilled;
                }
            }
        }
    }
    Console.WriteLine("Kills Done");

    foreach (IXLCell playerDeath in deathsRange.CellsUsed())
    {
        if (playerDeath.Value.ToString() == playerStats)
        {
            String roundName = playerDeath.CellLeft(3).Value.ToString();
            String seasonNumber = playerDeath.CellLeft(2).Value.ToString();

            var seasonStats = playerStatsList.Find(round => round.Round == roundName && round.Season == seasonNumber);
            if (seasonStats != null)
            {
                String playerDied = playerDeath.CellRight(2).Value.ToString();
                if (seasonStats.Death == "todelete")
                {
                    seasonStats.Death = playerDied;
                }
                else
                {
                    seasonStats.Death = seasonStats.Death + " " + playerDied;
                }
            }
        }
    }
    Console.WriteLine("Deaths Done");

    //Calculates the number of kills in 1 season
    foreach (var round in playerStatsList.ToList())
    {
        if (round.Kills != "/")
        {
            int count = round.Kills.Split(' ').Length;
            round.KillsTotal = count;
        }
    }
    Console.WriteLine("Kill Count Done");

    //Gets the list for first damages
    statsList = statsDocument.Worksheet(3);
    IXLRange firstDamageRange = statsList.Range(1, 4, statsList.RangeUsed()!.RowCount(), 4);

    foreach (IXLCell player in firstDamageRange.Cells())
    {
        if (player.Value.ToString().Equals(playerStats))
        {
            String roundName = player.CellLeft(3).Value.ToString();
            String seasonNumber = player.CellLeft(2).Value.ToString();

            var seasonStats = playerStatsList.Find(round => round.Round == roundName && round.Season == seasonNumber);
            if (seasonStats != null)
            {
                seasonStats.FirstDamage = "x";
            }
        }
    }
    Console.WriteLine("First Damage Done");

    //Gets the list for ironman
    statsList = statsDocument.Worksheet(4);
    IXLRange ironmanRange = statsList.Range(1, 4, statsList.RangeUsed()!.RowCount(), 4);

    foreach (IXLCell player in ironmanRange.Cells())
    {
        if (player.Value.ToString().Equals(playerStats))
        {
            String roundName = player.CellLeft(3).Value.ToString();
            String seasonNumber = player.CellLeft(2).Value.ToString();

            var seasonStats = playerStatsList.Find(round => round.Round == roundName && round.Season == seasonNumber);
            if (seasonStats != null)
            {
                seasonStats.Ironman = "x";
            }
        }
    }
    Console.WriteLine("Ironman Done");

    //Gets the list for pve deaths
    statsList = statsDocument.Worksheet(5);
    IXLRange pveRange = statsList.Range(1, 4, statsList.RangeUsed()!.RowCount(), 4);

    foreach (IXLCell player in pveRange.Cells())
    {
        if (player.Value.ToString().Equals(playerStats))
        {
            String roundName = player.CellLeft(3).Value.ToString();
            String seasonNumber = player.CellLeft(2).Value.ToString();

            var seasonStats = playerStatsList.Find(round => round.Round == roundName && round.Season == seasonNumber);
            if (seasonStats != null)
            {
                seasonStats.PveDeath = "x";
            }
        }
    }
    Console.WriteLine("PvE Deaths Done");

    //Gets the list for first deaths
    statsList = statsDocument.Worksheet(6);
    IXLRange firstDeathsRange = statsList.Range(1, 4, statsList.RangeUsed()!.RowCount(), 4);

    foreach (IXLCell player in firstDeathsRange.Cells())
    {
        if (player.Value.ToString().Equals(playerStats))
        {
            String roundName = player.CellLeft(3).Value.ToString();
            String seasonNumber = player.CellLeft(2).Value.ToString();

            var seasonStats = playerStatsList.Find(round => round.Round == roundName && round.Season == seasonNumber);
            if (seasonStats != null)
            {
                seasonStats.FirstDeath = "x";
            }
        }
    }
    Console.WriteLine("First Deaths Done");

    //Gets the list for first blood
    statsList = statsDocument.Worksheet(7);
    IXLRange firstBloodRange = statsList.Range(1, 4, statsList.RangeUsed()!.RowCount(), 4);

    foreach (IXLCell player in firstBloodRange.Cells())
    {
        if (player.Value.ToString().Equals(playerStats))
        {
            String roundName = player.CellLeft(3).Value.ToString();
            String seasonNumber = player.CellLeft(2).Value.ToString();

            var seasonStats = playerStatsList.Find(round => round.Round == roundName && round.Season == seasonNumber);
            if (seasonStats != null)
            {
                seasonStats.FirstBlood = "x";
            }
        }
    }
    Console.WriteLine("First Blood Done");

    //Gets the list for most kills
    statsList = statsDocument.Worksheet(8);
    IXLRange topFragsRange = statsList.Range(1, 4, statsList.RangeUsed()!.RowCount(), 4);

    foreach (IXLCell player in topFragsRange.Cells())
    {
        if (player.Value.ToString().Equals(playerStats))
        {
            String roundName = player.CellLeft(3).Value.ToString();
            String seasonNumber = player.CellLeft(2).Value.ToString();

            var seasonStats = playerStatsList.Find(round => round.Round == roundName && round.Season == seasonNumber);
            if (seasonStats != null)
            {
                seasonStats.TopFrag = "x";
            }
        }
    }
    Console.WriteLine("Top Frags Done");

    //Gets the list for runner ups
    statsList = statsDocument.Worksheet(9);
    IXLRange runnerUpsRange = statsList.Range(1, 4, statsList.RangeUsed()!.RowCount(), 4);

    foreach (IXLCell player in runnerUpsRange.Cells())
    {
        if (player.Value.ToString().Equals(playerStats))
        {
            String roundName = player.CellLeft(3).Value.ToString();
            String seasonNumber = player.CellLeft(2).Value.ToString();

            var seasonStats = playerStatsList.Find(round => round.Round == roundName && round.Season == seasonNumber);
            if (seasonStats != null)
            {
                seasonStats.RunnerUp = "x";
            }
        }
    }
    Console.WriteLine("Runner Ups Done");

    //Gets the list for wins
    statsList = statsDocument.Worksheet(11);
    IXLRange winRange = statsList.Range(1, 4, statsList.RangeUsed()!.RowCount(), 4);

    foreach (IXLCell player in winRange.Cells())
    {
        if (player.Value.ToString().Equals(playerStats))
        {
            String roundName = player.CellLeft(3).Value.ToString();
            String seasonNumber = player.CellLeft(2).Value.ToString();

            var seasonStats = playerStatsList.Find(round => round.Round == roundName && round.Season == seasonNumber);
            if (seasonStats != null)
            {
                seasonStats.Win = "x";
            }
        }
    }
    Console.WriteLine("Wins Done");

    //Gets the list for alives wins, if not a win (dragon rush) logs in console
    statsList = statsDocument.Worksheet(10);
    IXLRange aliveRange = statsList.Range(1, 4, statsList.RangeUsed()!.RowCount(), 4);

    foreach (IXLCell player in aliveRange.Cells())
    {
        if (player.Value.ToString().Equals(playerStats))
        {
            String roundName = player.CellLeft(3).Value.ToString();
            String seasonNumber = player.CellLeft(2).Value.ToString();

            var seasonStats = playerStatsList.Find(round => round.Round == roundName && round.Season == seasonNumber);
            if (seasonStats != null)
            {
                if (seasonStats.Win == "x")
                {
                    seasonStats.Win = "o";
                }
                else
                {
                    Console.WriteLine("Alive no win found!");
                }
            }
        }
    }
    Console.WriteLine("Alives Done");

    //Saves the new stat doc
    DataExporter.SavePlayerStats(playerStatsList);

    //Deletes temporary filled cells
    DataExporter.ClearEmptyCells();

    Console.ForegroundColor = ConsoleColor.Green;
    Console.WriteLine("Stats are now compiled for " + playerStats + "!");
}
else if (Statfunction == 4)
{
    string filePath;
    string postFolder;
    int skipSheet;

    switch (statDoc)
    {
        case 1:
            filePath = "..\\..\\..\\Global Docs\\Global RR Stats Community Document.xlsx";
            postFolder = "Reddit";
            skipSheet = 8;
            break;
        case 2:
            filePath = "..\\..\\..\\Global Docs\\Non-Reddit Stats Community Document.xlsx";
            postFolder = "NonReddit";
            skipSheet = 6;
            break;
        case 3:
            filePath = "..\\..\\..\\Global Docs\\Global Live Round Stats Community Document.xlsx";
            postFolder = "Live";
            skipSheet = 6;
            break;
        default:
            filePath = "..\\..\\..\\Global Docs\\Global RR Stats Community Document.xlsx";
            postFolder = "Reddit";
            skipSheet = 8;
            break;
    }

    if (!File.Exists(filePath))
    {
        Console.ForegroundColor = ConsoleColor.Red;
        Console.WriteLine($"Error: File not found at {filePath}");
        return;
    }

    //Big list of variables that are saved at the end
    //Variables for the rounds list
    List<RoundList> roundList = new List<RoundList>();
    //Variables for the all rosters doc
    List<RostersList> rostersList = new List<RostersList>();
    //Variables for the all rosters without non-reddits
    List<RostersList> rostersListNR = new List<RostersList>();
    //Variables for the kill list
    List<KillsList> killsList = new List<KillsList>();
    //Variables for the team list
    List<TeamsList> teamsList = new List<TeamsList>();
    //Variables for the first damage list
    List<StatsList> firstDamageList = new List<StatsList>();
    //Variables for the ironman list
    List<StatsList> ironmanList = new List<StatsList>();
    //Variables for the pve death list
    List<StatsList> pveDeathList = new List<StatsList>();
    //Variables for the first death list
    List<StatsList> firstDeathList = new List<StatsList>();
    //Variables for the first blood list
    List<StatsList> firstBloodList = new List<StatsList>();
    //Variables for the top kills list
    List<StatsList> topFragList = new List<StatsList>();
    //Variables for the runner up list
    List<StatsList> runnerUpList = new List<StatsList>();
    //Variables for the alive list
    List<StatsList> aliveList = new List<StatsList>();
    //Variables for the wins list
    List<StatsList> winList = new List<StatsList>();
    //Variables for the round debut list
    List<RoundDebut> roundDebutsList = new List<RoundDebut>();
    //Variables for the round debut list without non-reddits
    List<RoundDebut> roundDebutsListNR = new List<RoundDebut>();
    //Variables for the global stats
    List<GlobalStats> globalStatsList = new List<GlobalStats>();
    //Variables for the kill records
    List<KillRecords> killRecordsList = new List<KillRecords>();
    //Unique PvE Deaths
    List<PveCausesList> pveCausesList = new List<PveCausesList>();
    List<GamemodesList> gamemodesList = new List<GamemodesList>();

    //Goes through every stats tabs on the doc
    using var globalDocument = new XLWorkbook(filePath);
    for (int sheet = skipSheet; sheet <= globalDocument.Worksheets.Count; sheet++)
    {
        List<String> roundRoster = new List<String>();
        RedditPosts redditPosts = new RedditPosts();
        //Collecting the Name, the total amount of seasons & the date of S1 for the round
        var roundPage = globalDocument.Worksheet(sheet);
        String round_name = roundPage.Name;
        int round_totalseasons = (roundPage.Columns().Count() - 1) / 3;
        DateTime round_debutdate = roundPage.Cell(2, 2).GetDateTime();
        String post_name = "";
        if (round_name.Contains("Sheet"))
        {
            round_name = "???";
            post_name = "3 Question Marks";
        }
        else
        {
            post_name = round_name;
        }
        roundList.Add(new RoundList(round_name, round_totalseasons, round_debutdate));
        Console.ForegroundColor = ConsoleColor.Green;
        Console.WriteLine("Working on " + round_name + ", " + (round_totalseasons).ToString() + " Seasons!");

        //Goes through every season on the round sheet
        for (int season = 1; season <= (roundPage.Columns().Count() - 1); season += 3)
        {
            //Gets the row for the start of the death log & sets cell locations relative to the season
            IXLCell victimsStart = roundPage.Search("Kill List (include alive)").First();
            var rangeUsed = roundPage.RangeUsed();
            int lastDataRow;
            if (rangeUsed == null)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("ERROR: " + round_name + " is empty!");
                return;
            }
            else
            {
                lastDataRow = rangeUsed.LastRowUsed().RowNumber();
            }
            int firstDataRow = victimsStart.WorksheetRow().RowNumber() + 1;
            int firstDataColumn = season + 1;
            int middleDataColumn = season + 2;
            int lastDataColumn = season + 3;
            String season_debutant = "";

            //Sets variables for stats logic
            List<String> seasonRoster = new List<String>();
            List<String> seasonDebutant = new List<String>();
            List<String> seasonTopKills = new List<String>();
            List<String> seasonWinnerAlive = new List<String>();
            List<String> seasonWinnerDead = new List<String>();
            List<String> seasonAlive = new List<String>();
            List<String> seasonRunnerUps = new List<String>();
            List<String> seasonTeams = new List<String>();
            Dictionary<String, int> rp_teamkills = new Dictionary<String, int>();
            Dictionary<String, int> killboard = new Dictionary<String, int>();
            IXLCell winnerCell = roundPage.Cell(1, 1);
            IXLCell winnerCell2 = roundPage.Cell(1, 1);
            IXLCell lastAliveCell = roundPage.Cell(1, 1);
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
            int crossover_season = 0;
            IXLRange teamRange = roundPage.Range(9, firstDataColumn, firstDataRow - 2, firstDataColumn);


            //Sets round named to be changed for crossovers and ??? to not be called Sheet
            round_name = roundPage.Name;
            if (round_name.Contains("Sheet"))
            {
                round_name = "???";
            }
            String season_number = roundPage.Cell(1, firstDataColumn).GetString();
            DateTime season_date = roundPage.Cell(2, firstDataColumn).GetDateTime();

            //Checks for season date not working chronologically
            if (season > 1)
            {
                if (season_date < roundPage.Cell(2, firstDataColumn).CellLeft(3).GetDateTime())
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine("ERROR: " + round_name + " S" + season_number + " has an invalid date!");
                }
            }

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
                crossover_season = 1;
            }

            if (crossover_season == 0)
            {
                //Get seasons data for the all rosters list
                rostersList.Add(new RostersList(round_name, season_number, season_date));

                //Get seasons data for all rounds except for the non-reddit releases
                if (!roundPage.Cell(1, lastDataColumn).GetString().Equals("NR"))
                {
                    rostersListNR.Add(new RostersList(round_name, season_number, season_date));
                }
            }

            String gamemode = roundPage.Cell(4, firstDataColumn).GetString();
            String[] gamemodeSplit = gamemode.Split(',');
            foreach (String scenario in gamemodeSplit)
            {
                if (gamemodesList.Any(gm => gm.Gamemode == scenario))
                {
                    GamemodesList.UpdateGamemodes(gamemodesList, scenario);
                }
                else
                {
                    gamemodesList.Add(new GamemodesList(scenario, 1));
                }
            }

            //Get the teams for the season
            //Skips FFA seasons since no teams
            teamSize = teamRange.RowsUsed().Count();
            if (!roundPage.Cell(3, middleDataColumn).GetString().Equals("FFA"))
            {
                //Loops the Cells in the team list
                foreach (IXLCell cell in teamRange.CellsUsed())
                {
                    string team = cell.GetString();

                    //Adds the team to the list of the season
                    seasonTeams.Add(team);

                    //Adds the team info the to team list, skips if player is a solo
                    if (crossover_season == 0)
                    {
                        if (team.Contains(","))
                        {
                            teamsList.Add(new TeamsList(round_name, season_number, season_date, team));
                        }
                    }
                }

                if (crossover_season == 0)
                {
                    //Get the team color and adds it to the team list, skips if player is a solo
                    IXLRange teamColorsRange = roundPage.Range(9, lastDataColumn, 9 + (teamSize - 1), lastDataColumn);
                    foreach (IXLCell cell in teamColorsRange.Cells())
                    {
                        String teamColor = cell.GetString();
                        String team = cell.CellLeft(2).GetString();

                        if (team.Contains(","))
                        {
                            if (!teamColor.Equals(""))
                            {
                                var roundTeamColor = teamsList.Find(round => round.Round == round_name && round.Season == season_number && round.Team == team);
                                if (roundTeamColor != null)
                                {
                                    roundTeamColor.TeamColor = teamColor;
                                }
                            }
                        }
                    }
                }
            }

            //Loops through all the victim cells
            IXLRange victimRange = roundPage.Range(firstDataRow, firstDataColumn, lastDataRow, firstDataColumn);
            seasonSize = victimRange.RowsUsed().Count();
            foreach (IXLCell cell in victimRange.CellsUsed())
            {
                string value = cell.GetString();

                if (crossover_season == 0)
                {
                    //Checks if its the players debut round, sets the date if it is
                    //If new players sets all the variables for them
                    if (roundDebutsList.Any(round => round.Player == value))
                    {
                        var roundDebutList = roundDebutsList.Find(round => round.Player == value);
                        if (roundDebutList != null)
                        {
                            if (season_date < roundDebutList.Date)
                            {
                                RoundDebut.UpdateRoundDebut(roundDebutList, round_name, season_number, season_date);
                            }
                        }
                    }
                    else
                    {
                        roundDebutsList.Add(new RoundDebut(round_name, season_number, season_date, value));

                        //Adds new player to the Global Stats
                        globalStatsList.Add(new GlobalStats(value));
                    }

                    //Checks for players debut round but also excludes non-reddit rounds
                    if (roundDebutsListNR.Any(round => round.Player == value))
                    {
                        if (!roundPage.Cell(1, lastDataColumn).GetString().Equals("NR"))
                        {
                            var roundDebutListNR = roundDebutsListNR.Find(round => round.Player == value);
                            if (roundDebutListNR != null)
                            {
                                if (season_date < roundDebutListNR.Date)
                                {
                                    RoundDebut.UpdateRoundDebut(roundDebutListNR, round_name, season_number, season_date);
                                }
                            }
                        }
                    }
                    else
                    {
                        if (!roundPage.Cell(1, lastDataColumn).GetString().Equals("NR"))
                        {
                            roundDebutsListNR.Add(new RoundDebut(round_name, season_number, season_date, value));
                        }
                    }
                }

                //Add Error Messages for suicides.
                if (roundPage.Cell(cell.WorksheetRow().RowNumber(), lastDataColumn).GetString().Equals(value))
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine("ERROR: " + value + " suicided! " + round_name + " " + season_number);
                }

                //If the players didn't die gets +1 alive on the global stats, else +1 death
                if (roundPage.Cell(cell.WorksheetRow().RowNumber(), lastDataColumn).GetString().Equals("Nothing"))
                {
                    seasonAlive.Add(value);
                    if (crossover_season == 0)
                    {
                        GlobalStats.UpdateAlives(globalStatsList, value);

                        //Add to alive list
                        aliveList.Add(new StatsList(round_name, season_number, season_date, value));
                    }
                }
                else
                {
                    if (crossover_season == 0)
                    {
                        GlobalStats.UpdateDeaths(globalStatsList, value);
                    }
                }

                //Makes roster for the season, skips players who show up twice with respawns gamemodes
                //Also adds +1 seasons played for the global stats
                if (!seasonRoster.Contains(value))
                {
                    seasonRoster.Add(value);
                    if (crossover_season == 0)
                    {
                        GlobalStats.UpdateSeasonsPlayed(globalStatsList, value);
                    }
                }

                //Makes roster for the round, adds new players
                //Also adds +1 unique round for the global stats
                if (!roundRoster.Contains(value))
                {
                    roundRoster.Add(value);
                    seasonDebutant.Add(value);
                    GlobalStats.UpdateTotalUniques(globalStatsList, value);
                }
            }
            //Formats the debutants for reddit posts
            seasonAlive.Sort();
            seasonDebutant.Sort();
            seasonRoster.Sort();
            foreach (String debutant in seasonDebutant)
            {
                season_debutant += debutant + ", ";
            }
            if (season_debutant.Length > 0)
            {
                season_debutant = season_debutant.Remove(season_debutant.Length - 2);
                redditPosts.Debutants.Add("**S" + roundPage.Cell(1, firstDataColumn).GetString() + " (" + (season_debutant.Count(c => c == ',') + 1) + "):** " + season_debutant + Environment.NewLine);
            }
            else
            {
                redditPosts.Debutants.Add("**S" + roundPage.Cell(1, firstDataColumn).GetString() + " (" + season_debutant.Count(c => c == ',') + "):** " + Environment.NewLine);
            }

            foreach (String player in seasonRoster)
            {
                String lastPlayed = roundPage.Cell(1, firstDataColumn).GetString();
                String lastSeason = "";
                if (firstDataColumn > 2)
                {
                    lastSeason = roundPage.Cell(1, firstDataColumn).CellLeft(3).GetString();
                }
                else
                {
                    lastSeason = "N/A";
                }
                RedditPostsPlayed.UpdateParticipations(redditPosts, player, lastPlayed, lastSeason);
            }

            if (crossover_season == 0)
            {
                var roundRosterList = rostersList.Find(round => round.Round == round_name && round.Season == season_number);
                if (roundRosterList != null)
                {
                    roundRosterList.Roster = seasonRoster;
                }
                //Adds rosters to a list for the sheet, skips non-reddit for the alternate page
                var roundRosterListNR = rostersListNR.Find(round => round.Round == round_name && round.Season == season_number);
                if (roundRosterListNR != null)
                {
                    roundRosterListNR.Roster = seasonRoster;
                }
            }

            if (!roundPage.Cell(3, middleDataColumn).GetString().Equals("FFA"))
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
            if (roundPage.Cell(firstDataRow, firstDataColumn).GetString().Equals(roundPage.Cell(firstDataRow + 1, lastDataColumn).GetString())
                && roundPage.Cell(firstDataRow + 1, firstDataColumn).GetString().Equals(roundPage.Cell(firstDataRow, lastDataColumn).GetString()))
            {
                redditPosts.FirstDeath.Add("**S" + roundPage.Cell(1, firstDataColumn).GetString() + ":** " + roundPage.Cell(firstDataRow, firstDataColumn).GetString() + " & " + roundPage.Cell(firstDataRow + 1, firstDataColumn).GetString() + " (Double Kill)" + Environment.NewLine);

                if (crossover_season == 0)
                {
                    String player1 = roundPage.Cell(firstDataRow, firstDataColumn).GetString();
                    String player2 = roundPage.Cell(firstDataRow + 1, firstDataColumn).GetString();
                    GlobalStats.UpdateFirstDeaths(globalStatsList, player1);
                    GlobalStats.UpdateFirstDeaths(globalStatsList, player2);

                    //Add to first death list
                    firstDeathList.Add(new StatsList(round_name, season_number, season_date, player1));
                    firstDeathList.Add(new StatsList(round_name, season_number, season_date, player2));
                }
            }
            else
            {
                if (crossover_season == 0)
                {
                    String player = roundPage.Cell(firstDataRow, firstDataColumn).GetString();
                    GlobalStats.UpdateFirstDeaths(globalStatsList, player);

                    //Add to first death list
                    firstDeathList.Add(new StatsList(round_name, season_number, season_date, player));
                }

                //Double the stats for round exception
                if (round_name.Equals("Game Changer")
                    && season_number.Equals("5"))
                {
                    String secondHalf = roundPage.Cell(firstDataRow, firstDataColumn).CellBelow().GetString();
                    GlobalStats.UpdateFirstDeaths(globalStatsList, secondHalf);
                    redditPosts.FirstDeath.Add("**S" + roundPage.Cell(1, firstDataColumn).GetString() + ":** " + roundPage.Cell(firstDataRow, firstDataColumn).GetString() + " & " + roundPage.Cell(firstDataRow, firstDataColumn).CellBelow().GetString() + " (" + roundPage.Cell(firstDataRow, firstDataColumn).CellRight(2).GetString() + ")" + Environment.NewLine);

                    //Add to first death list
                    firstDeathList.Add(new StatsList(round_name, season_number, season_date, secondHalf));
                }
                else
                {
                    if (roundPage.Cell(firstDataRow, firstDataColumn).CellRight(2).GetString().Equals(""))
                    {
                        String method = roundPage.Cell(firstDataRow, firstDataColumn).CellRight(1).GetString();
                        String pvedeath = PveCausesList.GetPveCause(method);
                        redditPosts.FirstDeath.Add("**S" + roundPage.Cell(1, firstDataColumn).GetString() + ":** " + roundPage.Cell(firstDataRow, firstDataColumn).GetString() + " (" + pvedeath + ")" + Environment.NewLine);

                    }
                    else
                    {
                        redditPosts.FirstDeath.Add("**S" + roundPage.Cell(1, firstDataColumn).GetString() + ":** " + roundPage.Cell(firstDataRow, firstDataColumn).GetString() + " (" + roundPage.Cell(firstDataRow, firstDataColumn).CellRight(2).GetString() + ")" + Environment.NewLine);
                    }
                }
            }

            //Gets ironman for the season
            //Different Range for Party Of One since ironman takes 5 rows for that sheet
            IXLRange ironmanRange = roundPage.Range(5, firstDataColumn, 5, lastDataColumn);
            IXLRange POOironmanRange = roundPage.Range(5, firstDataColumn, 9, lastDataColumn);
            String ironman_post = "";
            String ironman_time = "";
            if (round_name.Equals("Party of One"))
            {
                ironman_post = "**S" + roundPage.Cell(1, firstDataColumn).GetString() + ":** ";
                foreach (IXLCell cell in POOironmanRange.CellsUsed())
                {
                    string value = cell.GetString();
                    if (crossover_season == 0)
                    {
                        GlobalStats.UpdateIronmans(globalStatsList, value);

                        //Add to ironman list
                        ironmanList.Add(new StatsList(round_name, season_number, season_date, value));
                    }

                    ironman_post = ironman_post + value + ", ";
                }
                ironman_post = ironman_post.Remove(ironman_post.Length - 2);
                if (roundPage.Cell(10, firstDataColumn).GetString().Equals(""))
                {
                    Console.ForegroundColor = ConsoleColor.Yellow;
                    //Console.WriteLine("ERROR: Ironman time missing! " + round_name + " " + season_number);
                }
                else
                {
                    ironman_time = " (" + roundPage.Cell(10, firstDataColumn).GetString() + ":" + roundPage.Cell(10, middleDataColumn).GetString() + ":" + roundPage.Cell(10, lastDataColumn).GetString() + ")";
                }
                redditPosts.Ironman.Add(ironman_post + ironman_time + Environment.NewLine);
            }
            else
            {
                ironman_post = "**S" + roundPage.Cell(1, firstDataColumn).GetString() + ":** ";
                foreach (IXLCell cell in ironmanRange.CellsUsed())
                {
                    string value = cell.GetString();
                    if (crossover_season == 0)
                    {
                        GlobalStats.UpdateIronmans(globalStatsList, value);

                        //Add to ironman list
                        ironmanList.Add(new StatsList(round_name, season_number, season_date, value));
                    }

                    ironman_post = ironman_post + value + ", ";
                }
                ironman_post = ironman_post.Remove(ironman_post.Length - 2);
                if (roundPage.Cell(6, firstDataColumn).GetString().Equals(""))
                {
                    Console.ForegroundColor = ConsoleColor.Yellow;
                    //Console.WriteLine("ERROR: Ironman time missing! " + round_name + " " + season_number);
                }
                else
                {
                    ironman_time = " (" + roundPage.Cell(6, firstDataColumn).GetString() + ":" + roundPage.Cell(6, middleDataColumn).GetString() + ":" + roundPage.Cell(6, lastDataColumn).GetString() + ")";
                }
                redditPosts.Ironman.Add(ironman_post + ironman_time + Environment.NewLine);
            }

            //Gets first damage for the season
            //Different Range for Party Of One since ironman takes 5 rows for that sheet
            IXLRange fdRange = roundPage.Range(7, firstDataColumn, 7, lastDataColumn);
            IXLRange POOfdRange = roundPage.Range(11, firstDataColumn, 11, lastDataColumn);
            String fd_post = "";
            String fd_time = "";
            if (round_name.Equals("Party of One"))
            {
                fd_post = "**S" + roundPage.Cell(1, firstDataColumn).GetString() + ":** ";
                foreach (IXLCell cell in POOfdRange.CellsUsed())
                {
                    string value = cell.GetString();
                    if (crossover_season == 0)
                    {
                        GlobalStats.UpdateFirstDamages(globalStatsList, value);

                        //Add to first damage list
                        firstDamageList.Add(new StatsList(round_name, season_number, season_date, value));
                    }

                    fd_post = fd_post + value + ", ";
                }

                fd_post = fd_post.Remove(fd_post.Length - 2);
                if (roundPage.Cell(12, firstDataColumn).GetString().Equals(""))
                {
                    Console.ForegroundColor = ConsoleColor.Yellow;
                    //Console.WriteLine("ERROR: First Damage time missing! " + round_name + " " + season_number);
                }
                else
                {
                    fd_time = " (" + roundPage.Cell(12, firstDataColumn).GetString() + ":" + roundPage.Cell(12, middleDataColumn).GetString() + ":" + roundPage.Cell(12, lastDataColumn).GetString() + ")";
                }
                redditPosts.FirstDamage.Add(fd_post + fd_time + Environment.NewLine);
            }
            else
            {
                fd_post = "**S" + roundPage.Cell(1, firstDataColumn).GetString() + ":** ";
                foreach (IXLCell cell in fdRange.CellsUsed())
                {
                    string value = cell.GetString();
                    if (crossover_season == 0)
                    {
                        GlobalStats.UpdateFirstDamages(globalStatsList, value);

                        //Add to first damage list
                        firstDamageList.Add(new StatsList(round_name, season_number, season_date, value));
                    }

                    fd_post = fd_post + value + ", ";
                }

                fd_post = fd_post.Remove(fd_post.Length - 2);
                if (roundPage.Cell(8, firstDataColumn).GetString().Equals(""))
                {
                    Console.ForegroundColor = ConsoleColor.Yellow;
                    //Console.WriteLine("ERROR: First Damage time missing! " + round_name + " " + season_number);
                }
                else
                {
                    fd_time = " (" + roundPage.Cell(8, firstDataColumn).GetString() + ":" + roundPage.Cell(8, middleDataColumn).GetString() + ":" + roundPage.Cell(8, lastDataColumn).GetString() + ")";
                }
                redditPosts.FirstDamage.Add(fd_post + fd_time + Environment.NewLine);
            }

            //Loops through all the kiler cells
            IXLRange killerRange = roundPage.Range(firstDataRow, lastDataColumn, firstDataRow + (seasonSize - 1), lastDataColumn);
            foreach (IXLCell cell in killerRange.Cells())
            {
                String killer = cell.GetString();
                String victim = cell.CellLeft(2).GetString();
                String method = cell.CellLeft().GetString();

                //Checks if killer is PvE or Player
                if (globalStatsList.Any(p => p.Player == killer))
                {
                    //Sets values for the kill list
                    if (crossover_season == 0)
                    {
                        killsList.Add(new KillsList(round_name, season_number, season_date, victim, method, killer));

                        //Adds +1 kill on global stats
                        GlobalStats.UpdateKills(globalStatsList, killer);
                    }

                    RedditPostsKills.UpdateKills(redditPosts, killer, victim, season_number);

                    //Figures out the killboard of the season
                    if (killboard.ContainsKey(killer))
                    {
                        killboard[killer] += 1;
                    }
                    else
                    {
                        killboard.Add(killer, 1);
                    }

                    //Check if there was a double kill for first blood, otherwise gives it to the first player found
                    if (first_blood == 0)
                    {
                        if (killer.Equals(cell.CellBelow().CellLeft(2).GetString())
                            && cell.CellBelow().GetString().Equals(cell.CellLeft(2).GetString()))
                        {
                            first_blood += 2;
                            redditPosts.FirstBlood.Add("**S" + roundPage.Cell(1, firstDataColumn).GetString() + ":** " + killer + " & " + cell.CellBelow().GetString() + " (Double Kill)" + Environment.NewLine);

                            if (crossover_season == 0)
                            {
                                String killer2 = cell.CellBelow().GetString();
                                GlobalStats.UpdateFirstBloods(globalStatsList, killer);
                                GlobalStats.UpdateFirstBloods(globalStatsList, cell.CellBelow().GetString());

                                //Add to first blood list
                                firstBloodList.Add(new StatsList(round_name, season_number, season_date, killer));
                                firstBloodList.Add(new StatsList(round_name, season_number, season_date, killer2));
                            }
                        }
                        else
                        {
                            first_blood += 1;
                            if (crossover_season == 0)
                            {
                                GlobalStats.UpdateFirstBloods(globalStatsList, killer);

                                //Add to first blood list
                                firstBloodList.Add(new StatsList(round_name, season_number, season_date, killer));
                            }

                            //Double the stats for round exception
                            if (round_name.Equals("Game Changer")
                                && season_number.Equals("5"))
                            {
                                String secondHalf = cell.CellBelow().GetString();
                                GlobalStats.UpdateFirstBloods(globalStatsList, secondHalf);
                                redditPosts.FirstBlood.Add("**S" + roundPage.Cell(1, firstDataColumn).GetString() + ":** " + killer + " & " + cell.CellBelow().GetString() + " (" + cell.CellLeft(2).GetString() + " & " + cell.CellBelow().CellLeft(2).GetString() + ")" + Environment.NewLine);

                                //Add to first blood list
                                firstBloodList.Add(new StatsList(round_name, season_number, season_date, secondHalf));
                            }
                            else
                            {
                                redditPosts.FirstBlood.Add("**S" + roundPage.Cell(1, firstDataColumn).GetString() + ":** " + killer + " (" + cell.CellLeft(2).GetString() + ")" + Environment.NewLine);
                            }
                        }
                    }
                }
                else
                {
                    //Adds +1 PvE Death for the player
                    if (!killer.Equals("Nothing"))
                    {
                        if (crossover_season == 0)
                        {
                            String pveVictim = roundPage.Cell(cell.WorksheetRow().RowNumber(), firstDataColumn).GetString();
                            GlobalStats.UpdatePveDeaths(globalStatsList, pveVictim);

                            //Add to pve list
                            pveDeathList.Add(new StatsList(round_name, season_number, season_date, pveVictim));
                        }

                        //Filters all the unique pve deaths
                        if (killer.Equals(""))
                        {
                            String pvedeath = PveCausesList.GetPveCause(method);

                            if (crossover_season == 0)
                            {
                                //Sets values for the kill list
                                killsList.Add(new KillsList(round_name, season_number, season_date, victim, method, pvedeath));

                                if (pveCausesList.Any(p => p.PveCause == pvedeath))
                                {
                                    PveCausesList.UpdatePveCauses(pveCausesList, pvedeath);
                                }
                                else
                                {
                                    pveCausesList.Add(new PveCausesList(pvedeath, 1));
                                }
                            }

                            RedditPostsPve.UpdatePve(redditPosts, pvedeath, victim, season_number);
                        }
                        else
                        {
                            if (crossover_season == 0)
                            {
                                //Sets values for the kill list
                                killsList.Add(new KillsList(round_name, season_number, season_date, victim, method, killer));

                                if (pveCausesList.Any(p => p.PveCause == killer))
                                {
                                    PveCausesList.UpdatePveCauses(pveCausesList, killer);
                                }
                                else
                                {
                                    pveCausesList.Add(new PveCausesList(killer, 1));
                                }
                            }

                            RedditPostsPve.UpdatePve(redditPosts, killer, victim, season_number);
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
                        seasonTopKills.Add(killer);

                        if (crossover_season == 0)
                        {
                            GlobalStats.UpdateTopFrags(globalStatsList, killer);

                            //Add to top frag list
                            topFragList.Add(new StatsList(round_name, season_number, season_date, killer));
                        }
                    }

                    //Checks if the player that got kills beat their kill record
                    if (crossover_season == 0)
                    {
                        if (killRecordsList.Any(p => p.Player == killer))
                        {
                            KillRecords.UpdateKillRecord(killRecordsList, killer, killboard[killer], round_name, season_number, season_date);
                        }
                        else
                        {
                            //If first round with kills sets the kill record
                            killRecordsList.Add(new KillRecords(killer, killboard[killer], round_name, season_number, season_date));
                        }
                    }
                }
                seasonTopKills.Sort();
                String topkills = "";
                foreach (String topkiller in seasonTopKills)
                {
                    topkills = topkills + topkiller + ", ";
                }
                topkills = topkills.Remove(topkills.Length - 2);
                redditPosts.MostKills.Add("**S" + roundPage.Cell(1, firstDataColumn).GetString() + ":** " + topkills + " (" + topFragAmount + ")" + Environment.NewLine);
            }

            if (!roundPage.Cell(3, middleDataColumn).GetString().Equals("FFA"))
            {
                foreach (String team in seasonTeams)
                {
                    String[] team_player = team.Split(separator);

                    foreach (String player in team_player)
                    {
                        if (rp_teamkills.ContainsKey(team))
                        {
                            if (killboard.ContainsKey(player))
                            {
                                rp_teamkills[team] += killboard[player];
                            }
                            else
                            {
                                rp_teamkills[team] += 0;
                            }
                        }
                        else
                        {
                            if (killboard.ContainsKey(player))
                            {
                                rp_teamkills.Add(team, killboard[player]);
                            }
                            else
                            {
                                rp_teamkills.Add(team, 0);
                            }
                        }
                    }
                }
                String most_team_kills = "";
                int teamTopFragAmount = rp_teamkills.Values.Max();
                foreach (String team in rp_teamkills.Keys)
                {
                    if (rp_teamkills[team] == teamTopFragAmount)
                    {
                        String[] team_player = team.Split(separator);

                        foreach (String player in team_player)
                        {
                            if (killboard.ContainsKey(player))
                            {
                                most_team_kills += player + " (" + killboard[player] + "), ";
                            }
                            else
                            {
                                most_team_kills += player + " (0), ";
                            }
                        }

                        if (!most_team_kills.Equals(""))
                        {
                            most_team_kills = most_team_kills.Remove(most_team_kills.Length - 2);
                            most_team_kills += " & ";
                        }
                    }
                }
                most_team_kills = most_team_kills.Remove(most_team_kills.Length - 3);
                redditPosts.MostKillsTeam.Add("**S" + roundPage.Cell(1, firstDataColumn).GetString() + ":** " + most_team_kills + Environment.NewLine);
            }
            else
            {
                redditPosts.MostKillsTeam.Add("**S" + roundPage.Cell(1, firstDataColumn).GetString() + ":** " + "N/A" + Environment.NewLine);
            }

            //Get the winners of the season
            //If Nothing is a regular season ending and gives the win to the last player on the list
            //Else is either a double kill win or no wins and is figured out to give the wins needed
            if (roundPage.Cell(firstDataRow + (seasonSize - 1), lastDataColumn).GetString().Equals("Nothing"))
            {
                //Gets the last season winner
                String seasonWinner = roundPage.Cell(firstDataRow + (seasonSize - 1), firstDataColumn).GetString();
                winnerCell = roundPage.Cell(firstDataRow + (seasonSize - 1), firstDataColumn);

                if (roundPage.Cell(4, firstDataColumn).GetString().Contains("Dragon Rush") ||
                    roundPage.Cell(4, firstDataColumn).GetString().Contains("Wither Rush") ||
                    roundPage.Cell(4, firstDataColumn).GetString().Contains("Realm Rush") ||
                    roundPage.Cell(4, firstDataColumn).GetString().Contains("Bolas Rush") ||
                    roundPage.Cell(4, firstDataColumn).GetString().Contains("Escape From Gaia") ||
                    roundPage.Cell(4, firstDataColumn).GetString().Contains("Trouble In Paradise") ||
                    roundPage.Cell(4, firstDataColumn).GetString().Contains("Dragon Rush Deviation Version") ||
                    roundPage.Cell(4, firstDataColumn).GetString().Contains("Hydra Rush"))
                {
                    IXLCell dragonRushCell = roundPage.Cell(firstDataRow + (seasonSize - 1), lastDataColumn);
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

                lastAliveCell = roundPage.Cell(firstDataRow + (seasonSize - 1), lastDataColumn);
                while (lastAliveCell.CellAbove().GetString().Equals("Nothing"))
                {
                    lastAliveCell = lastAliveCell.CellAbove();
                }

                //If FFA no need to look for teams, else looks for the team
                if (roundPage.Cell(3, middleDataColumn).GetString().Equals("FFA"))
                {
                    if (seasonAlive.Contains(seasonWinner))
                    {
                        seasonWinnerAlive.Add(seasonWinner);
                    }
                    else
                    {
                        seasonWinnerDead.Add(seasonWinner);
                    }

                    if (crossover_season == 0)
                    {
                        GlobalStats.UpdateWins(globalStatsList, seasonWinner);

                        //Add to winner list
                        winList.Add(new StatsList(round_name, season_number, season_date, seasonWinner));
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
                                if (seasonAlive.Contains(winner))
                                {
                                    seasonWinnerAlive.Add(winner);
                                }
                                else
                                {
                                    seasonWinnerDead.Add(winner);
                                }

                                if (crossover_season == 0)
                                {
                                    GlobalStats.UpdateWins(globalStatsList, winner);

                                    //Add to winner list
                                    winList.Add(new StatsList(round_name, season_number, season_date, winner));
                                }
                            }
                        }
                    }
                }

                //Detects double kill runner ups
                if (lastAliveCell.CellAbove().GetString().Equals(lastAliveCell.CellLeft(2).CellAbove(2).GetString())
                    && lastAliveCell.CellAbove().CellAbove().GetString().Equals(lastAliveCell.CellLeft(2).CellAbove().GetString()))
                {
                    if (roundPage.Cell(3, middleDataColumn).GetString().Equals("FFA"))
                    {
                        double_kill_runnerup = 1;
                    }
                    else
                    {
                        if (!winningTeam.Contains(lastAliveCell.CellAbove().GetString()) && !winningTeam.Contains(lastAliveCell.CellAbove(2).GetString()))
                        {
                            double_kill_runnerup = 1;
                        }
                    }
                }
            }
            else
            {
                //Check for a double kill ending
                if (roundPage.Cell(firstDataRow + (seasonSize - 1), lastDataColumn).GetString().Equals(roundPage.Cell(firstDataRow + (seasonSize - 2), firstDataColumn).GetString())
                    && roundPage.Cell(firstDataRow + (seasonSize - 2), lastDataColumn).GetString().Equals(roundPage.Cell(firstDataRow + (seasonSize - 1), firstDataColumn).GetString()))
                {
                    //Double kill ending so 2 winners
                    double_kill_ending = 1;
                    String seasonWinner1 = roundPage.Cell(firstDataRow + (seasonSize - 1), firstDataColumn).GetString();
                    String seasonWinner2 = roundPage.Cell(firstDataRow + (seasonSize - 2), firstDataColumn).GetString();
                    winnerCell = roundPage.Cell(firstDataRow + (seasonSize - 1), firstDataColumn);
                    winnerCell2 = roundPage.Cell(firstDataRow + (seasonSize - 2), firstDataColumn);

                    //If FFA no need to look for teams, else looks for the team
                    if (roundPage.Cell(3, middleDataColumn).GetString().Equals("FFA"))
                    {
                        if (seasonAlive.Contains(seasonWinner1))
                        {
                            seasonWinnerAlive.Add(seasonWinner1);
                        }
                        else
                        {
                            seasonWinnerDead.Add(seasonWinner1);
                        }

                        if (seasonAlive.Contains(seasonWinner2))
                        {
                            seasonWinnerAlive.Add(seasonWinner2);
                        }
                        else
                        {
                            seasonWinnerDead.Add(seasonWinner2);
                        }

                        if (crossover_season == 0)
                        {
                            GlobalStats.UpdateWins(globalStatsList, seasonWinner1);
                            GlobalStats.UpdateWins(globalStatsList, seasonWinner2);

                            //Add to winner list
                            winList.Add(new StatsList(round_name, season_number, season_date, seasonWinner1));
                            winList.Add(new StatsList(round_name, season_number, season_date, seasonWinner2));
                        }
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
                                    if (seasonAlive.Contains(winner))
                                    {
                                        seasonWinnerAlive.Add(winner);
                                    }
                                    else
                                    {
                                        seasonWinnerDead.Add(winner);
                                    }

                                    if (crossover_season == 0)
                                    {
                                        GlobalStats.UpdateWins(globalStatsList, winner);

                                        //Add to winner list
                                        winList.Add(new StatsList(round_name, season_number, season_date, winner));
                                    }
                                }
                            }
                            if (team.Contains(seasonWinner2))
                            {
                                winningTeam2 = team;

                                //Splits the team string to get each player and gives them a win
                                String[] winners = team.Split(separator);
                                foreach (String winner in winners)
                                {
                                    if (seasonAlive.Contains(winner))
                                    {
                                        seasonWinnerAlive.Add(winner);
                                    }
                                    else
                                    {
                                        seasonWinnerDead.Add(winner);
                                    }

                                    if (crossover_season == 0)
                                    {
                                        GlobalStats.UpdateWins(globalStatsList, winner);

                                        //Add to winner list
                                        winList.Add(new StatsList(round_name, season_number, season_date, winner));
                                    }
                                }
                            }
                        }
                    }
                }
                else
                {
                    dragon_win = 1;

                    seasonWinnerDead.Add("Ender Dragon");
                }
            }
            String winner_post = " ";
            seasonWinnerAlive.Sort();
            seasonWinnerDead.Sort();
            foreach (String winner in seasonWinnerAlive)
            {
                if (killboard.ContainsKey(winner))
                {
                    winner_post += winner + " (" + killboard[winner] + "), ";
                }
                else
                {
                    winner_post += winner + " (0), ";
                }
            }
            if (seasonWinnerDead.Count == 0)
            {
                winner_post = winner_post.Remove(winner_post.Length - 2);
                winner_post = winner_post + "***";
            }
            else
            {
                if (!winner_post.Equals(" "))
                {
                    winner_post = winner_post.Remove(winner_post.Length - 2);
                    winner_post = winner_post + "**, *";
                }
                else
                {
                    winner_post = "** *";
                }
            }
            foreach (String winner in seasonWinnerDead)
            {
                if (killboard.ContainsKey(winner))
                {
                    winner_post += winner + " (" + killboard[winner] + "), ";
                }
                else
                {
                    if (winner.Equals("Ender Dragon"))
                    {
                        winner_post += winner + ", ";
                    }
                    else
                    {
                        winner_post += winner + " (0), ";
                    }
                }
            }
            winner_post = winner_post.Remove(winner_post.Length - 2);
            winner_post = winner_post + "*";
            redditPosts.Winners.Add("**S" + roundPage.Cell(1, firstDataColumn).GetString() + ":" + winner_post + Environment.NewLine);

            //Get the runner ups of the season
            //If FFA it has to be the player above
            //Else figures out the next team after the winners
            if (dragon_win == 0)
            {
                if (dragon_rush_ru == 0)
                {
                    if (double_kill_runnerup == 0)
                    {
                        if (roundPage.Cell(3, middleDataColumn).GetString().Equals("FFA"))
                        {
                            if (double_kill_ending == 1)
                            {
                                seasonRunnerUps.Add(winnerCell2.CellAbove().GetString());

                                if (crossover_season == 0)
                                {
                                    GlobalStats.UpdateRunnerUps(globalStatsList, winnerCell2.CellAbove().GetString());

                                    //Add to runner up list
                                    runnerUpList.Add(new StatsList(round_name, season_number, season_date, winnerCell2.CellAbove().GetString()));
                                }
                            }
                            else
                            {
                                seasonRunnerUps.Add(winnerCell.CellAbove().GetString());

                                if (crossover_season == 0)
                                {
                                    GlobalStats.UpdateRunnerUps(globalStatsList, winnerCell.CellAbove().GetString());

                                    //Add to runner up list
                                    runnerUpList.Add(new StatsList(round_name, season_number, season_date, winnerCell.CellAbove().GetString()));
                                }
                            }
                        }
                        else
                        {
                            if (double_kill_ending == 1)
                            {

                                while (winningTeam.Contains(winnerCell2.CellAbove().GetString()) || winningTeam2.Contains(winnerCell2.CellAbove().GetString()))
                                {
                                    winnerCell2 = winnerCell2.CellAbove();
                                }

                                //Figures out the full team of runner ups
                                String seasonRunnerUp = winnerCell2.CellAbove().GetString();
                                foreach (String team in seasonTeams)
                                {
                                    if (team.Contains(seasonRunnerUp))
                                    {
                                        //Splits the team string to get each player and gives them a runner up
                                        String[] runnerups = team.Split(separator);
                                        foreach (String runner_up in runnerups)
                                        {
                                            seasonRunnerUps.Add(runner_up);

                                            if (crossover_season == 0)
                                            {
                                                GlobalStats.UpdateRunnerUps(globalStatsList, runner_up);

                                                //Add to runner up list
                                                runnerUpList.Add(new StatsList(round_name, season_number, season_date, runner_up));
                                            }
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
                                            seasonRunnerUps.Add(runner_up);

                                            if (crossover_season == 0)
                                            {
                                                GlobalStats.UpdateRunnerUps(globalStatsList, runner_up);

                                                //Add to runner up list
                                                runnerUpList.Add(new StatsList(round_name, season_number, season_date, runner_up));
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                    else
                    {
                        if (roundPage.Cell(3, middleDataColumn).GetString().Equals("FFA"))
                        {
                            seasonRunnerUps.Add(winnerCell.CellAbove().GetString());
                            seasonRunnerUps.Add(winnerCell.CellAbove(2).GetString());

                            if (crossover_season == 0)
                            {
                                GlobalStats.UpdateRunnerUps(globalStatsList, winnerCell.CellAbove().GetString());
                                GlobalStats.UpdateRunnerUps(globalStatsList, winnerCell.CellAbove(2).GetString());

                                //Add to runner up list
                                runnerUpList.Add(new StatsList(round_name, season_number, season_date, winnerCell.CellAbove().GetString()));
                                runnerUpList.Add(new StatsList(round_name, season_number, season_date, winnerCell.CellAbove(2).GetString()));
                            }
                        }
                        else
                        {
                            while (winningTeam.Contains(winnerCell.CellAbove().GetString()))
                            {
                                winnerCell = winnerCell.CellAbove();
                            }

                            //Figures out the full team of runner ups
                            String seasonRunnerUp = winnerCell.CellAbove().GetString();
                            String seasonRunnerUp2 = winnerCell.CellAbove(2).GetString();
                            foreach (String team in seasonTeams)
                            {
                                if (team.Contains(seasonRunnerUp))
                                {
                                    //Splits the team string to get each player and gives them a runner up
                                    String[] runnerups = team.Split(separator);
                                    foreach (String runner_up in runnerups)
                                    {
                                        seasonRunnerUps.Add(runner_up);

                                        if (crossover_season == 0)
                                        {
                                            GlobalStats.UpdateRunnerUps(globalStatsList, runner_up);

                                            //Add to runner up list
                                            runnerUpList.Add(new StatsList(round_name, season_number, season_date, runner_up));
                                        }
                                    }
                                }

                                if (team.Contains(seasonRunnerUp2))
                                {
                                    //Splits the team string to get each player and gives them a runner up
                                    String[] runnerups = team.Split(separator);
                                    foreach (String runner_up in runnerups)
                                    {
                                        seasonRunnerUps.Add(runner_up);

                                        if (crossover_season == 0)
                                        {
                                            GlobalStats.UpdateRunnerUps(globalStatsList, runner_up);

                                            //Add to runner up list
                                            runnerUpList.Add(new StatsList(round_name, season_number, season_date, runner_up));
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                else
                {
                    IXLCell runnerUpCheck = roundPage.Cell(firstDataRow + (seasonSize - 1), lastDataColumn);
                    IXLCell runnerUpPlayer = roundPage.Cell(firstDataRow + (seasonSize - 1), firstDataColumn);

                    while (runnerUpCheck.GetString().Equals("Nothing"))
                    {
                        if (!runnerUpCheck.CellLeft().GetString().Equals("Winner"))
                        {
                            seasonRunnerUps.Add(runnerUpPlayer.GetString());

                            if (crossover_season == 0)
                            {
                                GlobalStats.UpdateRunnerUps(globalStatsList, runnerUpPlayer.GetString());

                                runnerUpList.Add(new StatsList(round_name, season_number, season_date, runnerUpPlayer.GetString()));
                            }
                        }
                        runnerUpCheck = runnerUpCheck.CellAbove();
                        runnerUpPlayer = runnerUpPlayer.CellAbove();
                    }
                }
            }
            else
            {
                //Dragon wins the season
                String seasonRunnerUp = roundPage.Cell(firstDataRow + (seasonSize - 1), firstDataColumn).GetString();

                if (roundPage.Cell(3, middleDataColumn).GetString().Equals("FFA"))
                {
                    seasonRunnerUps.Add(seasonRunnerUp);

                    if (crossover_season == 0)
                    {
                        GlobalStats.UpdateRunnerUps(globalStatsList, seasonRunnerUp);

                        runnerUpList.Add(new StatsList(round_name, season_number, season_date, seasonRunnerUp));
                    }
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
                                seasonRunnerUps.Add(runner_up);

                                if (crossover_season == 0)
                                {
                                    GlobalStats.UpdateRunnerUps(globalStatsList, runner_up);

                                    //Add to runner up list
                                    runnerUpList.Add(new StatsList(round_name, season_number, season_date, runner_up));
                                }
                            }
                        }
                    }
                }

            }

            String runnerup_post = "";
            seasonRunnerUps.Sort();
            foreach (String runner_up in seasonRunnerUps)
            {
                if (killboard.ContainsKey(runner_up))
                {
                    runnerup_post += runner_up + " (" + killboard[runner_up] + "), ";
                }
                else
                {
                    runnerup_post += runner_up + " (0), ";
                }
            }
            if (seasonTeams.Count == 1 && !seasonWinnerDead.First().Equals("Ender Dragon"))
            {
                redditPosts.RunnerUps.Add("**S" + roundPage.Cell(1, firstDataColumn).GetString() + ":** " + "N/A" + Environment.NewLine);
            }
            else
            {
                runnerup_post = runnerup_post.Remove(runnerup_post.Length - 2);
                redditPosts.RunnerUps.Add("**S" + roundPage.Cell(1, firstDataColumn).GetString() + ":** " + runnerup_post + Environment.NewLine);
            }
        }

        var roundRosterSize = roundList.Find(round => round.Round == round_name);
        if (roundRosterSize != null)
        {
            roundRosterSize.RosterSize = roundRoster.Count;
        }
        String rppath = "..\\..\\..\\Reddit Posts\\" + postFolder + "\\" + post_name + ".txt";
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
                            "191st", "192nd", "193rd", "194th", "195th", "196th", "197th", "198th", "199th", "200th",
                            "201st", "202nd", "203rd", "204th", "205th", "206th", "207th", "208th", "209th", "210th",
                            "211th", "212th", "213th", "214th", "215th", "216th", "217th", "218th", "219th", "220th",
                            "221st", "222nd", "223rd", "224th", "225th", "226th", "227th", "228th", "229th", "230th",
                            "231st", "232nd", "233rd", "234th", "235th", "236th", "237th", "238th", "239th", "240th",
                            "241st", "242nd", "243rd", "244th", "245th", "246th", "247th", "248th", "249th", "250th",
                            "251st", "252nd", "253rd", "254th", "255th", "256th", "257th", "258th", "259th", "260th",
                            "261st", "262nd", "263rd", "264th", "265th", "266th", "267th", "268th", "269th", "270th",
                            "271st", "272nd", "273rd", "274th", "275th", "276th", "277th", "278th", "279th", "280th",
                            "281st", "282nd", "283rd", "284th", "285th", "286th", "287th", "288th", "289th", "290th",
                            "291st", "292nd", "293rd", "294th", "295th", "296th", "297th", "298th", "299th", "300th",};
        int ranking = 0;
        int ties = 1;
        int currentkill = 0;

        File.WriteAllText(rppath, "## " + post_name + " Statistics" + Environment.NewLine);
        File.AppendAllText(rppath, Environment.NewLine + "---");
        File.AppendAllText(rppath, Environment.NewLine + "### Winners" + Environment.NewLine + Environment.NewLine);
        foreach (String winner in redditPosts.Winners)
        {
            File.AppendAllText(rppath, winner + Environment.NewLine);
        }
        File.AppendAllText(rppath, "---");
        File.AppendAllText(rppath, Environment.NewLine + "### Runner Ups" + Environment.NewLine + Environment.NewLine);
        foreach (String runner_up in redditPosts.RunnerUps)
        {
            File.AppendAllText(rppath, runner_up + Environment.NewLine);
        }
        File.AppendAllText(rppath, "---");
        File.AppendAllText(rppath, Environment.NewLine + "### Most Kills" + Environment.NewLine + Environment.NewLine);
        foreach (String topkiller in redditPosts.MostKills)
        {
            File.AppendAllText(rppath, topkiller + Environment.NewLine);
        }
        File.AppendAllText(rppath, "---");
        File.AppendAllText(rppath, Environment.NewLine + "### Most Kills (Team)" + Environment.NewLine + Environment.NewLine);
        foreach (String topkillerteam in redditPosts.MostKillsTeam)
        {
            File.AppendAllText(rppath, topkillerteam + Environment.NewLine);
        }
        File.AppendAllText(rppath, "---");
        File.AppendAllText(rppath, Environment.NewLine + "### First Damage" + Environment.NewLine + Environment.NewLine);
        foreach (String firstdamages in redditPosts.FirstDamage)
        {
            File.AppendAllText(rppath, firstdamages + Environment.NewLine);
        }
        File.AppendAllText(rppath, "---");
        File.AppendAllText(rppath, Environment.NewLine + "### Ironman" + Environment.NewLine + Environment.NewLine);
        foreach (String ironmans in redditPosts.Ironman)
        {
            File.AppendAllText(rppath, ironmans + Environment.NewLine);
        }
        File.AppendAllText(rppath, "---");
        File.AppendAllText(rppath, Environment.NewLine + "### First Blood" + Environment.NewLine + Environment.NewLine);
        foreach (String firstbloods in redditPosts.FirstBlood)
        {
            File.AppendAllText(rppath, firstbloods + Environment.NewLine);
        }
        File.AppendAllText(rppath, "---");
        File.AppendAllText(rppath, Environment.NewLine + "### First Death" + Environment.NewLine + Environment.NewLine);
        foreach (String firstdeaths in redditPosts.FirstDeath)
        {
            File.AppendAllText(rppath, firstdeaths + Environment.NewLine);
        }
        File.AppendAllText(rppath, "---");
        File.AppendAllText(rppath, Environment.NewLine + "### Kills" + Environment.NewLine + Environment.NewLine);
        foreach (RedditPostsKills kills in redditPosts.Kills)
        {
            kills.KillsList = kills.KillsList.Remove(kills.KillsList.Length - 2);
        }
        redditPosts.Kills = redditPosts.Kills.OrderByDescending(x => x.KillsAmount).ThenBy(x => x.Player).ToList();
        foreach (RedditPostsKills kills in redditPosts.Kills)
        {
            if (currentkill > 0)
            {
                if (kills.KillsAmount == currentkill)
                {
                    ties += 1;
                }
                else
                {
                    ranking += ties;
                    ties = 1;
                }
            }
            File.AppendAllText(rppath, "**" + placement[ranking] + " - " + kills.Player + " (" + kills.KillsAmount + "):** " + kills.KillsList + Environment.NewLine + Environment.NewLine);
            currentkill = kills.KillsAmount;
        }
        File.AppendAllText(rppath, "---");
        File.AppendAllText(rppath, Environment.NewLine + "### PvE Deaths" + Environment.NewLine + Environment.NewLine);
        foreach (RedditPostsPve pveDeaths in redditPosts.PveDeaths)
        {
            pveDeaths.DeathsList = pveDeaths.DeathsList.Remove(pveDeaths.DeathsList.Length - 2);
        }
        redditPosts.PveDeaths = redditPosts.PveDeaths.OrderByDescending(x => x.DeathsAmount).ThenBy(x => x.PveCause).ToList();
        foreach (RedditPostsPve pveDeaths in redditPosts.PveDeaths)
        {
            File.AppendAllText(rppath, "**" + pveDeaths.PveCause + " (" + pveDeaths.DeathsAmount + "):** " + pveDeaths.DeathsList + Environment.NewLine + Environment.NewLine);
        }
        File.AppendAllText(rppath, "---");
        File.AppendAllText(rppath, Environment.NewLine + "### Participation" + Environment.NewLine + Environment.NewLine);
        for (int seasons = (roundPage.Columns().Count() - 1) / 3; seasons > 0; seasons--)
        {
            String season_part = "";
            foreach (RedditPostsPlayed player in redditPosts.Participations)
            {
                if (player.TotalPlayed == seasons)
                {
                    season_part += player.Player + " " + player.SeasonsPlayed + ")" + ", ";
                }
            }
            int count = season_part.Count(c => c == ' ') / 2;
            if (count > 0)
            {
                season_part = season_part.Remove(season_part.Length - 2);
            }
            if (seasons == 1)
            {
                File.AppendAllText(rppath, "**" + seasons.ToString() + " Season (" + count.ToString() + "):** " + season_part + Environment.NewLine + Environment.NewLine);
            }
            else
            {
                File.AppendAllText(rppath, "**" + seasons.ToString() + " Seasons (" + count.ToString() + "):** " + season_part + Environment.NewLine + Environment.NewLine);
            }
        }
        File.AppendAllText(rppath, "---");
        File.AppendAllText(rppath, Environment.NewLine + "### Debutants" + Environment.NewLine + Environment.NewLine);
        foreach (String debutants in redditPosts.Debutants)
        {
            File.AppendAllText(rppath, debutants + Environment.NewLine);
        }
        File.AppendAllText(rppath, "---");
    }

    //Updates KDR & KPR of player
    //If infinite sets it to the amount of kills they have
    foreach (GlobalStats globalStats in globalStatsList)
    {
        if (globalStats.Deaths.Equals(0))
        {
            globalStats.KDR = Convert.ToDouble(globalStats.Kills);
        }
        else
        {
            globalStats.KDR = Convert.ToDouble(globalStats.Kills) / Convert.ToDouble(globalStats.Deaths);
        }
        globalStats.KPR = Convert.ToDouble(globalStats.Kills) / Convert.ToDouble(globalStats.SeasonsPlayed);
    }

    //Takes all the lists and adds them to new docs to save
    var roundlist = new XLWorkbook();
    var statscompiled = new XLWorkbook();
    var rrdebut = new XLWorkbook();
    var globalstats = new XLWorkbook();

    //Making Round List Page
    var round_list = roundlist.AddWorksheet("Round List");
    round_list.Column("D").Style.NumberFormat.Format = "dd mmm, yyyy";
    round_list.Cell(1, 1).InsertData(roundList);
    round_list.Sort(4, XLSortOrder.Ascending);

    //Making PvE List Page
    var pve_list = roundlist.AddWorksheet("PvE List");
    pve_list.Cell(1, 1).InsertData(pveCausesList);
    pve_list.Sort(2, XLSortOrder.Descending);

    //Making PvE List Page
    var gm_list = roundlist.AddWorksheet("Gamemode List");
    gm_list.Cell(1, 1).InsertData(gamemodesList);
    gm_list.Sort(2, XLSortOrder.Descending);

    //Making All Rosters Page
    var allrosters = roundlist.AddWorksheet("All Rosters");
    allrosters.Column("C").Style.NumberFormat.Format = "mm/dd/yyyy";
    int currentRow = 1;
    foreach (var round in rostersList)
    {
        allrosters.Cell(currentRow, 1).Value = round.Round;
        allrosters.Cell(currentRow, 2).Value = round.Season;
        allrosters.Cell(currentRow, 3).Value = round.Date;
        allrosters.Cell(currentRow, 4).InsertData(round.Roster, transpose: true);
        currentRow++;
    }
    allrosters.Sort(3, XLSortOrder.Ascending);

    //Making NR All Rosters Page
    var allrosters_nr = roundlist.AddWorksheet("All Rosters (NR)");
    allrosters_nr.Column("C").Style.NumberFormat.Format = "mm/dd/yyyy";
    int currentRowNR = 1;
    foreach (var round in rostersListNR)
    {
        allrosters_nr.Cell(currentRowNR, 1).Value = round.Round;
        allrosters_nr.Cell(currentRowNR, 2).Value = round.Season;
        allrosters_nr.Cell(currentRowNR, 3).Value = round.Date;
        allrosters_nr.Cell(currentRowNR, 4).InsertData(round.Roster, transpose: true);
        currentRowNR++;
    }
    allrosters_nr.Sort(3, XLSortOrder.Ascending);

    //Making Kills Page
    var allkills = statscompiled.AddWorksheet("All Kills");
    allkills.Column("C").Style.NumberFormat.Format = "mm/dd/yyyy";
    allkills.Cell(1, 1).InsertData(killsList);
    allkills.Sort(3, XLSortOrder.Ascending);

    //Making Teams Page
    var allteams = statscompiled.AddWorksheet("All Teams");
    allteams.Column("C").Style.NumberFormat.Format = "mm/dd/yyyy";
    allteams.Cell(1, 1).InsertData(teamsList);
    allteams.Sort(3, XLSortOrder.Ascending);

    //First Damage list
    var firstdamage = statscompiled.AddWorksheet("First Damage");
    firstdamage.Column("C").Style.NumberFormat.Format = "mm/dd/yyyy";
    firstdamage.Cell(1, 1).InsertData(firstDamageList);
    firstdamage.Sort(3, XLSortOrder.Ascending);

    //Ironman list
    var ironman = statscompiled.AddWorksheet("Ironman");
    ironman.Column("C").Style.NumberFormat.Format = "mm/dd/yyyy";
    ironman.Cell(1, 1).InsertData(ironmanList);
    ironman.Sort(3, XLSortOrder.Ascending);

    //PvE Death list
    var pve_death = statscompiled.AddWorksheet("PvE Deaths");
    pve_death.Column("C").Style.NumberFormat.Format = "mm/dd/yyyy";
    pve_death.Cell(1, 1).InsertData(pveDeathList);
    pve_death.Sort(3, XLSortOrder.Ascending);

    //First Death list
    var firstdeath = statscompiled.AddWorksheet("First Death");
    firstdeath.Column("C").Style.NumberFormat.Format = "mm/dd/yyyy";
    firstdeath.Cell(1, 1).InsertData(firstDeathList);
    firstdeath.Sort(3, XLSortOrder.Ascending);

    //First Blood list
    var firstblood = statscompiled.AddWorksheet("First Blood");
    firstblood.Column("C").Style.NumberFormat.Format = "mm/dd/yyyy";
    firstblood.Cell(1, 1).InsertData(firstBloodList);
    firstblood.Sort(3, XLSortOrder.Ascending);

    //Most Kills list
    var topfrags = statscompiled.AddWorksheet("Top Frags");
    topfrags.Column("C").Style.NumberFormat.Format = "mm/dd/yyyy";
    topfrags.Cell(1, 1).InsertData(topFragList);
    topfrags.Sort(3, XLSortOrder.Ascending);

    //Runner Up list
    var runnerup = statscompiled.AddWorksheet("Runner Ups");
    runnerup.Column("C").Style.NumberFormat.Format = "mm/dd/yyyy";
    runnerup.Cell(1, 1).InsertData(runnerUpList);
    runnerup.Sort(3, XLSortOrder.Ascending);

    //Alive list
    var alive = statscompiled.AddWorksheet("Alive");
    alive.Column("C").Style.NumberFormat.Format = "mm/dd/yyyy";
    alive.Cell(1, 1).InsertData(aliveList);
    alive.Sort(3, XLSortOrder.Ascending);

    //Win list
    var win = statscompiled.AddWorksheet("Wins");
    win.Column("C").Style.NumberFormat.Format = "mm/dd/yyyy";
    win.Cell(1, 1).InsertData(winList);
    win.Sort(3, XLSortOrder.Ascending);

    //RR Debuts
    var rr_debut = rrdebut.AddWorksheet("RR Debuts");
    rr_debut.Column("C").Style.NumberFormat.Format = "mm/dd/yyyy";
    rr_debut.Cell(1, 1).InsertData(roundDebutsList);
    rr_debut.Sort(3, XLSortOrder.Ascending);

    //RR Debuts (No NR)
    var rr_debut_nr = rrdebut.AddWorksheet("RR Debuts (No NR)");
    rr_debut_nr.Column("C").Style.NumberFormat.Format = "mm/dd/yyyy";
    rr_debut_nr.Cell(1, 1).InsertData(roundDebutsListNR);
    rr_debut_nr.Sort(3, XLSortOrder.Ascending);

    //Global Stats
    var global_stats = globalstats.AddWorksheet("Global Stats");
    global_stats.Cell(1, 1).InsertData(globalStatsList);
    global_stats.Sort(1, XLSortOrder.Ascending);

    //Kill Records
    var kill_records = globalstats.AddWorksheet("Kill Records");
    kill_records.Cell(1, 1).InsertData(killRecordsList);
    kill_records.Sort(1, XLSortOrder.Ascending);

    //Saves the new docs
    roundlist.SaveAs("..\\..\\..\\Stats Sheet\\" + postFolder + "\\Round_List.xlsx");
    statscompiled.SaveAs("..\\..\\..\\Stats Sheet\\" + postFolder + "\\Stats_Compiled.xlsx");
    globalstats.SaveAs("..\\..\\..\\Stats Sheet\\" + postFolder + "\\Global_Stats.xlsx");
    rrdebut.SaveAs("..\\..\\..\\Stats Sheet\\" + postFolder + "\\RR_Debuts.xlsx");
    Console.ForegroundColor = ConsoleColor.Green;
    Console.WriteLine("Stats are now compiled!");
}