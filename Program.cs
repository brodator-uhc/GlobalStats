using ClosedXML.Excel;
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
        String roundName = roundPage.Name;
        int roundTotalSeasons = (roundPage.Columns().Count() - 1) / 3;
        DateTime round_debutdate = roundPage.Cell(2, 2).GetDateTime();
        String redditPostName = "";
        if (roundName.Contains("Sheet"))
        {
            roundName = "???";
            redditPostName = "3 Question Marks";
        }
        else
        {
            redditPostName = roundName;
        }
        roundList.Add(new RoundList(roundName, roundTotalSeasons, round_debutdate));
        Console.ForegroundColor = ConsoleColor.Green;
        Console.WriteLine("Working on " + roundName + ", " + roundTotalSeasons.ToString() + " Seasons!");

        //Goes through every season on the round sheet
        for (int season = 1; season <= (roundPage.Columns().Count() - 1); season += 3)
        {
            //Gets the row for the start of the death log & sets cell locations relative to the season
            int firstDataColumn = season + 1;
            int middleDataColumn = season + 2;
            int lastDataColumn = season + 3;
            SeasonInfo seasonInfo = new SeasonInfo(roundPage, firstDataColumn);
            IXLCell victimsStart = roundPage.Search("Kill List (include alive)").First();
            var rangeUsed = roundPage.RangeUsed();
            int lastDataRow;
            if (rangeUsed == null)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("ERROR: " + seasonInfo.SeasonName + " is empty!");
                return;
            }
            else
            {
                lastDataRow = rangeUsed.LastRowUsed().RowNumber();
            }
            int firstDataRow = victimsStart.WorksheetRow().RowNumber() + 1;
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
            WinnerInfo winnerInfo = new WinnerInfo(roundPage);
            int seasonSize = 0;
            int teamSize = 0;
            int first_blood = 0;
            IXLRange teamRange = roundPage.Range(9, firstDataColumn, firstDataRow - 2, firstDataColumn);

            //Checks for season date not working chronologically
            if (season > 1)
            {
                if (seasonInfo.SeasonDate < roundPage.Cell(2, firstDataColumn).CellLeft(3).GetDateTime())
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine("ERROR: " + seasonInfo.SeasonName + " S" + seasonInfo.SeasonNumber + " has an invalid date!");
                }
            }

            if (seasonInfo.IsCrossoverSeason == false)
            {
                //Get seasons data for the all rosters list
                rostersList.Add(new RostersList(seasonInfo.SeasonName, seasonInfo.SeasonNumber, seasonInfo.SeasonDate));

                //Get seasons data for all rounds except for the non-reddit releases
                if (!roundPage.Cell(1, lastDataColumn).GetString().Equals("NR"))
                {
                    rostersListNR.Add(new RostersList(seasonInfo.SeasonName, seasonInfo.SeasonNumber, seasonInfo.SeasonDate));
                }

                //Get gamemode list
                String[] gamemodeSplit = seasonInfo.SeasonGamemodes.Split(',');
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
            }

            //Get the teams for the season
            //Skips FFA seasons since no teams
            teamSize = teamRange.RowsUsed().Count();
            if (seasonInfo.IsFFA == false)
            {
                //Loops the Cells in the team list
                foreach (IXLCell cell in teamRange.CellsUsed())
                {
                    string team = cell.GetString();

                    //Adds the team to the list of the season
                    seasonTeams.Add(team);

                    //Adds the team info the to team list, skips if player is a solo
                    if (seasonInfo.IsCrossoverSeason == false)
                    {
                        if (team.Contains(","))
                        {
                            teamsList.Add(new TeamsList(seasonInfo.SeasonName, seasonInfo.SeasonNumber, seasonInfo.SeasonDate, team));
                        }
                    }
                }

                if (seasonInfo.IsCrossoverSeason == false)
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
                                var roundTeamColor = teamsList.Find(round => round.Round == seasonInfo.SeasonName && round.Season == seasonInfo.SeasonNumber && round.Team == team);
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
            winnerInfo.LastDataRowCell = roundPage.Cell(firstDataRow + (seasonSize - 1), firstDataColumn);
            foreach (IXLCell cell in victimRange.CellsUsed())
            {
                string value = cell.GetString();

                if (seasonInfo.IsCrossoverSeason == false)
                {
                    //Checks if its the players debut round, sets the date if it is
                    //If new players sets all the variables for them
                    if (roundDebutsList.Any(round => round.Player == value))
                    {
                        var roundDebutList = roundDebutsList.Find(round => round.Player == value);
                        if (roundDebutList != null)
                        {
                            if (seasonInfo.SeasonDate < roundDebutList.Date)
                            {
                                RoundDebut.UpdateRoundDebut(roundDebutList, seasonInfo.SeasonName, seasonInfo.SeasonNumber, seasonInfo.SeasonDate);
                            }
                        }
                    }
                    else
                    {
                        roundDebutsList.Add(new RoundDebut(seasonInfo.SeasonName, seasonInfo.SeasonNumber, seasonInfo.SeasonDate, value));

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
                                if (seasonInfo.SeasonDate < roundDebutListNR.Date)
                                {
                                    RoundDebut.UpdateRoundDebut(roundDebutListNR, seasonInfo.SeasonName, seasonInfo.SeasonNumber, seasonInfo.SeasonDate);
                                }
                            }
                        }
                    }
                    else
                    {
                        if (!roundPage.Cell(1, lastDataColumn).GetString().Equals("NR"))
                        {
                            roundDebutsListNR.Add(new RoundDebut(seasonInfo.SeasonName, seasonInfo.SeasonNumber, seasonInfo.SeasonDate, value));
                        }
                    }
                }

                //Add Error Messages for suicides.
                if (roundPage.Cell(cell.WorksheetRow().RowNumber(), lastDataColumn).GetString().Equals(value))
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine("ERROR: " + value + " suicided! " + seasonInfo.SeasonName + " " + seasonInfo.SeasonNumber);
                }

                //If the players didn't die gets +1 alive on the global stats, else +1 death
                if (roundPage.Cell(cell.WorksheetRow().RowNumber(), lastDataColumn).GetString().Equals("Nothing"))
                {
                    seasonAlive.Add(value);
                    if (seasonInfo.IsCrossoverSeason == false)
                    {
                        GlobalStats.UpdateAlives(globalStatsList, value);

                        //Add to alive list
                        aliveList.Add(new StatsList(seasonInfo, value));
                    }
                }
                else
                {
                    if (seasonInfo.IsCrossoverSeason == false)
                    {
                        GlobalStats.UpdateDeaths(globalStatsList, value);
                    }
                }

                //Makes roster for the season, skips players who show up twice with respawns gamemodes
                //Also adds +1 seasons played for the global stats
                if (!seasonRoster.Contains(value))
                {
                    seasonRoster.Add(value);
                    if (seasonInfo.IsCrossoverSeason == false)
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

            if (seasonInfo.IsCrossoverSeason == false)
            {
                var roundRosterList = rostersList.Find(round => round.Round == seasonInfo.SeasonName && round.Season == seasonInfo.SeasonNumber);
                if (roundRosterList != null)
                {
                    roundRosterList.Roster = seasonRoster;
                }
                //Adds rosters to a list for the sheet, skips non-reddit for the alternate page
                var roundRosterListNR = rostersListNR.Find(round => round.Round == seasonInfo.SeasonName && round.Season == seasonInfo.SeasonNumber);
                if (roundRosterListNR != null)
                {
                    roundRosterListNR.Roster = seasonRoster;
                }
            }

            if (seasonInfo.IsFFA == false)
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
                        Console.WriteLine("ERROR: Player " + player + " missing in teams! " + seasonInfo.SeasonName + " " + seasonInfo.SeasonNumber);
                    }

                    if (playercheck > 1)
                    {
                        Console.ForegroundColor = ConsoleColor.Red;
                        Console.WriteLine("ERROR: Player " + player + " duplicate in teams! " + seasonInfo.SeasonName + " " + seasonInfo.SeasonNumber);
                    }

                    playercheck = 0;
                }

                //Verify if player is misspelled or missing in victims
                int teamcheck = 0;
                foreach (String team in seasonTeams)
                {
                    String[] teamplayers = team.Split(',');

                    foreach (String teamplayer in teamplayers)
                    {
                        if (seasonRoster.Contains(teamplayer))
                        {
                            teamcheck += 1;
                        }

                        if (teamcheck == 0)
                        {
                            Console.ForegroundColor = ConsoleColor.Red;
                            Console.WriteLine("ERROR: Player " + teamplayer + " missing in victims! " + seasonInfo.SeasonName + " " + seasonInfo.SeasonNumber);
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

                if (seasonInfo.IsCrossoverSeason == false)
                {
                    String player1 = roundPage.Cell(firstDataRow, firstDataColumn).GetString();
                    String player2 = roundPage.Cell(firstDataRow + 1, firstDataColumn).GetString();
                    GlobalStats.UpdateFirstDeaths(globalStatsList, player1);
                    GlobalStats.UpdateFirstDeaths(globalStatsList, player2);

                    //Add to first death list
                    firstDeathList.Add(new StatsList(seasonInfo, player1));
                    firstDeathList.Add(new StatsList(seasonInfo, player2));
                }
            }
            else
            {
                if (seasonInfo.IsCrossoverSeason == false)
                {
                    String player = roundPage.Cell(firstDataRow, firstDataColumn).GetString();
                    GlobalStats.UpdateFirstDeaths(globalStatsList, player);

                    //Add to first death list
                    firstDeathList.Add(new StatsList(seasonInfo, player));
                }

                //Double the stats for round exception
                if (seasonInfo.SeasonName.Equals("Game Changer")
                    && seasonInfo.SeasonNumber.Equals("5"))
                {
                    String secondHalf = roundPage.Cell(firstDataRow, firstDataColumn).CellBelow().GetString();
                    GlobalStats.UpdateFirstDeaths(globalStatsList, secondHalf);
                    redditPosts.FirstDeath.Add("**S" + roundPage.Cell(1, firstDataColumn).GetString() + ":** " + roundPage.Cell(firstDataRow, firstDataColumn).GetString() + " & " + roundPage.Cell(firstDataRow, firstDataColumn).CellBelow().GetString() + " (" + roundPage.Cell(firstDataRow, firstDataColumn).CellRight(2).GetString() + ")" + Environment.NewLine);

                    //Add to first death list
                    firstDeathList.Add(new StatsList(seasonInfo, secondHalf));
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
            if (seasonInfo.SeasonName.Equals("Party of One"))
            {
                ironman_post = "**S" + roundPage.Cell(1, firstDataColumn).GetString() + ":** ";
                foreach (IXLCell cell in POOironmanRange.CellsUsed())
                {
                    string value = cell.GetString();
                    if (seasonInfo.IsCrossoverSeason == false)
                    {
                        GlobalStats.UpdateIronmans(globalStatsList, value);

                        //Add to ironman list
                        ironmanList.Add(new StatsList(seasonInfo, value));
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
                    if (seasonInfo.IsCrossoverSeason == false)
                    {
                        GlobalStats.UpdateIronmans(globalStatsList, value);

                        //Add to ironman list
                        ironmanList.Add(new StatsList(seasonInfo, value));
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
            if (seasonInfo.SeasonName.Equals("Party of One"))
            {
                fd_post = "**S" + roundPage.Cell(1, firstDataColumn).GetString() + ":** ";
                foreach (IXLCell cell in POOfdRange.CellsUsed())
                {
                    string value = cell.GetString();
                    if (seasonInfo.IsCrossoverSeason == false)
                    {
                        GlobalStats.UpdateFirstDamages(globalStatsList, value);

                        //Add to first damage list
                        firstDamageList.Add(new StatsList(seasonInfo, value));
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
                    if (seasonInfo.IsCrossoverSeason == false)
                    {
                        GlobalStats.UpdateFirstDamages(globalStatsList, value);

                        //Add to first damage list
                        firstDamageList.Add(new StatsList(seasonInfo, value));
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
                    if (seasonInfo.IsCrossoverSeason == false)
                    {
                        killsList.Add(new KillsList(seasonInfo.SeasonName, seasonInfo.SeasonNumber, seasonInfo.SeasonDate, victim, method, killer));

                        //Adds +1 kill on global stats
                        GlobalStats.UpdateKills(globalStatsList, killer);
                    }

                    RedditPostsKills.UpdateKills(redditPosts, killer, victim, seasonInfo.SeasonNumber);

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

                            if (seasonInfo.IsCrossoverSeason == false)
                            {
                                String killer2 = cell.CellBelow().GetString();
                                GlobalStats.UpdateFirstBloods(globalStatsList, killer);
                                GlobalStats.UpdateFirstBloods(globalStatsList, cell.CellBelow().GetString());

                                //Add to first blood list
                                firstBloodList.Add(new StatsList(seasonInfo, killer));
                                firstBloodList.Add(new StatsList(seasonInfo, killer2));
                            }
                        }
                        else
                        {
                            first_blood += 1;
                            if (seasonInfo.IsCrossoverSeason == false)
                            {
                                GlobalStats.UpdateFirstBloods(globalStatsList, killer);

                                //Add to first blood list
                                firstBloodList.Add(new StatsList(seasonInfo, killer));
                            }

                            //Double the stats for round exception
                            if (seasonInfo.SeasonName.Equals("Game Changer")
                                && seasonInfo.SeasonNumber.Equals("5"))
                            {
                                String secondHalf = cell.CellBelow().GetString();
                                GlobalStats.UpdateFirstBloods(globalStatsList, secondHalf);
                                redditPosts.FirstBlood.Add("**S" + roundPage.Cell(1, firstDataColumn).GetString() + ":** " + killer + " & " + cell.CellBelow().GetString() + " (" + cell.CellLeft(2).GetString() + " & " + cell.CellBelow().CellLeft(2).GetString() + ")" + Environment.NewLine);

                                //Add to first blood list
                                firstBloodList.Add(new StatsList(seasonInfo, secondHalf));
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
                        if (seasonInfo.IsCrossoverSeason == false)
                        {
                            String pveVictim = roundPage.Cell(cell.WorksheetRow().RowNumber(), firstDataColumn).GetString();
                            GlobalStats.UpdatePveDeaths(globalStatsList, pveVictim);

                            //Add to pve list
                            pveDeathList.Add(new StatsList(seasonInfo, pveVictim));
                        }

                        //Filters all the unique pve deaths
                        if (killer.Equals(""))
                        {
                            String pvedeath = PveCausesList.GetPveCause(method);

                            if (seasonInfo.IsCrossoverSeason == false)
                            {
                                //Sets values for the kill list
                                killsList.Add(new KillsList(seasonInfo.SeasonName, seasonInfo.SeasonNumber, seasonInfo.SeasonDate, victim, method, pvedeath));

                                if (pveCausesList.Any(p => p.PveCause == pvedeath))
                                {
                                    PveCausesList.UpdatePveCauses(pveCausesList, pvedeath);
                                }
                                else
                                {
                                    pveCausesList.Add(new PveCausesList(pvedeath, 1));
                                }
                            }

                            RedditPostsPve.UpdatePve(redditPosts, pvedeath, victim, seasonInfo.SeasonNumber);
                        }
                        else
                        {
                            if (seasonInfo.IsCrossoverSeason == false)
                            {
                                //Sets values for the kill list
                                killsList.Add(new KillsList(seasonInfo.SeasonName, seasonInfo.SeasonNumber, seasonInfo.SeasonDate, victim, method, killer));

                                if (pveCausesList.Any(p => p.PveCause == killer))
                                {
                                    PveCausesList.UpdatePveCauses(pveCausesList, killer);
                                }
                                else
                                {
                                    pveCausesList.Add(new PveCausesList(killer, 1));
                                }
                            }

                            RedditPostsPve.UpdatePve(redditPosts, killer, victim, seasonInfo.SeasonNumber);
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

                        if (seasonInfo.IsCrossoverSeason == false)
                        {
                            GlobalStats.UpdateTopFrags(globalStatsList, killer);

                            //Add to top frag list
                            topFragList.Add(new StatsList(seasonInfo, killer));
                        }
                    }

                    //Checks if the player that got kills beat their kill record
                    if (seasonInfo.IsCrossoverSeason == false)
                    {
                        if (killRecordsList.Any(p => p.Player == killer))
                        {
                            KillRecords.UpdateKillRecord(killRecordsList, killer, killboard[killer], seasonInfo.SeasonName, seasonInfo.SeasonNumber, seasonInfo.SeasonDate);
                        }
                        else
                        {
                            //If first round with kills sets the kill record
                            killRecordsList.Add(new KillRecords(killer, killboard[killer], seasonInfo.SeasonName, seasonInfo.SeasonNumber, seasonInfo.SeasonDate));
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

            if (seasonInfo.IsFFA == false)
            {
                foreach (String team in seasonTeams)
                {
                    String[] team_player = team.Split(',');

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
                        String[] team_player = team.Split(',');

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
            String seasonNumberPost = roundPage.Cell(1, firstDataColumn).GetString();
            WinnerFinder.GetWinners(globalStatsList, winList, seasonInfo, winnerInfo, seasonAlive, seasonTeams, seasonWinnerAlive, seasonWinnerDead);
            RedditPostFormat.FormatWins(redditPosts, seasonInfo, seasonTeams, seasonWinnerAlive, seasonWinnerDead, killboard, seasonNumberPost);

            //Get the runner ups of the season
            //If FFA it has to be the player above
            //Else figures out the next team after the winners
            RunnerUpFinder.GetRunnerUps(globalStatsList, runnerUpList, seasonInfo, winnerInfo, seasonRunnerUps, seasonTeams);
            RedditPostFormat.FormatRunnerUps(redditPosts, seasonInfo, seasonRunnerUps, seasonTeams, seasonWinnerDead, killboard);
        }

        //Updates the roster size of the round
        RoundList.UpdateRosterSize(roundList, roundName, roundRoster.Count);

        //Save Reddit Post of the round into text file (Markdown formatting).
        RedditPostCompiler.SaveRedditPost(redditPosts, roundTotalSeasons, postFolder, redditPostName);
    }

    //Updates KDR & KPR of player
    //If infinite sets it to the amount of kills they have
    GlobalStats.UpdateKDRs(globalStatsList);

    //Takes all the lists and adds them to new docs to save
    DataExporter.SaveRoundList(roundList, pveCausesList, gamemodesList, rostersList, rostersListNR, postFolder);
    DataExporter.SaveRoundDebut(roundDebutsList, roundDebutsListNR, postFolder);
    DataExporter.SaveGlobalStats(globalStatsList, killRecordsList, postFolder);
    DataExporter.SaveCompiledStats(killsList, teamsList, firstDamageList, ironmanList, pveDeathList, firstDeathList, firstBloodList, topFragList, runnerUpList, aliveList, winList, postFolder);

    Console.ForegroundColor = ConsoleColor.Green;
    Console.WriteLine("Stats are now compiled!");
}