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

    //Select the stats doc to use
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
    List<RoundList> roundList = [];
    List<PveCausesList> pveCausesList = [];
    List<GamemodesList> gamemodesList = [];
    List<TeamTypeList> teamTypeList = [];
    List<RostersList> rostersList = [];
    List<RostersList> rostersListNR = [];
    //Variables for the round debut list
    List<RoundDebut> roundDebutsList = [];
    List<RoundDebut> roundDebutsListNR = [];
    //Variables for the global stats
    List<GlobalStats> globalStatsList = [];
    List<KillRecords> killRecordsList = [];
    //Variables for the stats list
    List<KillsList> killsList = [];
    List<TeamsList> teamsList = [];
    List<StatsList> firstDamageList = [];
    List<StatsList> ironmanList = [];
    List<StatsList> pveDeathList = [];
    List<StatsList> firstDeathList = [];
    List<StatsList> firstBloodList = [];
    List<StatsList> topFragList = [];
    List<StatsList> runnerUpList = [];
    List<StatsList> aliveList = [];
    List<StatsList> winList = [];

    //Goes through every stats tabs on the doc
    using var globalDocument = new XLWorkbook(filePath);
    for (int sheet = skipSheet; sheet <= globalDocument.Worksheets.Count; sheet++)
    {
        List<String> roundRoster = [];
        RedditPosts redditPosts = new();

        //Collecting the Name, the total amount of seasons & the date of S1 for the round
        var roundPage = globalDocument.Worksheet(sheet);
        String roundName = roundPage.Name;
        int roundTotalSeasons = (roundPage.Columns().Count() - 1) / 3;
        DateTime roundDebutDate = roundPage.Cell(2, 2).GetDateTime();
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
        roundList.Add(new RoundList(roundName, roundTotalSeasons, roundDebutDate));

        Console.ForegroundColor = ConsoleColor.Green;
        Console.WriteLine("Working on " + roundName + ", " + roundTotalSeasons.ToString() + " Seasons!");

        //Goes through every season on the round sheet
        for (int season = 1; season <= (roundPage.Columns().Count() - 1); season += 3)
        {
            //Gets the row for the start of the death log & sets cell locations relative to the season
            //Also Sets variables for stats logic
            int firstDataColumn = season + 1;
            SeasonInfo seasonInfo = new(roundPage, firstDataColumn);
            SeasonLists seasonLists = new();
            WinnerInfo winnerInfo = new(roundPage);
            Dictionary<String, int> killboard = [];
            Dictionary<String, int> teamKillboard = [];
            String seasonNumberPost = roundPage.Cell(1, firstDataColumn).GetString();
            IXLCell victimsStart = roundPage.Search("Kill List (include alive)").First();
            int firstDataRow = victimsStart.WorksheetRow().RowNumber() + 1;
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

            //Checks for season date not working chronologically
            if (season > 1)
            {
                IXLCell lastSeasonDate = roundPage.Cell(2, firstDataColumn).CellLeft(3);
                StatsVerification.VerifyDate(seasonInfo, lastSeasonDate);
            }

            //Get the gamemode and team type list for the season
            ScenarioFinder.GetScenarios(gamemodesList, teamTypeList, seasonInfo);

            //Get the teams for the season
            //Skips FFA seasons since no teams
            IXLRange teamRange = roundPage.Range(9, firstDataColumn, firstDataRow - 2, firstDataColumn);
            seasonInfo.SeasonTeamSize = teamRange.RowsUsed().Count();
            IXLRange teamColorsRange = roundPage.Range(9, firstDataColumn + 2, 9 + (seasonInfo.SeasonTeamSize - 1), firstDataColumn + 2);
            TeamsAnalyzer.GetTeams(teamsList, seasonLists, seasonInfo, teamRange, teamColorsRange);

            //Loops through all the victim cells
            IXLCell seasonNumberCell = roundPage.Cell(1, firstDataColumn);
            IXLRange victimRange = roundPage.Range(firstDataRow, firstDataColumn, lastDataRow, firstDataColumn);
            DeathsAnalyzer.GetDeaths(globalStatsList, redditPosts, rostersList, rostersListNR, roundDebutsList, roundDebutsListNR, aliveList, roundRoster, seasonLists, seasonInfo, victimRange, seasonNumberPost, firstDataColumn, seasonNumberCell);

            //Verify IGNs for players in the deaths and teams for errors
            StatsVerification.VerifyPlayers(seasonInfo, seasonLists);

            //Figures out if there is a double kill for first death, otherwise add +1 to the first death
            IXLCell firstDeathVictim = roundPage.Cell(firstDataRow, firstDataColumn);
            IXLCell firstDeathKiller = roundPage.Cell(firstDataRow, firstDataColumn + 2);
            FirstDeathFinder.GetFirstDeath(globalStatsList, redditPosts, firstDeathList, seasonInfo, seasonNumberPost, firstDeathVictim, firstDeathKiller);

            //Gets ironman for the season
            //Different Range for Party Of One since ironman takes 5 rows for that sheet
            if (seasonInfo.SeasonName.Equals("Party of One"))
            {
                IXLRange ironmanRange = roundPage.Range(5, firstDataColumn, 9, firstDataColumn + 2);
                IXLCell ironmanTimeCell = roundPage.Cell(10, firstDataColumn);
                IronmanFinder.GetIronman(globalStatsList, redditPosts, ironmanList, seasonInfo, seasonNumberPost, ironmanRange, ironmanTimeCell);
            }
            else
            {
                IXLRange ironmanRange = roundPage.Range(5, firstDataColumn, 5, firstDataColumn + 2);
                IXLCell ironmanTimeCell = roundPage.Cell(6, firstDataColumn);
                IronmanFinder.GetIronman(globalStatsList, redditPosts, ironmanList, seasonInfo, seasonNumberPost, ironmanRange, ironmanTimeCell);
            }

            //Gets first damage for the season
            //Different Range for Party Of One since ironman takes 5 rows for that sheet
            if (seasonInfo.SeasonName.Equals("Party of One"))
            {
                IXLRange firstDamageRange = roundPage.Range(11, firstDataColumn, 11, firstDataColumn + 2);
                IXLCell firstDamageTimeCell = roundPage.Cell(12, firstDataColumn);
                FirstDamageFinder.GetFirstDamage(globalStatsList, redditPosts, firstDamageList, seasonInfo, seasonNumberPost, firstDamageRange, firstDamageTimeCell);
            }
            else
            {
                IXLRange firstDamageRange = roundPage.Range(7, firstDataColumn, 7, firstDataColumn + 2);
                IXLCell firstDamageTimeCell = roundPage.Cell(8, firstDataColumn);
                FirstDamageFinder.GetFirstDamage(globalStatsList, redditPosts, firstDamageList, seasonInfo, seasonNumberPost, firstDamageRange, firstDamageTimeCell);
            }

            //Loops through all the kiler cells
            IXLRange killerRange = roundPage.Range(firstDataRow, firstDataColumn + 2, firstDataRow + (seasonInfo.SeasonSize - 1), firstDataColumn + 2);
            KillsAnalyzer.GetKills(globalStatsList, redditPosts, killsList, firstBloodList, pveDeathList, pveCausesList, seasonInfo, killboard, seasonNumberPost, killerRange);

            //Gets top frags for the season
            //Skips rounds with 0 kills.
            KillboardAnalyzer.GetMostKills(globalStatsList, redditPosts, topFragList, killRecordsList, seasonInfo, seasonLists, killboard, teamKillboard, seasonNumberPost);
            RedditPostFormat.FormatTeamTopFrags(redditPosts, seasonInfo, killboard, teamKillboard, seasonNumberPost);

            //Get the winners of the season
            //If Nothing is a regular season ending and gives the win to the last player on the list
            //Else is either a double kill win or no wins and is figured out to give the wins needed
            winnerInfo.LastDataRowCell = roundPage.Cell(firstDataRow + (seasonInfo.SeasonSize - 1), firstDataColumn);
            WinnerFinder.GetWinners(globalStatsList, winList, seasonInfo, winnerInfo, seasonLists);
            RedditPostFormat.FormatWins(redditPosts, seasonLists, killboard, seasonNumberPost);

            //Get the runner ups of the season
            //If FFA it has to be the player above
            //Else figures out the next team after the winners
            RunnerUpFinder.GetRunnerUps(globalStatsList, runnerUpList, seasonInfo, winnerInfo, seasonLists);
            RedditPostFormat.FormatRunnerUps(redditPosts, seasonLists, killboard, seasonNumberPost);
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
    DataExporter.SaveRoundList(roundList, pveCausesList, gamemodesList, teamTypeList, rostersList, rostersListNR, postFolder);
    DataExporter.SaveRoundDebut(roundDebutsList, roundDebutsListNR, postFolder);
    DataExporter.SaveGlobalStats(globalStatsList, killRecordsList, postFolder);
    DataExporter.SaveCompiledStats(killsList, teamsList, firstDamageList, ironmanList, pveDeathList, firstDeathList, firstBloodList, topFragList, runnerUpList, aliveList, winList, postFolder);

    Console.ForegroundColor = ConsoleColor.Green;
    Console.WriteLine("Stats are now compiled!");
}