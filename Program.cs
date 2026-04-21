using ClosedXML.Excel;
using StatsAnalyzer;

// If 1 calculates list of round gaps
// If 2 calculates list of global gaps
// If 3 makes a personal stats sheet for a player
// If 4 calculates the global stats
int Statfunction = 1;
// Player to analyze for the stats sheet
String playerStats = "Kaismartypants";
// Select the stat doc for global stats
// 1 for reddit
// 2 for non-reddit
// 3 for live rounds
int statDoc = 1;

if (Statfunction == 1)
{
    String roundsListDoc = "..\\..\\..\\Stats Sheet\\Reddit\\Round_List.xlsx";
    List<RoundsGaps> roundsGapList = [];

    if (!File.Exists(roundsListDoc))
    {
        Console.ForegroundColor = ConsoleColor.Red;
        Console.WriteLine($"Error: File not found at {roundsListDoc}");
        return;
    }

    using var roundsDocument = new XLWorkbook(roundsListDoc);
    var roundList = roundsDocument.Worksheet(1);
    var rosterList = roundsDocument.Worksheet(5);

    //Calculate the global gaps between 2 season of 1 roudn
    RoundGapsCalculator.CalculateRoundGaps(roundsGapList, roundList, rosterList);

    //Saves the new docs
    DataExporter.SaveRoundsGaps(roundsGapList);
}
else if (Statfunction == 2)
{
    String playerListDoc = "..\\..\\..\\Stats Sheet\\Reddit\\Global_Stats.xlsx";
    String roundsListDoc = "..\\..\\..\\Stats Sheet\\Reddit\\Round_List.xlsx";
    List<GlobalGaps> globalGapList = [];

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

    using var roundsDocument = new XLWorkbook(roundsListDoc);
    var rosterList = roundsDocument.Worksheet(5);

    //Calculate the global gaps between 2 season played of any round
    GlobalGapsCalculator.CalculateGlobalGaps(globalGapList, playerList, rosterList);

    //Saves the gap doc
    DataExporter.SaveGlobalGaps(globalGapList);
}
else if (Statfunction == 3)
{
    //IGN of the player to analyze
    String roundsListDoc = "..\\..\\..\\Stats Sheet\\Reddit\\Round_List.xlsx";
    String statsListDoc = "..\\..\\..\\Stats Sheet\\Reddit\\Stats_Compiled.xlsx";
    List<PlayerStats> playerStatsList = [];

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
    var rosterList = roundsDocument.Worksheet(5);

    using var statsDocument = new XLWorkbook(statsListDoc);

    //Calculate the players stats for each rounds played
    PlayerStatsCalculator.CalculatePlayerStats(playerStatsList, rosterList, statsDocument, playerStats);

    //Saves the new stat doc
    DataExporter.SavePlayerStats(playerStatsList);

    //Deletes temporary filled cells
    String filePath = "..\\..\\..\\Stats Sheet\\Player Stats.xlsx";
    DataExporter.ClearEmptyCells(filePath);

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
            RedditPostFormat.FormatWins(redditPosts, seasonInfo, winnerInfo, seasonLists, killboard, seasonNumberPost);

            //Get the runner ups of the season
            //If FFA it has to be the player above
            //Else figures out the next team after the winners
            RunnerUpFinder.GetRunnerUps(globalStatsList, runnerUpList, seasonInfo, winnerInfo, seasonLists);
            RedditPostFormat.FormatRunnerUps(redditPosts, seasonInfo, winnerInfo, seasonLists, killboard, seasonNumberPost);
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