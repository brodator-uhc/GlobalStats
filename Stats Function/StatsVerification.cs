using ClosedXML.Excel;

namespace StatsAnalyzer
{
    public class StatsVerification
    {
        public static void VerifyDate(SeasonInfo seasonInfo, IXLCell lastSeasonDate)
        {
                if (seasonInfo.SeasonDate < lastSeasonDate.GetDateTime())
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine("ERROR: " + seasonInfo.SeasonName + " S" + seasonInfo.SeasonNumber + " has an invalid date!");
                }
        }
        public static void VerifyPlayers(SeasonInfo seasonInfo, SeasonLists seasonLists)
        {
            if (seasonInfo.IsFFA == false)
            {
                //Verify if player is misspelled or missing in teams
                int playerCheck = 0;
                foreach (String player in seasonLists.SeasonRoster)
                {
                    foreach (String team in seasonLists.SeasonTeams)
                    {
                        if (team.Contains(player))
                        {
                            playerCheck += 1;
                        }
                    }

                    if (playerCheck == 0)
                    {
                        Console.ForegroundColor = ConsoleColor.Red;
                        Console.WriteLine("ERROR: Player " + player + " missing in teams! " + seasonInfo.SeasonName + " " + seasonInfo.SeasonNumber);
                    }

                    if (playerCheck > 1)
                    {
                        Console.ForegroundColor = ConsoleColor.Red;
                        Console.WriteLine("ERROR: Player " + player + " duplicate in teams! " + seasonInfo.SeasonName + " " + seasonInfo.SeasonNumber);
                    }

                    playerCheck = 0;
                }

                //Verify if player is misspelled or missing in victims
                int teamCheck = 0;
                foreach (String team in seasonLists.SeasonTeams)
                {
                    String[] teamMembers = team.Split(',');

                    foreach (String player in teamMembers)
                    {
                        if (seasonLists.SeasonRoster.Contains(player))
                        {
                            teamCheck += 1;
                        }

                        if (teamCheck == 0)
                        {
                            Console.ForegroundColor = ConsoleColor.Red;
                            Console.WriteLine("ERROR: Player " + player + " missing in victims! " + seasonInfo.SeasonName + " " + seasonInfo.SeasonNumber);
                        }

                        teamCheck = 0;
                    }
                }
            }
        }
    }
}