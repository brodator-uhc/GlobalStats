using ClosedXML.Excel;

namespace StatsAnalyzer
{
    public class TeamsAnalyzer
    {
        public static void GetTeams(List<TeamsList> teamsList, SeasonLists seasonLists, SeasonInfo seasonInfo, IXLRange teamRange, IXLRange teamColorsRange)
        {
            if (seasonInfo.IsFFA == false)
            {
                //Loops the Cells in the team list
                foreach (IXLCell teamCell in teamRange.CellsUsed())
                {
                    string team = teamCell.GetString();
                    seasonLists.SeasonTeams.Add(team);

                    //Adds the team info the to team list, skips if player is a solo
                    if (seasonInfo.IsCrossoverSeason == false)
                    {
                        if (team.Contains(','))
                        {
                            teamsList.Add(new TeamsList(seasonInfo, team));
                        }
                    }
                }

                if (seasonInfo.IsCrossoverSeason == false)
                {
                    //Get the team color and adds it to the team list, skips if player is a solo
                    foreach (IXLCell teamColorCell in teamColorsRange.Cells())
                    {
                        String teamColor = teamColorCell.GetString();
                        String team = teamColorCell.CellLeft(2).GetString();

                        if (team.Contains(','))
                        {
                            if (!teamColor.Equals(""))
                            {
                                var roundTeamColor = teamsList.Find(r => r.Round == seasonInfo.SeasonName && r.Season == seasonInfo.SeasonNumber && r.Team == team);
                                roundTeamColor?.TeamColor = teamColor;
                            }
                        }
                    }
                }
            }
        }
    }
}