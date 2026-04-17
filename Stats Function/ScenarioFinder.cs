using ClosedXML.Excel;

namespace StatsAnalyzer
{
    public class ScenarioFinder
    {
        public static void GetScenarios(List<GamemodesList> gamemodesList, List<TeamTypeList> teamTypeList, SeasonInfo seasonInfo)
        {
            if (seasonInfo.IsCrossoverSeason == false)
            {
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

                String[] teamTypeSplit = seasonInfo.SeasonTeamType.Split(',');
                foreach (String scenario in teamTypeSplit)
                {
                    if (teamTypeList.Any(gm => gm.TeamType == scenario))
                    {
                        TeamTypeList.UpdateTeamType(teamTypeList, scenario);
                    }
                    else
                    {
                        teamTypeList.Add(new TeamTypeList(scenario, 1));
                    }
                }
            }
        }
    }
}