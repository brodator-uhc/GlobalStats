namespace StatsAnalyzer
{
    public class RedditPostCompiler
    {
        public static void SaveRedditPost(RedditPosts redditPosts, int totalSeasons, String postFolder, String redditPostName)
        {
            String rppath = "..\\..\\..\\Reddit Posts\\" + postFolder + "\\" + redditPostName + ".txt";
            String[] placement = [ "1st", "2nd", "3rd", "4th", "5th", "6th", "7th", "8th", "9th", "10th",
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
                            "291st", "292nd", "293rd", "294th", "295th", "296th", "297th", "298th", "299th", "300th"];

            //Winners
            File.WriteAllText(rppath, "## " + redditPostName + " Statistics" + Environment.NewLine);
            File.AppendAllText(rppath, Environment.NewLine + "---");
            File.AppendAllText(rppath, Environment.NewLine + "### Winners" + Environment.NewLine + Environment.NewLine);
            foreach (String winner in redditPosts.Winners)
            {
                File.AppendAllText(rppath, winner + Environment.NewLine);
            }
            File.AppendAllText(rppath, "---");

            //Runner Ups
            File.AppendAllText(rppath, Environment.NewLine + "### Runner Ups" + Environment.NewLine + Environment.NewLine);
            foreach (String runnerUp in redditPosts.RunnerUps)
            {
                File.AppendAllText(rppath, runnerUp + Environment.NewLine);
            }
            File.AppendAllText(rppath, "---");

            //Most Kills
            File.AppendAllText(rppath, Environment.NewLine + "### Most Kills" + Environment.NewLine + Environment.NewLine);
            foreach (String mostKill in redditPosts.MostKills)
            {
                File.AppendAllText(rppath, mostKill + Environment.NewLine);
            }
            File.AppendAllText(rppath, "---");

            //Most Kills (Team)
            File.AppendAllText(rppath, Environment.NewLine + "### Most Kills (Team)" + Environment.NewLine + Environment.NewLine);
            foreach (String mostKillTeam in redditPosts.MostKillsTeam)
            {
                File.AppendAllText(rppath, mostKillTeam + Environment.NewLine);
            }
            File.AppendAllText(rppath, "---");

            //First Damage
            File.AppendAllText(rppath, Environment.NewLine + "### First Damage" + Environment.NewLine + Environment.NewLine);
            foreach (String firstDamage in redditPosts.FirstDamage)
            {
                File.AppendAllText(rppath, firstDamage + Environment.NewLine);
            }
            File.AppendAllText(rppath, "---");

            //Ironman
            File.AppendAllText(rppath, Environment.NewLine + "### Ironman" + Environment.NewLine + Environment.NewLine);
            foreach (String ironman in redditPosts.Ironman)
            {
                File.AppendAllText(rppath, ironman + Environment.NewLine);
            }
            File.AppendAllText(rppath, "---");

            //First Blood
            File.AppendAllText(rppath, Environment.NewLine + "### First Blood" + Environment.NewLine + Environment.NewLine);
            foreach (String firstBlood in redditPosts.FirstBlood)
            {
                File.AppendAllText(rppath, firstBlood + Environment.NewLine);
            }
            File.AppendAllText(rppath, "---");

            //First Death
            File.AppendAllText(rppath, Environment.NewLine + "### First Death" + Environment.NewLine + Environment.NewLine);
            foreach (String firstDeath in redditPosts.FirstDeath)
            {
                File.AppendAllText(rppath, firstDeath + Environment.NewLine);
            }
            File.AppendAllText(rppath, "---");

            //Kills
            int ranking = 0;
            int ties = 1;
            int currentKill = 0;
            File.AppendAllText(rppath, Environment.NewLine + "### Kills" + Environment.NewLine + Environment.NewLine);
            foreach (RedditPostsKills kills in redditPosts.Kills)
            {
                kills.KillsList = kills.KillsList[..^2];
            }
            redditPosts.Kills = [.. redditPosts.Kills.OrderByDescending(x => x.KillsAmount).ThenBy(x => x.Player)];
            foreach (RedditPostsKills kills in redditPosts.Kills)
            {
                if (currentKill > 0)
                {
                    if (kills.KillsAmount == currentKill)
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
                currentKill = kills.KillsAmount;
            }
            File.AppendAllText(rppath, "---");

            //Pve Deaths
            File.AppendAllText(rppath, Environment.NewLine + "### PvE Deaths" + Environment.NewLine + Environment.NewLine);
            foreach (RedditPostsPve pveDeaths in redditPosts.PveDeaths)
            {
                pveDeaths.DeathsList = pveDeaths.DeathsList[..^2];
            }
            redditPosts.PveDeaths = [.. redditPosts.PveDeaths.OrderByDescending(x => x.DeathsAmount).ThenBy(x => x.PveCause)];
            foreach (RedditPostsPve pveDeaths in redditPosts.PveDeaths)
            {
                File.AppendAllText(rppath, "**" + pveDeaths.PveCause + " (" + pveDeaths.DeathsAmount + "):** " + pveDeaths.DeathsList + Environment.NewLine + Environment.NewLine);
            }
            File.AppendAllText(rppath, "---");

            //Participations
            File.AppendAllText(rppath, Environment.NewLine + "### Participation" + Environment.NewLine + Environment.NewLine);
            for (int seasons = totalSeasons; seasons > 0; seasons--)
            {
                String seasonPart = "";
                foreach (RedditPostsPlayed player in redditPosts.Participations)
                {
                    if (player.TotalPlayed == seasons)
                    {
                        seasonPart += player.Player + " " + player.SeasonsPlayed + ")" + ", ";
                    }
                }
                int count = seasonPart.Count(c => c == ' ') / 2;
                if (count > 0)
                {
                    seasonPart = seasonPart[..^2];
                }
                if (seasons == 1)
                {
                    File.AppendAllText(rppath, "**" + seasons.ToString() + " Season (" + count.ToString() + "):** " + seasonPart + Environment.NewLine + Environment.NewLine);
                }
                else
                {
                    File.AppendAllText(rppath, "**" + seasons.ToString() + " Seasons (" + count.ToString() + "):** " + seasonPart + Environment.NewLine + Environment.NewLine);
                }
            }
            File.AppendAllText(rppath, "---");

            //Debutants
            File.AppendAllText(rppath, Environment.NewLine + "### Debutants" + Environment.NewLine + Environment.NewLine);
            foreach (String debutants in redditPosts.Debutants)
            {
                File.AppendAllText(rppath, debutants + Environment.NewLine);
            }
            File.AppendAllText(rppath, "---");
        }
    }
}