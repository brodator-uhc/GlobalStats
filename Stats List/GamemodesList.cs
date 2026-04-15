namespace StatsAnalyzer
{
    public class GamemodesList(String gamemode, int timesUsed)
    {
        public String Gamemode { get; set; } = gamemode;
        public int TimesUsed { get; set; } = timesUsed;

        public static void UpdateGamemodes(List<GamemodesList> gamemodesLists, String gamemode)
        {
            var gamemodesStats = gamemodesLists.Find(p => p.Gamemode == gamemode);
            gamemodesStats?.TimesUsed += 1;
        }
    }
}