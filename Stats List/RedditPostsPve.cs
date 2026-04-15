namespace StatsAnalyzer
{
    public class RedditPostsPve(String pveCause, int deathsAmount, String deathsList)
    {
        public String PveCause { get; set; } = pveCause;
        public int DeathsAmount { get; set; } = deathsAmount;
        public String DeathsList { get; set; } = deathsList;

        public static void UpdatePve(RedditPosts redditPosts, String pveCause, String victim, String season)
        {
            var redditPostsPve = redditPosts.PveDeaths.Find(p => p.PveCause == pveCause);
            if (redditPostsPve != null)
            {
                redditPostsPve.DeathsAmount += 1;
                redditPostsPve.DeathsList = redditPostsPve.DeathsList + victim + " (S" + season + "), ";
            }
            else
            {
                String deathsList = victim + " (S" + season + "), ";
                redditPosts.PveDeaths.Add(new RedditPostsPve(pveCause, 1, deathsList));
            }
        }
    }
}