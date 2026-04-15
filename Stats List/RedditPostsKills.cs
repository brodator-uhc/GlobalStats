namespace StatsAnalyzer
{
    public class RedditPostsKills(String player, int killsAmount, String killsList)
    {
        public String Player { get; set; } = player;
        public int KillsAmount { get; set; } = killsAmount;
        public String KillsList { get; set; } = killsList;

        public static void UpdateKills(RedditPosts redditPosts, String player, String victim, String season)
        {
            var redditPostsKills = redditPosts.Kills.Find(p => p.Player == player);
            if (redditPostsKills != null)
            {
                redditPostsKills.KillsAmount += 1;
                redditPostsKills.KillsList = redditPostsKills.KillsList + victim + " (S" + season + "), ";
            }
            else
            {
                String killsList = victim + " (S" + season + "), ";
                redditPosts.Kills.Add(new RedditPostsKills(player, 1, killsList));
            }
        }
    }
}