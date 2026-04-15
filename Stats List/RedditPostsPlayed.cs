namespace StatsAnalyzer
{
    public class RedditPostsPlayed(String player, int totalPlayed, String lastPlayed, String seasonsPlayed)
    {
        public String Player { get; set; } = player;
        public int TotalPlayed { get; set; } = totalPlayed;
        public String LastPlayed { get; set; } = lastPlayed;
        public String SeasonsPlayed { get; set; } = seasonsPlayed;

        public static void UpdateParticipations(RedditPosts redditPosts, String player, String lastPlayed, String lastSeason)
        {
            var redditPostsParticipations = redditPosts.Participations.Find(p => p.Player == player);
            if (redditPostsParticipations != null)
            {
                redditPostsParticipations.TotalPlayed += 1;
            }
            else
            {
                redditPosts.Participations.Add(new RedditPostsPlayed(player, 1, lastPlayed, ""));
            }

            redditPostsParticipations = redditPosts.Participations.Find(p => p.Player == player);
            if (redditPostsParticipations!.LastPlayed.Equals(lastPlayed))
            {
                redditPostsParticipations.SeasonsPlayed = "(S" + lastPlayed;
            }
            else if (redditPostsParticipations.LastPlayed.Equals(lastSeason))
            {
                char char_season = redditPostsParticipations.SeasonsPlayed[^(lastSeason.Length + 2)];
                if (char_season.Equals('-'))
                {
                    redditPostsParticipations.SeasonsPlayed = redditPostsParticipations.SeasonsPlayed[..^(lastSeason.Length + 1)];
                    redditPostsParticipations.SeasonsPlayed += "S" + lastPlayed;
                }
                else
                {
                    redditPostsParticipations.SeasonsPlayed += "-S" + lastPlayed;
                }
                redditPostsParticipations.LastPlayed = lastPlayed;
            }
            else
            {
                redditPostsParticipations.SeasonsPlayed += ",S" + lastPlayed;
                redditPostsParticipations.LastPlayed = lastPlayed;
            }
        }
    }
}
