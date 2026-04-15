namespace StatsAnalyzer
{
    public class RedditPosts
    {
        public List<String> Winners { get; set; } = [];
        public List<String> RunnerUps { get; set; } = [];
        public List<String> MostKills { get; set; } = [];
        public List<String> MostKillsTeam { get; set; } = [];
        public List<String> FirstDamage { get; set; } = [];
        public List<String> Ironman { get; set; } = [];
        public List<String> FirstBlood { get; set; } = [];
        public List<String> FirstDeath { get; set; } = [];
        public List<RedditPostsKills> Kills { get; set; } = [];
        public List<RedditPostsPve> PveDeaths { get; set; } = [];
        public List<RedditPostsPlayed> Participations { get; set; } = [];
        public List<String> Debutants { get; set; } = [];
    }
}