namespace StatsAnalyzer
{
    public class PlayerStats(String round, String season, DateTime date)
    {
        public String Round { get; set; } = round;
        public String Season { get; set; } = season;
        public DateTime Date { get; set; } = date;
        public String TeamColor { get; set; } = "w";
        public String Team { get; set; } = "/";
        public int KillsTotal { get; set; } = 0;
        public String Kills { get; set; } = "/";
        public String PveDeath { get; set; } = "todelete";
        public String Death { get; set; } = "todelete";
        public String FirstDamage { get; set; } = "todelete";
        public String Ironman { get; set; } = "todelete";
        public String FirstDeath { get; set; } = "todelete";
        public String FirstBlood { get; set; } = "todelete";
        public String TopFrag { get; set; } = "todelete";
        public String RunnerUp { get; set; } = "todelete";
        public String Win { get; set; } = "todelete";

        public static String GetTeamColorChar(String teamColor)
        {
            String teamColorChar = "";
            teamColorChar = teamColor switch
            {
                "Black" => "b",
                "Blue" => "f",
                "Cyan" => "c",
                "Dark Blue" => "d",
                "Dark Gray" => "e",
                "Dark Green" => "g",
                "Dark Red" => "r",
                "Light Blue" => "a",
                "Light Gray" => "s",
                "Light Green" => "l",
                "Orange" => "o",
                "Pink" => "m",
                "Purple" => "p",
                "Red" => "t",
                "Yellow" => "y",
                _ => "w",
            };
            return teamColorChar;
        }
    }
}