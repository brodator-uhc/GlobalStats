namespace GlobalStats
{
    public class PlayerStats
    {
        public String Round { get; set; }
        public String Season { get; set; }
        public DateTime Date { get; set; }
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

        public PlayerStats(String round, String season, DateTime date)
        {
            Round = round;
            Season = season;
            Date = date;
        }

        public static String GetTeamColorChar(String teamColor)
        {
            String teamColorChar = "";
            switch (teamColor)
            {
                case "Black":
                    teamColorChar = "b";
                    break;
                case "Blue":
                    teamColorChar = "f";
                    break;
                case "Cyan":
                    teamColorChar = "c";
                    break;
                case "Dark Blue":
                    teamColorChar = "d";
                    break;
                case "Dark Gray":
                    teamColorChar = "e";
                    break;
                case "Dark Green":
                    teamColorChar = "g";
                    break;
                case "Dark Red":
                    teamColorChar = "r";
                    break;
                case "Light Blue":
                    teamColorChar = "a";
                    break;
                case "Light Gray":
                    teamColorChar = "s";
                    break;
                case "Light Green":
                    teamColorChar = "l";
                    break;
                case "Orange":
                    teamColorChar = "o";
                    break;
                case "Pink":
                    teamColorChar = "m";
                    break;
                case "Purple":
                    teamColorChar = "p";
                    break;
                case "Red":
                    teamColorChar = "t";
                    break;
                case "Yellow":
                    teamColorChar = "y";
                    break;
                default:
                    teamColorChar = "w";
                    break;
            }
            return teamColorChar;
        }
    }
}