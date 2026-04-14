namespace GlobalStats
{
    public class RostersList
    {
        public String Round { get; set; }
        public String Season { get; set; }
        public DateTime Date { get; set; }
        public List<String> Roster { get; set; } = new List<String>();

        public RostersList(String round, String season, DateTime date)
        {
            Round = round;
            Season = season;
            Date = date;
        }
    }
}