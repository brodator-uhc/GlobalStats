namespace StatsAnalyzer
{
    public class KillRecords(String player, int killRecord, String round, String season, DateTime date)
    {
        public String Player { get; set; } = player;
        public int KillRecord { get; set; } = killRecord;
        public String Round { get; set; } = round;
        public String Season { get; set; } = season;
        public DateTime Date { get; set; } = date;

        public static void UpdateKillRecord(List<KillRecords> killRecordsList, String player, int killRecord, String round, String season, DateTime date)
        {
            var playerKillRecord = killRecordsList.Find(p => p.Player == player);
            if (playerKillRecord != null)
            {
                if (killRecord > playerKillRecord.KillRecord)
                {
                    playerKillRecord.KillRecord = killRecord;
                    playerKillRecord.Round = round;
                    playerKillRecord.Season = season;
                    playerKillRecord.Date = date;
                }
                else if (killRecord == playerKillRecord.KillRecord)
                {
                    //If kill records are tied picks the first one that happened
                    if (date < playerKillRecord.Date)
                    {
                        playerKillRecord.Round = round;
                        playerKillRecord.Season = season;
                        playerKillRecord.Date = date;
                    }
                }
            }
        }
    }
}