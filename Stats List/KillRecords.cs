namespace StatsAnalyzer
{
    public class KillRecords(String player, int killRecord, SeasonInfo seasonInfo)
    {
        public String Player { get; set; } = player;
        public int KillRecord { get; set; } = killRecord;
        public String Round { get; set; } = seasonInfo.SeasonName;
        public String Season { get; set; } = seasonInfo.SeasonNumber;
        public DateTime Date { get; set; } = seasonInfo.SeasonDate;

        public static void UpdateKillRecord(List<KillRecords> killRecordsList, String player, int killRecord, SeasonInfo seasonInfo)
        {
            var playerKillRecord = killRecordsList.Find(p => p.Player == player);
            if (playerKillRecord != null)
            {
                if (killRecord > playerKillRecord.KillRecord)
                {
                    playerKillRecord.KillRecord = killRecord;
                    playerKillRecord.Round = seasonInfo.SeasonName;
                    playerKillRecord.Season = seasonInfo.SeasonNumber;
                    playerKillRecord.Date = seasonInfo.SeasonDate;
                }
                else if (killRecord == playerKillRecord.KillRecord)
                {
                    //If kill records are tied picks the first one that happened
                    if (seasonInfo.SeasonDate < playerKillRecord.Date)
                    {
                        playerKillRecord.Round = seasonInfo.SeasonName;
                        playerKillRecord.Season = seasonInfo.SeasonNumber;
                        playerKillRecord.Date = seasonInfo.SeasonDate;
                    }
                }
            }
        }
    }
}