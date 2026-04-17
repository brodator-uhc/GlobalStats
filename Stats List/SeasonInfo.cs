using ClosedXML.Excel;

namespace StatsAnalyzer
{
    public class SeasonInfo
    {
        public String SeasonName { get; set; }
        public String SeasonNumber { get; set; }
        public DateTime SeasonDate { get; set; }
        public String SeasonGamemodes { get; set; }
        public String SeasonTeamType { get; set; }
        public bool IsFFA { get; set; }
        public bool IsCrossoverSeason { get; set; }

        public SeasonInfo(IXLWorksheet roundPage, int firstDataColumn)
        {
            SeasonName = roundPage.Name;
            SeasonNumber = roundPage.Cell(1, firstDataColumn).GetString();
            SeasonDate = roundPage.Cell(2, firstDataColumn).GetDateTime();
            SeasonGamemodes = roundPage.Cell(4, firstDataColumn).GetString();
            SeasonTeamType = roundPage.Cell(3, firstDataColumn + 1).GetString();
            IsFFA = roundPage.Cell(3, firstDataColumn + 1).GetString().Equals("FFA");

            //Sets round named to be changed for crossovers and ??? to not be called Sheet
            if (SeasonName.Contains("Sheet"))
            {
                SeasonName = "???";
            }

            //Sets name of round to contain both names if crossovers
            if (SeasonName.Equals("Phobia") && SeasonNumber.Equals("20"))
            {
                SeasonName = "WMC x Phobia";
                SeasonNumber = "30/20";
            }
            if (SeasonName.Equals("Scattershot") && SeasonNumber.Equals("6"))
            {
                SeasonName = "The Melon Blooded x Scattershot";
                SeasonNumber = "40/6";
            }
            if (SeasonName.Equals("Cinema") && SeasonNumber.Equals("16b"))
            {
                SeasonName = "Phobia x Cinema";
                SeasonNumber = "28/16b";
            }

            //Only count one of the crossover round towards itself only
            if (SeasonName.Equals("WMC") && SeasonNumber.Equals("30") ||
                SeasonName.Equals("The Melon Blooded") && SeasonNumber.Equals("40") ||
                SeasonName.Equals("Phobia") && SeasonNumber.Equals("28"))
            {
                IsCrossoverSeason = true;
            }
        }
    }
}