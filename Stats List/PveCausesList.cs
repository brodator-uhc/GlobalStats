using ClosedXML.Excel;

namespace StatsAnalyzer
{
    public class PveCausesList(String pveCause, int deathsCaused)
    {
        public String PveCause { get; set; } = pveCause;
        public int DeathsCaused { get; set; } = deathsCaused;

        public static void UpdatePveCauses(List<PveCausesList> pveCausesList, String pveCause)
        {
            var pveCauseStats = pveCausesList.Find(p => p.PveCause == pveCause);
            pveCauseStats?.DeathsCaused += 1;
        }

        public static String GetPveCause(String method)
        {
            String pvedeath = "";

            switch (method)
            {
                case String m when m.Contains("lava") && !m.Contains("discovered"):
                pvedeath = "Lava";
                break;
                case String m when m.Contains("discovered"):
                pvedeath = "Magma";
                break;
                case String m when m.Contains("ground") || m.Contains("doomed") || m.Contains("fell") && !m.Contains("world"):
                pvedeath = "Fall";
                break;
                case String m when m.Contains("world"):
                pvedeath = "Void";
                break;
                case String m when m.Contains("drowned"):
                pvedeath = "Drowning";
                break;
                case String m when m.Contains("suffocated"):
                pvedeath = "Suffocation";
                break;
                case String m when m.Contains("burnt") || m.Contains("burned"):
                pvedeath = "Burning";
                break;
                case String m when m.Contains("starved"):
                pvedeath = "Starvation";
                break;
                case String m when m.Contains("fallout"):
                pvedeath = "Fallout";
                break;
                case String m when m.Contains("swords"):
                pvedeath = "Diamond Sword";
                break;
                case String m when m.Contains("water"):
                pvedeath = "Water";
                break;
                case String m when m.Contains("disqualified"):
                pvedeath = "Disqualified";
                break;
                case String m when m.Contains("bats"):
                pvedeath = "Bats";
                break;
                case String m when m.Contains("extra"):
                pvedeath = "Extra Damage";
                break;
                case String m when m.Contains("diamonds"):
                pvedeath = "Blood Diamonds";
                break;
                case String m when m.Contains("gambled"):
                pvedeath = "Gambling";
                break;
                case String m when m.Contains("button"):
                pvedeath = "Push The Button";
                break;
                case String m when m.Contains("hell"):
                pvedeath = "Go To Hell";
                break;
                case String m when m.Contains("comply"):
                pvedeath = "Comply";
                break;
                case String m when m.Contains("learned"):
                pvedeath = "Newtons Third Law";
                break;
                case String m when m.Contains("infiltrator"):
                pvedeath = "Infiltrator";
                break;
                case String m when m.Contains("love"):
                pvedeath = "Love";
                break;
                case String m when m.Contains("Design"):
                pvedeath = "Bed";
                break;
                case String m when m.Contains("blew"):
                pvedeath = "Explosion";
                break;
                case String m when m.Contains("sneaked"):
                pvedeath = "Sneaking";
                break;
                case String m when m.Contains("withered"):
                pvedeath = "Withered";
                break;
                case String m when m.Contains("timed") || m.Contains("disconnected") || m.Contains("offline"):
                pvedeath = "Left";
                break;
                case String m when m.Contains("stalagmite") || m.Contains("stalactite"):
                pvedeath = "Dripstone";
                break;
                case String m when m.Contains("anvil"):
                pvedeath = "Anvil";
                break;
                case String m when m.Contains("pricked"):
                pvedeath = "Cactus";
                break;
                case String m when m.Contains("poked"):
                pvedeath = "Sweet Berry Bush";
                break;
                case String m when m.Contains("kinetic"):
                pvedeath = "Elytra";
                break;
                case String m when m.Contains("bang"):
                pvedeath = "Firework";
                break;
                case String m when m.Contains("died"):
                pvedeath = "Death";
                break;
                case String m when m.Contains("flames"):
                pvedeath = "Fire";
                break;
                case String m when m.Contains("pummeled"):
                pvedeath = "Pummeled";
                break;
                case String m when m.Contains("magic"):
                pvedeath = "Potion";
                break;
                case String m when m.Contains("lightning"):
                pvedeath = "Lightning";
                break;
                case String m when m.Contains("shot"):
                pvedeath = "Arrow";
                break;
                default:
                pvedeath = "N/A";
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("ERROR: " + method + " does not have a PvE Category!");
                break;
            }

            return pvedeath;
        }
    }
}