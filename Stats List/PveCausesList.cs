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
            String pveDeath = "";

            switch (method)
            {
                case String m when m.Contains("lava") && !m.Contains("discovered"):
                pveDeath = "Lava";
                break;
                case String m when m.Contains("discovered"):
                pveDeath = "Magma";
                break;
                case String m when m.Contains("ground") || m.Contains("doomed") || m.Contains("fell") && !m.Contains("world"):
                pveDeath = "Fall";
                break;
                case String m when m.Contains("world"):
                pveDeath = "Void";
                break;
                case String m when m.Contains("drowned"):
                pveDeath = "Drowning";
                break;
                case String m when m.Contains("suffocated"):
                pveDeath = "Suffocation";
                break;
                case String m when m.Contains("burnt") || m.Contains("burned"):
                pveDeath = "Burning";
                break;
                case String m when m.Contains("starved"):
                pveDeath = "Starvation";
                break;
                case String m when m.Contains("fallout"):
                pveDeath = "Fallout";
                break;
                case String m when m.Contains("swords"):
                pveDeath = "Diamond Sword";
                break;
                case String m when m.Contains("water"):
                pveDeath = "Water";
                break;
                case String m when m.Contains("disqualified"):
                pveDeath = "Disqualified";
                break;
                case String m when m.Contains("bats"):
                pveDeath = "Bats";
                break;
                case String m when m.Contains("extra"):
                pveDeath = "Extra Damage";
                break;
                case String m when m.Contains("diamonds"):
                pveDeath = "Blood Diamonds";
                break;
                case String m when m.Contains("gambled"):
                pveDeath = "Gambling";
                break;
                case String m when m.Contains("button"):
                pveDeath = "Push The Button";
                break;
                case String m when m.Contains("hell"):
                pveDeath = "Go To Hell";
                break;
                case String m when m.Contains("comply"):
                pveDeath = "Comply";
                break;
                case String m when m.Contains("learned"):
                pveDeath = "Newtons Third Law";
                break;
                case String m when m.Contains("infiltrator"):
                pveDeath = "Infiltrator";
                break;
                case String m when m.Contains("love"):
                pveDeath = "Love";
                break;
                case String m when m.Contains("Design"):
                pveDeath = "Bed";
                break;
                case String m when m.Contains("blew"):
                pveDeath = "Explosion";
                break;
                case String m when m.Contains("sneaked"):
                pveDeath = "Sneaking";
                break;
                case String m when m.Contains("withered"):
                pveDeath = "Withered";
                break;
                case String m when m.Contains("timed") || m.Contains("disconnected") || m.Contains("offline"):
                pveDeath = "Left";
                break;
                case String m when m.Contains("stalagmite") || m.Contains("stalactite"):
                pveDeath = "Dripstone";
                break;
                case String m when m.Contains("anvil"):
                pveDeath = "Anvil";
                break;
                case String m when m.Contains("pricked"):
                pveDeath = "Cactus";
                break;
                case String m when m.Contains("poked"):
                pveDeath = "Sweet Berry Bush";
                break;
                case String m when m.Contains("kinetic"):
                pveDeath = "Elytra";
                break;
                case String m when m.Contains("bang"):
                pveDeath = "Firework";
                break;
                case String m when m.Contains("died"):
                pveDeath = "Death";
                break;
                case String m when m.Contains("flames"):
                pveDeath = "Fire";
                break;
                case String m when m.Contains("pummeled"):
                pveDeath = "Pummeled";
                break;
                case String m when m.Contains("magic"):
                pveDeath = "Potion";
                break;
                case String m when m.Contains("hot potato"):
                pveDeath = "Hot Potato";
                break;
                case String m when m.Contains("lightning"):
                pveDeath = "Lightning";
                break;
                case String m when m.Contains("shot"):
                pveDeath = "Arrow";
                break;
                default:
                pveDeath = "N/A";
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("ERROR: " + method + " does not have a PvE Category!");
                break;
            }
            return pveDeath;
        }
    }
}