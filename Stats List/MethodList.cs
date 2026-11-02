namespace StatsAnalyzer
{
    public class MethodList(String method, int timesUsed)
    {
        public String Method { get; set; } = method;
        public int TimesUsed { get; set; } = timesUsed;

        public static void UpdateMethod(List<MethodList> methodLists, String method)
        {
            var methodStats = methodLists.Find(p => p.Method == method);
            methodStats?.TimesUsed += 1;
        }
    }
}