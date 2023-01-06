namespace ExcelSheetReader.Helpers
{
    public static class CodeHelper
    {
        public static Dictionary<string, int> Codes { get; private set; } = new()
        {
            { "CA", 49 },
            { "IL", 71 },
            { "SO", 124 },
            { "IN", 244 },
            { "MX", 284 },
            { "US", 376 },
            { "N", 33 },
            { "Y", 34 },
            { "Z", 35 },
            { "Aeries Web", 0 },
            { "Manual Entry", 1 },
            { "Imported", 2 },
            { "JR", 1 },
            { "PREF", 0 },
            { "LEGAL", 1 },
            { "F", 568 },
            { "M", 569 },
            { "He/Him/His", 598 },
            { "She/Her/Hers", 599 },
            { "They/Them/Theirs", 600 },
            { "P", 0 },
            { "D", 0 }
        };
    }
}
