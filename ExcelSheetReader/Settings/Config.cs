namespace ExcelSheetReader.Settings
{
    public static class Config
    {
        public static string FilePath { get; private set; }
        public static List<SheetInfo> WorkBookInfo { get; private set; }

        static Config()
        {
            FilePath = @"C:\Users\michaelc\Downloads\Seed Data - Student.xlsx";
            WorkBookInfo = new List<SheetInfo>()
            {
                new SheetInfo("Student", false, "TEMPStudent"),
                new SheetInfo("StudentRace", true),
                new SheetInfo("StudentName", true),
                new SheetInfo("StudentGender", true),
                new SheetInfo("StudentPronoun", true),
                new SheetInfo("Address", false, "TEMPAddress"),
                new SheetInfo("StudentAddress", true),
                new SheetInfo("Phone", false, "TEMPPhone"),
                new SheetInfo("StudentPhone", true),
                new SheetInfo("Email", false, "TEMPEmail"),
                new SheetInfo("StudentEmail", true),
                new SheetInfo("StudentSocial", true),
                new SheetInfo("StudentUser", true),
            };
        }
    }

    public class SheetInfo
    {
        public string SheetName { get; private set; }
        public bool IsJunctionTable { get; private set; }
        public string TempVarName { get; private set; }

        public SheetInfo(
            string sheetName, bool isJunctionTable, string tempVarName = "TEMPVariable"
        )
        {
            SheetName = sheetName;
            IsJunctionTable = isJunctionTable;
            TempVarName = tempVarName;
        }
    }
}
