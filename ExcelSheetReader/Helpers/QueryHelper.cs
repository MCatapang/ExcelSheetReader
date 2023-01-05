namespace ExcelSheetReader.Helpers
{
    public static class QueryHelper
    {
        public static Dictionary<string, string> Queries { get; private set; } = new()
        {
            { "Student", "INSERT IGNORE INTO Student ( `StudentId`, `StateId`, `Birthdate`, `BirthCity`, `BirthStateCode`, `BirthCountryCode`, `EthnicityCode`, `Initial9thGradeYear`, `CohortYear`, `RecordSourceCode`)" }
        };
    }
}
