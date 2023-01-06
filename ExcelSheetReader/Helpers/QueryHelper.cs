namespace ExcelSheetReader.Helpers
{
    public static class QueryHelper
    {
        public static Dictionary<string, string> Queries { get; private set; } = new()
        {
            { "Student", "INSERT IGNORE INTO Student (`StudentRecordId`, `StudentId`, `StateId`, `Birthdate`, `BirthCity`, `BirthStateCode`, `BirthCountryCode`, `EthnicityCode`, `Initial9thGradeYear`, `CohortYear`, `RecordSourceCode`)" },
            { "StudentRace", "INSERT IGNORE INTO StudentRace (`StudentRecordId`, `RaceCode`, `SortOrder`)" },
            { "StudentName", "INSERT IGNORE INTO StudentName (`StudentRecordId`, `Nametype`, `FirstName`, `LastName`, `MiddleName`, `SuffixCode`)" },
            { "StudentGender", "INSERT IGNORE INTO StudentGender (`StudentRecordId`, `GenderType`, `GenderCode`)" },
            { "StudentPronoun", "INSERT IGNORE INTO StudentPronoun (`StudentRecordId`, `PronounCode`)" },
            { "Address", "INSERT IGNORE INTO Address (`AddressRecordId`, `Address1`, `Address2`, `City`, `StateCode`, `ZipCode`, `ZipExtension`, `CountyCode`, `CountryCode`, `GridCode`, `CensusBlock`, `Latitude`, `Longitude`, `ValidationNotes`, `ValidationDate`)" },
            { "StudentAddress", "INSERT IGNORE INTO StudentAddress (`StudentRecordId`, `AddressRecordId`, `AddressType`, `StartDate`, `EndDate`)" },
            { "Phone", "INSERT IGNORE INTO Phone (`PhoneRecordId`, `CountryCode`, `AreaCode`, `PhoneNumber`, `Extension`)" },
            { "StudentPhone", "INSERT IGNORE INTO StudentPhone (`StudentRecordId`, `PhoneRecordId`, `PhoneType`)" },
            { "Email", "INSERT IGNORE INTO Email (`EmailRecordId`, `EmailAddress`)" },
            { "StudentEmail", "INSERT IGNORE INTO StudentEmail (`StudentRecordId`, `EmailRecordId`, `EmailType`)" },
            { "StudentSocial", "INSERT IGNORE INTO StudentSocial (`StudentRecordId`, `EncryptedData`)" },
            { "StudentUser", "INSERT IGNORE INTO StudentUser (`StudentRecordId`, `UserRecordId`)" }
        };
    }
}
