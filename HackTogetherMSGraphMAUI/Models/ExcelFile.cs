
namespace HackTogetherMSGraphMAUI.Models
{
    /// <summary>
    /// Model for Excel file
    /// </summary>
    public class ExcelFile
    {
        public string Id { get; set; }

        /// <summary>
        /// File full name
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// File last modified date and time
        /// </summary>
        public string LastModifiedDateTime { get; set; }

        /// <summary>
        /// File created date and time
        /// </summary>
        public string CreatedDateTime { get; set; }

        /// <summary>
        /// Number of sheets in file
        /// </summary>
        public int Sheets { get; set; }

        /// <summary>
        /// Number of tables in file
        /// </summary>
        public int Tables { get; set; }

    }
}
