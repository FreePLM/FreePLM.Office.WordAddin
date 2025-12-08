using System;

namespace FreePLM.Office.WordAddin.Models
{
    /// <summary>
    /// Document information (without file content)
    /// </summary>
    public class DocumentReadDto
    {
        public string ObjectId { get; set; }
        public string FileName { get; set; }
        public string CurrentRevision { get; set; }
        public DocumentStatus Status { get; set; }
        public string Owner { get; set; }
        public string Group { get; set; }
        public string Role { get; set; }
        public string Project { get; set; }
        public DateTime CreatedDate { get; set; }
        public string CreatedBy { get; set; }
        public DateTime ModifiedDate { get; set; }
        public string ModifiedBy { get; set; }
        public long FileSize { get; set; }
        public bool IsCheckedOut { get; set; }
        public string CheckedOutBy { get; set; }
        public DateTime? CheckedOutDate { get; set; }
    }
}
