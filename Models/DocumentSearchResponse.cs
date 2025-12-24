using System;
using System.Collections.Generic;

namespace FreePLM.Office.WordAddin.Models
{
    public class DocumentSearchResponse
    {
        public List<DocumentSearchResultDto> Documents { get; set; }
        public int TotalCount { get; set; }
        public int PageNumber { get; set; }
        public int PageSize { get; set; }
        public int TotalPages { get; set; }
    }

    public class DocumentSearchResultDto
    {
        public string ObjectId { get; set; }
        public string FileName { get; set; }
        public string CurrentRevision { get; set; }
        public DocumentStatus Status { get; set; }
        public string Owner { get; set; }
        public string Project { get; set; }
        public string Group { get; set; }
        public string Role { get; set; }
        public DateTime CreatedDate { get; set; }
        public DateTime ModifiedDate { get; set; }
        public bool IsCheckedOut { get; set; }
        public string CheckedOutBy { get; set; }
        public DateTime? CheckedOutDate { get; set; }
    }
}
