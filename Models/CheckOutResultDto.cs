using System;

namespace FreePLM.Office.WordAddin.Models
{
    /// <summary>
    /// Result of a checkout operation
    /// </summary>
    public class CheckOutResultDto
    {
        public bool Success { get; set; }
        public string ObjectId { get; set; }
        public string Revision { get; set; }
        public string DownloadUrl { get; set; }
        public DateTime CheckedOutDate { get; set; }
        public string Message { get; set; }
    }
}
