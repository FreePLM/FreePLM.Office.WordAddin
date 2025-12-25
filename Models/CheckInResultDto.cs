using System;

namespace FreePLM.Office.WordAddin.Models
{
    /// <summary>
    /// Result of a checkin operation
    /// </summary>
    public class CheckInResultDto
    {
        public bool Success { get; set; }
        public string ObjectId { get; set; }
        public string NewRevision { get; set; }
        public string PreviousRevision { get; set; }
        public DateTime CheckedInDate { get; set; }
        public string Message { get; set; }
        public bool CloseAfterCheckIn { get; set; }
    }
}
