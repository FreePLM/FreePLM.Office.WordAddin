namespace FreePLM.Office.WordAddin.Models
{
    /// <summary>
    /// Request to check in a document
    /// </summary>
    public class CheckInDto
    {
        public string ObjectId { get; set; }
        public string Comment { get; set; }
        public bool CreateMajorRevision { get; set; }
        public DocumentStatus? NewStatus { get; set; }
    }
}
