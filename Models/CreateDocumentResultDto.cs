namespace FreePLM.Office.WordAddin.Models
{
    /// <summary>
    /// Result from creating a new document
    /// </summary>
    public class CreateDocumentResultDto
    {
        public bool Success { get; set; }
        public string ObjectId { get; set; }
        public string FileName { get; set; }
        public string Revision { get; set; }
        public bool CheckedOut { get; set; }
        public string Message { get; set; }
    }
}
