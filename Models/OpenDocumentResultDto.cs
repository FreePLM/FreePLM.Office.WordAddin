namespace FreePLM.Office.WordAddin.Models
{
    public class OpenDocumentResultDto
    {
        public string ObjectId { get; set; }
        public string FileName { get; set; }
        public string Revision { get; set; }
        public byte[] FileContent { get; set; }
        public bool IsCheckedOut { get; set; }
        public string CheckedOutBy { get; set; }
        public DocumentStatus Status { get; set; }
    }
}
