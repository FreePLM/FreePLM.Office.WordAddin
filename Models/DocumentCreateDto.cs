namespace FreePLM.Office.WordAddin.Models
{
    /// <summary>
    /// Request to create a new document
    /// </summary>
    public class DocumentCreateDto
    {
        public string FileName { get; set; }
        public string Owner { get; set; }
        public string Group { get; set; }
        public string Role { get; set; }
        public string Project { get; set; }
        public string Comment { get; set; }
    }
}
