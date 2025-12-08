namespace FreePLM.Office.WordAddin.Models
{
    /// <summary>
    /// Request to check out a document
    /// </summary>
    public class CheckOutDto
    {
        public string ObjectId { get; set; }
        public string Comment { get; set; }
        public string MachineName { get; set; }
    }
}
