using System.Collections.Generic;

namespace FreePLM.Office.WordAddin.Models
{
    /// <summary>
    /// API error response
    /// </summary>
    public class ApiErrorResponse
    {
        public bool Success { get; set; }
        public string Error { get; set; }
        public string Message { get; set; }
        public Dictionary<string, object> Details { get; set; }
    }
}
