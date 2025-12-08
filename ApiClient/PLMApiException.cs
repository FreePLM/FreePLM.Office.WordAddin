using System;
using System.Net;

namespace FreePLM.Office.WordAddin.ApiClient
{
    /// <summary>
    /// Exception thrown when PLM API call fails
    /// </summary>
    public class PLMApiException : Exception
    {
        public HttpStatusCode StatusCode { get; }
        public string ErrorCode { get; }

        public PLMApiException(string message, HttpStatusCode statusCode, string errorCode = null)
            : base(message)
        {
            StatusCode = statusCode;
            ErrorCode = errorCode;
        }

        public bool IsDocumentLocked => ErrorCode == "DocumentLocked";
        public bool IsLockNotHeld => ErrorCode == "LockNotHeld";
        public bool IsNotFound => StatusCode == HttpStatusCode.NotFound;
    }
}
