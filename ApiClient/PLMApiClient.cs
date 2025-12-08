using System;
using System.IO;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;
using FreePLM.Office.WordAddin.Models;
using Newtonsoft.Json;

namespace FreePLM.Office.WordAddin.ApiClient
{
    /// <summary>
    /// API client for communicating with the FreePLM service
    /// </summary>
    public class PLMApiClient : IDisposable
    {
        private readonly HttpClient _httpClient;
        private readonly string _baseUrl;

        public PLMApiClient(string baseUrl = "http://localhost:5000")
        {
            _baseUrl = baseUrl;
            _httpClient = new HttpClient
            {
                BaseAddress = new Uri(_baseUrl),
                Timeout = TimeSpan.FromMinutes(5)
            };
            _httpClient.DefaultRequestHeaders.Accept.Add(
                new MediaTypeWithQualityHeaderValue("application/json"));
        }

        #region Document Operations

        /// <summary>
        /// Get document information by ObjectId
        /// </summary>
        public async Task<DocumentReadDto> GetDocumentAsync(string objectId)
        {
            var response = await _httpClient.GetAsync($"/api/documents/{objectId}");
            response.EnsureSuccessStatusCode();

            var json = await response.Content.ReadAsStringAsync();
            return JsonConvert.DeserializeObject<DocumentReadDto>(json);
        }

        /// <summary>
        /// Download document file content
        /// </summary>
        public async Task<byte[]> DownloadFileAsync(string objectId, string revision = null)
        {
            var url = $"/api/documents/{objectId}/content";
            if (!string.IsNullOrEmpty(revision))
            {
                url += $"?revision={revision}";
            }

            var response = await _httpClient.GetAsync(url);
            response.EnsureSuccessStatusCode();

            return await response.Content.ReadAsByteArrayAsync();
        }

        #endregion

        #region CheckOut/CheckIn Operations

        /// <summary>
        /// Check out a document
        /// </summary>
        public async Task<CheckOutResultDto> CheckOutAsync(string objectId, string comment = null)
        {
            var request = new CheckOutDto
            {
                ObjectId = objectId,
                Comment = comment,
                MachineName = Environment.MachineName
            };

            var json = JsonConvert.SerializeObject(request);
            var content = new StringContent(json, Encoding.UTF8, "application/json");

            var response = await _httpClient.PostAsync("/api/checkout", content);

            var responseJson = await response.Content.ReadAsStringAsync();

            if (!response.IsSuccessStatusCode)
            {
                var error = JsonConvert.DeserializeObject<ApiErrorResponse>(responseJson);
                throw new PLMApiException(error.Message, response.StatusCode, error.Error);
            }

            return JsonConvert.DeserializeObject<CheckOutResultDto>(responseJson);
        }

        /// <summary>
        /// Check in a document
        /// </summary>
        public async Task<CheckInResultDto> CheckInAsync(
            string objectId,
            byte[] fileContent,
            string comment,
            bool createMajorRevision = false,
            DocumentStatus? newStatus = null)
        {
            using (var multipart = new MultipartFormDataContent())
            {
                // Add form fields
                multipart.Add(new StringContent(objectId), "objectId");
                multipart.Add(new StringContent(comment), "comment");
                multipart.Add(new StringContent(createMajorRevision.ToString()), "createMajorRevision");

                if (newStatus.HasValue)
                {
                    multipart.Add(new StringContent(newStatus.Value.ToString()), "newStatus");
                }

                // Add file content
                var fileContentPart = new ByteArrayContent(fileContent);
                fileContentPart.Headers.ContentType = MediaTypeHeaderValue.Parse("application/octet-stream");
                multipart.Add(fileContentPart, "file", "document.docx");

                var response = await _httpClient.PostAsync("/api/checkin", multipart);

                var responseJson = await response.Content.ReadAsStringAsync();

                if (!response.IsSuccessStatusCode)
                {
                    var error = JsonConvert.DeserializeObject<ApiErrorResponse>(responseJson);
                    throw new PLMApiException(error.Message, response.StatusCode, error.Error);
                }

                return JsonConvert.DeserializeObject<CheckInResultDto>(responseJson);
            }
        }

        /// <summary>
        /// Cancel checkout (undo checkout without creating revision)
        /// </summary>
        public async Task CancelCheckOutAsync(string objectId)
        {
            var request = new { objectId = objectId };
            var json = JsonConvert.SerializeObject(request);
            var content = new StringContent(json, Encoding.UTF8, "application/json");

            var response = await _httpClient.PostAsync("/api/checkout/cancel", content);

            if (!response.IsSuccessStatusCode)
            {
                var responseJson = await response.Content.ReadAsStringAsync();
                var error = JsonConvert.DeserializeObject<ApiErrorResponse>(responseJson);
                throw new PLMApiException(error.Message, response.StatusCode, error.Error);
            }
        }

        /// <summary>
        /// Get checkout/lock status
        /// </summary>
        public async Task<bool> IsCheckedOutAsync(string objectId)
        {
            var doc = await GetDocumentAsync(objectId);
            return doc.IsCheckedOut;
        }

        #endregion

        #region Workflow Operations

        /// <summary>
        /// Change document status
        /// </summary>
        public async Task ChangeStatusAsync(string objectId, DocumentStatus newStatus, string comment)
        {
            var request = new
            {
                objectId = objectId,
                newStatus = newStatus,
                comment = comment
            };

            var json = JsonConvert.SerializeObject(request);
            var content = new StringContent(json, Encoding.UTF8, "application/json");

            var response = await _httpClient.PostAsync("/api/workflow/status", content);

            if (!response.IsSuccessStatusCode)
            {
                var responseJson = await response.Content.ReadAsStringAsync();
                var error = JsonConvert.DeserializeObject<ApiErrorResponse>(responseJson);
                throw new PLMApiException(error.Message, response.StatusCode, error.Error);
            }
        }

        #endregion

        #region Health Check

        /// <summary>
        /// Check if the FreePLM service is running
        /// </summary>
        public async Task<bool> IsServiceAvailableAsync()
        {
            try
            {
                var response = await _httpClient.GetAsync("/api/health");
                return response.IsSuccessStatusCode;
            }
            catch
            {
                return false;
            }
        }

        #endregion

        public void Dispose()
        {
            _httpClient?.Dispose();
        }
    }
}
