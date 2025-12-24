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

        /// <summary>
        /// Show check-in UI dialog and process check-in
        /// </summary>
        public async Task<CheckInResultDto> CheckInUIAsync(string objectId, string fileName, string currentRevision)
        {
            var requestData = new
            {
                objectId = objectId,
                fileName = fileName,
                currentRevision = currentRevision
            };

            var json = JsonConvert.SerializeObject(requestData);
            var content = new StringContent(json, Encoding.UTF8, "application/json");

            var response = await _httpClient.PostAsync("/api/checkout/checkin-ui", content);

            if (!response.IsSuccessStatusCode)
            {
                var errorContent = await response.Content.ReadAsStringAsync();
                throw new PLMApiException($"Check-in UI failed: {errorContent}", response.StatusCode);
            }

            var jsonResponse = await response.Content.ReadAsStringAsync();
            var result = JsonConvert.DeserializeObject<dynamic>(jsonResponse);

            // Check if user cancelled
            if (result.success == false)
            {
                return null; // User cancelled
            }

            return new CheckInResultDto
            {
                Success = result.success,
                ObjectId = result.objectId,
                NewRevision = result.newRevision,
                PreviousRevision = result.previousRevision,
                CheckedInDate = result.checkedInDate,
                Message = result.message
            };
        }

        #region Document Operations

        /// <summary>
        /// Show create document UI dialog and process creation
        /// </summary>
        public async Task<CreateDocumentResultDto> CreateDocumentUIAsync()
        {
            var response = await _httpClient.PostAsync("/api/document/create-ui", null);

            var responseJson = await response.Content.ReadAsStringAsync();

            if (!response.IsSuccessStatusCode)
            {
                var error = JsonConvert.DeserializeObject<ApiErrorResponse>(responseJson);
                throw new PLMApiException(error?.Message ?? "Create document failed", response.StatusCode, error?.Error);
            }

            var result = JsonConvert.DeserializeObject<dynamic>(responseJson);

            // Check if user cancelled
            if (result.success == false)
            {
                return null; // User cancelled
            }

            return new CreateDocumentResultDto
            {
                Success = result.success,
                ObjectId = result.objectId,
                FileName = result.fileName,
                Revision = result.revision,
                CheckedOut = result.checkedOut,
                Message = result.message
            };
        }

        /// <summary>
        /// Show save as UI dialog to bring existing document into PLM
        /// </summary>
        public async Task<CreateDocumentResultDto> SaveAsUIAsync(string filePath, string currentFileName)
        {
            using (var multipart = new MultipartFormDataContent())
            {
                // Add current filename
                if (!string.IsNullOrEmpty(currentFileName))
                {
                    multipart.Add(new StringContent(currentFileName), "currentFileName");
                }

                // Add file content - use FileShare.ReadWrite to handle Word's lingering locks
                byte[] fileContent = null;
                int maxRetries = 5;
                int retryDelayMs = 500;

                for (int i = 0; i < maxRetries; i++)
                {
                    try
                    {
                        using (var fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                        {
                            fileContent = new byte[fileStream.Length];
                            fileStream.Read(fileContent, 0, (int)fileStream.Length);
                        }
                        break; // Success - exit retry loop
                    }
                    catch (IOException ex)
                    {
                        if (i == maxRetries - 1)
                        {
                            throw new IOException($"Unable to read file after {maxRetries} attempts: {filePath}", ex);
                        }
                        // File still locked, wait and retry
                        System.Threading.Thread.Sleep(retryDelayMs);
                    }
                }

                var fileContentPart = new ByteArrayContent(fileContent);
                fileContentPart.Headers.ContentType = MediaTypeHeaderValue.Parse("application/octet-stream");
                multipart.Add(fileContentPart, "file", currentFileName ?? "document.docx");

                var response = await _httpClient.PostAsync("/api/document/saveas-ui", multipart);

                var responseJson = await response.Content.ReadAsStringAsync();

                if (!response.IsSuccessStatusCode)
                {
                    var error = JsonConvert.DeserializeObject<ApiErrorResponse>(responseJson);
                    throw new PLMApiException(error?.Message ?? "Save As failed", response.StatusCode, error?.Error);
                }

                var result = JsonConvert.DeserializeObject<dynamic>(responseJson);

                // Check if user cancelled
                if (result.success == false)
                {
                    return null; // User cancelled
                }

                return new CreateDocumentResultDto
                {
                    Success = result.success,
                    ObjectId = result.objectId,
                    FileName = result.fileName,
                    Revision = result.revision,
                    CheckedOut = result.checkedOut,
                    Message = result.message
                };
            }
        }

        /// <summary>
        /// Create a new document
        /// </summary>
        public async Task<CreateDocumentResultDto> CreateDocumentAsync(string fileName, string comment = null)
        {
            var request = new
            {
                fileName = fileName,
                owner = "user@example.com",
                group = "Default",
                role = "Default",
                project = "Default",
                machineName = Environment.MachineName,
                comment = comment ?? "Initial document creation"
            };

            var json = JsonConvert.SerializeObject(request);
            var content = new StringContent(json, Encoding.UTF8, "application/json");

            var response = await _httpClient.PostAsync("/api/document/create", content);

            var responseJson = await response.Content.ReadAsStringAsync();

            if (!response.IsSuccessStatusCode)
            {
                var error = JsonConvert.DeserializeObject<ApiErrorResponse>(responseJson);
                throw new PLMApiException(error.Message, response.StatusCode, error.Error);
            }

            return JsonConvert.DeserializeObject<CreateDocumentResultDto>(responseJson);
        }

        #endregion

        #region Search Operations

        /// <summary>
        /// Show open document UI dialog and return ObjectId
        /// </summary>
        public async Task<OpenUIResultDto> OpenUIAsync()
        {
            var response = await _httpClient.PostAsync("/api/document/open-ui", null);

            var responseJson = await response.Content.ReadAsStringAsync();

            if (!response.IsSuccessStatusCode)
            {
                var error = JsonConvert.DeserializeObject<ApiErrorResponse>(responseJson);
                throw new PLMApiException(error?.Message ?? "Open UI failed", response.StatusCode, error?.Error);
            }

            var result = JsonConvert.DeserializeObject<dynamic>(responseJson);

            // Check if user cancelled
            if (result.success == false)
            {
                return null; // User cancelled
            }

            return new OpenUIResultDto
            {
                Success = result.success,
                ObjectId = result.objectId
            };
        }

        /// <summary>
        /// Show search UI dialog and return selected document
        /// </summary>
        public async Task<SearchUIResultDto> SearchUIAsync()
        {
            var response = await _httpClient.PostAsync("/api/search/search-ui", null);

            var responseJson = await response.Content.ReadAsStringAsync();

            if (!response.IsSuccessStatusCode)
            {
                var error = JsonConvert.DeserializeObject<ApiErrorResponse>(responseJson);
                throw new PLMApiException(error?.Message ?? "Search UI failed", response.StatusCode, error?.Error);
            }

            var result = JsonConvert.DeserializeObject<dynamic>(responseJson);

            // Check if user cancelled
            if (result.success == false)
            {
                return null; // User cancelled
            }

            return new SearchUIResultDto
            {
                Success = result.success,
                ObjectId = result.objectId,
                FileName = result.fileName
            };
        }

        /// <summary>
        /// Search for documents
        /// </summary>
        public async Task<DocumentSearchResponse> SearchDocumentsAsync(
            string objectId = null,
            string fileName = null,
            string project = null,
            string owner = null,
            string status = null)
        {
            var queryParams = new System.Collections.Generic.List<string>();

            if (!string.IsNullOrEmpty(objectId))
                queryParams.Add($"objectId={Uri.EscapeDataString(objectId)}");
            if (!string.IsNullOrEmpty(fileName))
                queryParams.Add($"fileName={Uri.EscapeDataString(fileName)}");
            if (!string.IsNullOrEmpty(project))
                queryParams.Add($"project={Uri.EscapeDataString(project)}");
            if (!string.IsNullOrEmpty(owner))
                queryParams.Add($"owner={Uri.EscapeDataString(owner)}");
            if (!string.IsNullOrEmpty(status))
                queryParams.Add($"status={Uri.EscapeDataString(status)}");

            var queryString = queryParams.Count > 0 ? "?" + string.Join("&", queryParams) : "";
            var url = $"/api/documents/search{queryString}";

            var response = await _httpClient.GetAsync(url);
            response.EnsureSuccessStatusCode();

            var json = await response.Content.ReadAsStringAsync();
            return JsonConvert.DeserializeObject<DocumentSearchResponse>(json);
        }

        /// <summary>
        /// Open a document from PLM (downloads and returns document info)
        /// </summary>
        public async Task<OpenDocumentResultDto> OpenDocumentAsync(string objectId, string revision = null)
        {
            // Get document info
            var document = await GetDocumentAsync(objectId);

            // Download file content
            var fileContent = await DownloadFileAsync(objectId, revision);

            return new OpenDocumentResultDto
            {
                ObjectId = document.ObjectId,
                FileName = document.FileName,
                Revision = document.CurrentRevision,
                FileContent = fileContent,
                IsCheckedOut = document.IsCheckedOut,
                CheckedOutBy = document.CheckedOutBy,
                Status = document.Status
            };
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
