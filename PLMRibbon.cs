using System;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using FreePLM.Office.WordAddin.ApiClient;
using FreePLM.Office.WordAddin.Helpers;
using FreePLM.Office.WordAddin.Models;
using Microsoft.Office.Core;
using Word = Microsoft.Office.Interop.Word;

namespace FreePLM.Office.WordAddin
{
    [ComVisible(true)]
    public class PLMRibbon : IRibbonExtensibility
    {
        private IRibbonUI _ribbon;
        private PLMApiClient _apiClient;
        private string _currentObjectId;
        private DocumentReadDto _currentDocument;

        public PLMRibbon()
        {
            _apiClient = new PLMApiClient();
        }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("FreePLM.Office.WordAddin.PLMRibbon.xml");
        }

        #endregion

        #region Ribbon Callbacks

        public void Ribbon_Load(IRibbonUI ribbonUI)
        {
            _ribbon = ribbonUI;
        }

        #endregion

        #region Dynamic Label Callbacks

        public string GetObjectIdLabel(IRibbonControl control)
        {
            var objectId = GetObjectIdFromDocument();
            return string.IsNullOrEmpty(objectId) ? "No PLM Document" : $"ID: {objectId}";
        }

        public string GetRevisionLabel(IRibbonControl control)
        {
            try
            {
                var datFile = GetDatFileForActiveDocument();
                if (datFile != null)
                {
                    return $"Rev: {datFile.CurrentRevision}";
                }

                // Fallback to custom properties (for backwards compatibility)
                if (Globals.ThisAddIn.Application.Documents.Count > 0)
                {
                    var doc = Globals.ThisAddIn.Application.ActiveDocument;
                    var revision = GetDocumentProperty(doc, "PLM_Revision");
                    if (!string.IsNullOrEmpty(revision))
                    {
                        return $"Rev: {revision}";
                    }
                }
            }
            catch { }
            return "Revision: --";
        }

        public string GetStatusLabel(IRibbonControl control)
        {
            try
            {
                var datFile = GetDatFileForActiveDocument();
                if (datFile != null)
                {
                    return $"Status: {datFile.Status}";
                }

                // Fallback to custom properties (for backwards compatibility)
                if (Globals.ThisAddIn.Application.Documents.Count > 0)
                {
                    var doc = Globals.ThisAddIn.Application.ActiveDocument;
                    var status = GetDocumentProperty(doc, "PLM_Status");
                    if (!string.IsNullOrEmpty(status))
                    {
                        return $"Status: {status}";
                    }
                }
            }
            catch { }
            return "Status: --";
        }

        public string GetCheckedOutLabel(IRibbonControl control)
        {
            try
            {
                var datFile = GetDatFileForActiveDocument();
                if (datFile != null)
                {
                    if (datFile.IsCheckedOut)
                    {
                        return "✓ Checked Out";
                    }
                    else
                    {
                        return "✗ Read Only";
                    }
                }

                // Fallback to custom properties (for backwards compatibility)
                if (Globals.ThisAddIn.Application.Documents.Count > 0)
                {
                    var doc = Globals.ThisAddIn.Application.ActiveDocument;
                    var checkedOut = GetDocumentProperty(doc, "PLM_CheckedOut");

                    if (checkedOut == "true")
                    {
                        return "✓ Checked Out";
                    }
                    else if (!string.IsNullOrEmpty(checkedOut))
                    {
                        return "✗ Read Only";
                    }
                }
            }
            catch { }
            return "-- Not PLM --";
        }

        #endregion

        #region Button Click Handlers

        public async void CheckOutButton_Click(IRibbonControl control)
        {
            try
            {
                // Check if service is available
                if (!await _apiClient.IsServiceAvailableAsync())
                {
                    MessageBox.Show(
                        "Cannot connect to FreePLM service.\n\nPlease ensure the FreePLM service is running on localhost:5000",
                        "Service Unavailable",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                    return;
                }

                // Get ObjectId from document custom property
                var objectId = GetObjectIdFromDocument();
                var doc = Globals.ThisAddIn.Application.ActiveDocument;

                if (string.IsNullOrEmpty(objectId))
                {
                    MessageBox.Show(
                        "This document is not a PLM document.\n\nPlease open a PLM document first.",
                        "Not a PLM Document",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                    return;
                }

                // Check if already checked out locally
                var checkedOut = GetDocumentProperty(doc, "PLM_CheckedOut");
                if (checkedOut == "true")
                {
                    MessageBox.Show(
                        "Document is already checked out.",
                        "Already Checked Out",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                    return;
                }

                // Perform checkout without prompting for comment
                var result = await _apiClient.CheckOutAsync(objectId, "Checked out from Word Add-in");

                if (result.Success)
                {
                    // Download the file
                    var fileBytes = await _apiClient.DownloadFileAsync(result.ObjectId);

                    // Get document info
                    var docInfo = await _apiClient.GetDocumentAsync(result.ObjectId);

                    // Prepare file path
                    var tempPath = Path.Combine(Path.GetTempPath(), "FreePLM", result.ObjectId);
                    Directory.CreateDirectory(tempPath);
                    var filePath = Path.Combine(tempPath, docInfo.FileName);

                    // Close current document FIRST (so file is not locked)
                    doc.Close(false);

                    // Now save checked-out file to temp location
                    File.WriteAllBytes(filePath, fileBytes);

                    // Create .dat sidecar file (file is checked out)
                    CreateDatFile(filePath, docInfo, isCheckedOut: true);

                    // Open the checked out file in Word
                    var wordDoc = Globals.ThisAddIn.Application.Documents.Open(filePath);

                    // Store ObjectId in document custom properties
                    SetDocumentProperty(wordDoc, "PLM_ObjectId", result.ObjectId);
                    SetDocumentProperty(wordDoc, "PLM_Revision", result.Revision);
                    SetDocumentProperty(wordDoc, "PLM_CheckedOut", "true");

                    _currentObjectId = result.ObjectId;
                    _currentDocument = docInfo;

                    UpdateRibbonState();

                    // Show status in Word status bar instead of popup
                    Globals.ThisAddIn.Application.StatusBar = $"PLM: Document {result.ObjectId} checked out (Rev {result.Revision})";
                }
            }
            catch (PLMApiException ex)
            {
                if (ex.IsDocumentLocked)
                {
                    MessageBox.Show(
                        $"Document is already checked out by another user.\n\n{ex.Message}",
                        "Document Locked",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                }
                else
                {
                    MessageBox.Show(
                        $"Error checking out document:\n\n{ex.Message}",
                        "Error",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"Unexpected error:\n\n{ex.Message}",
                    "Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        public async void CheckInButton_Click(IRibbonControl control)
        {
            try
            {
                var doc = Globals.ThisAddIn.Application.ActiveDocument;
                var objectId = GetObjectIdFromDocument();

                if (string.IsNullOrEmpty(objectId))
                {
                    MessageBox.Show(
                        "This document is not checked out from FreePLM.",
                        "Not Checked Out",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                    return;
                }

                // Get current revision
                var currentRevision = GetDocumentProperty(doc, "PLM_Revision") ?? "A.01";
                var fileName = doc.Name;
                var originalPath = doc.FullName;

                // Save the document first
                doc.Save();

                // Prepare temp location
                var tempDir = Path.Combine(Path.GetTempPath(), "FreePLM", objectId);
                Directory.CreateDirectory(tempDir);
                var tempFilePath = Path.Combine(tempDir, fileName);

                // Check if document is already in the temp location
                var isInTempLocation = originalPath.Equals(tempFilePath, StringComparison.OrdinalIgnoreCase);

                if (!isInTempLocation)
                {
                    // Need to copy to temp location
                    File.Copy(originalPath, tempFilePath, true);
                }

                // Call the WPF UI-enabled API endpoint (shows dialog with comment + close option)
                var result = await _apiClient.CheckInUIAsync(objectId, fileName, currentRevision);

                if (result == null)
                {
                    // User cancelled in the WPF check-in dialog
                    return;
                }

                if (result.Success)
                {
                    // Update the .dat file to reflect checked-in status
                    var directory = Path.GetDirectoryName(originalPath);
                    var datFile = DatFileHelper.ReadDatFile(directory, objectId);
                    if (datFile != null)
                    {
                        datFile.IsCheckedOut = false;
                        datFile.CheckedOutBy = null;
                        datFile.CheckedOutDate = null;
                        datFile.CurrentRevision = result.NewRevision;
                        datFile.LastSyncDate = DateTime.UtcNow;
                        DatFileHelper.WriteDatFile(directory, datFile);
                    }

                    UpdateRibbonState();

                    // Close document if user requested it in the WPF dialog
                    if (result.CloseAfterCheckIn)
                    {
                        doc.Close(false);

                        MessageBox.Show(
                            $"Document checked in successfully!\n\nNew Revision: {result.NewRevision}\nPrevious: {result.PreviousRevision}\n\nDocument has been closed.",
                            "Check In Successful",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Information);
                    }
                    else
                    {
                        // Update custom properties to reflect new status
                        SetDocumentProperty(doc, "PLM_Revision", result.NewRevision);
                        SetDocumentProperty(doc, "PLM_CheckedOut", "false");

                        MessageBox.Show(
                            $"Document checked in successfully!\n\nNew Revision: {result.NewRevision}\nPrevious: {result.PreviousRevision}\n\nDocument remains open (read-only).",
                            "Check In Successful",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Information);
                    }
                }
            }
            catch (PLMApiException ex)
            {
                MessageBox.Show(
                    $"Error checking in document:\n\n{ex.Message}",
                    "Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"Unexpected error:\n\n{ex.Message}",
                    "Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        public async void CancelCheckOutButton_Click(IRibbonControl control)
        {
            try
            {
                var objectId = GetObjectIdFromDocument();
                if (string.IsNullOrEmpty(objectId))
                {
                    MessageBox.Show(
                        "This document is not checked out from FreePLM.",
                        "Not Checked Out",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                    return;
                }

                var confirm = MessageBox.Show(
                    "Are you sure you want to cancel the checkout?\n\nAny changes will be lost.",
                    "Confirm Cancel",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Warning);

                if (confirm == DialogResult.Yes)
                {
                    await _apiClient.CancelCheckOutAsync(objectId);

                    MessageBox.Show(
                        "Checkout cancelled successfully.",
                        "Cancelled",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);

                    // Close the document without saving
                    Globals.ThisAddIn.Application.ActiveDocument.Close(false);

                    UpdateRibbonState();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"Error cancelling checkout:\n\n{ex.Message}",
                    "Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        public async void RefreshButton_Click(IRibbonControl control)
        {
            await RefreshDocumentInfo();
        }

        public void ViewPropertiesButton_Click(IRibbonControl control)
        {
            var objectId = GetObjectIdFromDocument();
            if (string.IsNullOrEmpty(objectId) || _currentDocument == null)
            {
                MessageBox.Show(
                    "No PLM document information available.",
                    "No Document",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
                return;
            }

            var info = $"ObjectId: {_currentDocument.ObjectId}\n" +
                      $"File Name: {_currentDocument.FileName}\n" +
                      $"Revision: {_currentDocument.CurrentRevision}\n" +
                      $"Status: {_currentDocument.Status}\n" +
                      $"Owner: {_currentDocument.Owner}\n" +
                      $"Group: {_currentDocument.Group}\n" +
                      $"Role: {_currentDocument.Role}\n" +
                      $"Project: {_currentDocument.Project}\n" +
                      $"Created: {_currentDocument.CreatedDate:g} by {_currentDocument.CreatedBy}\n" +
                      $"Modified: {_currentDocument.ModifiedDate:g} by {_currentDocument.ModifiedBy}\n" +
                      $"Checked Out: {(_currentDocument.IsCheckedOut ? $"Yes, by {_currentDocument.CheckedOutBy}" : "No")}";

            MessageBox.Show(info, "Document Properties", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        public async void NewButton_Click(IRibbonControl control)
        {
            try
            {
                // Check if service is available
                if (!await _apiClient.IsServiceAvailableAsync())
                {
                    MessageBox.Show(
                        "Cannot connect to FreePLM service.\n\nPlease ensure the FreePLM service is running on localhost:5000",
                        "Service Unavailable",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                    return;
                }

                // Show UI dialog and create the document
                var result = await _apiClient.CreateDocumentUIAsync();

                if (result == null)
                {
                    // User cancelled in the UI
                    return;
                }

                if (!result.Success)
                {
                    MessageBox.Show(
                        $"Failed to create document:\n\n{result.Message}",
                        "Error",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                    return;
                }

                // Create a new Word document
                var app = Globals.ThisAddIn.Application;
                var doc = app.Documents.Add();

                // Set PLM properties from API response
                SetDocumentProperty(doc, "PLM_ObjectId", result.ObjectId);
                SetDocumentProperty(doc, "PLM_Revision", result.Revision);
                SetDocumentProperty(doc, "PLM_CheckedOut", "true");
                SetDocumentProperty(doc, "PLM_Status", "Private");

                // Save the document to temp location
                var tempDir = Path.Combine(Path.GetTempPath(), "FreePLM", result.ObjectId);
                Directory.CreateDirectory(tempDir);
                var filePath = Path.Combine(tempDir, result.FileName);

                doc.SaveAs2(filePath);

                // Create .dat sidecar file (new document is checked out)
                CreateDatFile(filePath, result.ObjectId, result.FileName, result.Revision, isCheckedOut: true);

                UpdateRibbonState();

                MessageBox.Show(
                    $"New PLM document created!\n\n" +
                    $"Object ID: {result.ObjectId}\n" +
                    $"File Name: {result.FileName}\n" +
                    $"Revision: {result.Revision}\n\n" +
                    $"The document is checked out and ready for editing.\n" +
                    $"Save your changes, then use 'Check In' to create the first revision.",
                    "Document Created",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"Error creating new document:\n\n{ex.Message}",
                    "Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        public void NewFromButton_Click(IRibbonControl control)
        {
            MessageBox.Show(
                "New From: This will let you select a local Word file and register it as a new PLM document.\n\n" +
                "You'll be prompted for: Local file, Group, Role, Project, and comment.\n\n" +
                "Feature coming in backend implementation!",
                "New From Local File",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
        }

        public async void OpenButton_Click(IRibbonControl control)
        {
            try
            {
                // Check if service is available
                if (!await _apiClient.IsServiceAvailableAsync())
                {
                    MessageBox.Show(
                        "Cannot connect to FreePLM service.\n\nPlease ensure the FreePLM service is running on localhost:5000",
                        "Service Unavailable",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                    return;
                }

                // Show WPF open dialog via the service
                var result = await _apiClient.OpenUIAsync();

                if (result != null && !string.IsNullOrEmpty(result.ObjectId))
                {
                    // Open the document
                    await OpenDocumentFromPLM(result.ObjectId);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"Error opening document: {ex.Message}",
                    "Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        public async void SaveButton_Click(IRibbonControl control)
        {
            try
            {
                var objectId = GetObjectIdFromDocument();
                if (string.IsNullOrEmpty(objectId))
                {
                    MessageBox.Show(
                        "This document is not a PLM document.\n\nUse 'Save As' to save it as a new PLM document.",
                        "Not a PLM Document",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                    return;
                }

                // Check if checked out (from document properties, not API)
                var doc = Globals.ThisAddIn.Application.ActiveDocument;
                var checkedOut = GetDocumentProperty(doc, "PLM_CheckedOut");
                if (checkedOut != "true")
                {
                    MessageBox.Show(
                        "Document must be checked out before you can save changes.\n\nPlease check out the document first.",
                        "Not Checked Out",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                    return;
                }

                // Save the Word document
                doc.Save();

                MessageBox.Show(
                    "Document saved locally.\n\nRemember to 'Check In' when you're done to create a new revision in PLM.",
                    "Saved",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error saving: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public async void SaveAsButton_Click(IRibbonControl control)
        {
            try
            {
                // Check if service is available
                if (!await _apiClient.IsServiceAvailableAsync())
                {
                    MessageBox.Show(
                        "Cannot connect to FreePLM service.\n\nPlease ensure the FreePLM service is running on localhost:5000",
                        "Service Unavailable",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                    return;
                }

                // Get active document
                var app = Globals.ThisAddIn.Application;
                if (app.Documents.Count == 0)
                {
                    MessageBox.Show(
                        "No document is currently open.",
                        "No Document",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                    return;
                }

                var doc = app.ActiveDocument;
                var currentFileName = doc.Name;

                // Check if document has been saved (has a valid path)
                string currentPath;
                bool isNewDocument = false;

                try
                {
                    currentPath = doc.FullName;
                    // If FullName doesn't throw and doesn't start with "Document", it's been saved
                    if (currentPath.StartsWith("Document") || !File.Exists(currentPath))
                    {
                        isNewDocument = true;
                    }
                }
                catch
                {
                    isNewDocument = true;
                    currentPath = null;
                }

                // If document hasn't been saved, save to temp location first
                if (isNewDocument)
                {
                    var tempPath = Path.Combine(Path.GetTempPath(), $"PLM_Temp_{Guid.NewGuid()}.docx");
                    doc.SaveAs2(tempPath);
                    currentPath = tempPath;
                }
                else
                {
                    // Save any pending changes
                    doc.Save();
                }

                // Check if document already has PLM properties
                var existingObjectId = GetDocumentProperty(doc, "PLM_ObjectId");
                if (!string.IsNullOrEmpty(existingObjectId))
                {
                    var dialogResult = MessageBox.Show(
                        $"This document is already in PLM (ObjectId: {existingObjectId}).\n\n" +
                        "Do you want to save it as a NEW document with a different ObjectId?\n\n" +
                        "The original document will remain unchanged.",
                        "Document Already in PLM",
                        MessageBoxButtons.YesNo,
                        MessageBoxIcon.Question);

                    if (dialogResult != DialogResult.Yes)
                    {
                        return;
                    }
                }

                // Show UI and save to PLM
                var result = await _apiClient.SaveAsUIAsync(currentPath, currentFileName);

                if (result == null)
                {
                    // User cancelled in the UI
                    return;
                }

                if (!result.Success)
                {
                    MessageBox.Show(
                        $"Failed to save document to PLM:\n\n{result.Message}",
                        "Error",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                    return;
                }

                // Set PLM properties on the document
                SetDocumentProperty(doc, "PLM_ObjectId", result.ObjectId);
                SetDocumentProperty(doc, "PLM_Revision", result.Revision);
                SetDocumentProperty(doc, "PLM_CheckedOut", "true");
                SetDocumentProperty(doc, "PLM_Status", "Private");

                // Save to temp location with new filename
                var tempDir = Path.Combine(Path.GetTempPath(), "FreePLM", result.ObjectId);
                Directory.CreateDirectory(tempDir);
                var newFilePath = Path.Combine(tempDir, result.FileName);

                doc.SaveAs2(newFilePath);

                UpdateRibbonState();

                MessageBox.Show(
                    $"Document saved to PLM successfully!\n\n" +
                    $"Object ID: {result.ObjectId}\n" +
                    $"File Name: {result.FileName}\n" +
                    $"Revision: {result.Revision}\n\n" +
                    $"The document is checked out and ready for editing.",
                    "Save As Successful",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"Error saving document to PLM:\n\n{ex.Message}",
                    "Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        public async void GetLatestButton_Click(IRibbonControl control)
        {
            try
            {
                var objectId = GetObjectIdFromDocument();
                if (string.IsNullOrEmpty(objectId))
                {
                    objectId = PromptForObjectId();
                    if (string.IsNullOrEmpty(objectId)) return;
                }

                // Download latest version (read-only)
                var fileBytes = await _apiClient.DownloadFileAsync(objectId);
                var doc = await _apiClient.GetDocumentAsync(objectId);

                // Save to temp location
                var tempPath = Path.Combine(Path.GetTempPath(), "FreePLM_ReadOnly", objectId);
                Directory.CreateDirectory(tempPath);

                var filePath = Path.Combine(tempPath, doc.FileName);
                File.WriteAllBytes(filePath, fileBytes);

                // Set file to read-only
                File.SetAttributes(filePath, FileAttributes.ReadOnly);

                // Close current document if open
                if (Globals.ThisAddIn.Application.Documents.Count > 0)
                {
                    Globals.ThisAddIn.Application.ActiveDocument.Close(false);
                }

                // Open as read-only
                var wordDoc = Globals.ThisAddIn.Application.Documents.Open(filePath, ReadOnly: true);

                MessageBox.Show(
                    $"Latest version opened as READ-ONLY.\n\nObjectId: {objectId}\nRevision: {doc.CurrentRevision}\n\nTo make changes, use 'Check Out'.",
                    "Get Latest",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public async void SubmitToWorkflowButton_Click(IRibbonControl control)
        {
            try
            {
                var objectId = GetObjectIdFromDocument();
                if (string.IsNullOrEmpty(objectId))
                {
                    MessageBox.Show("No PLM document is open.", "No Document", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                var doc = await _apiClient.GetDocumentAsync(objectId);

                // Show status transition dialog
                var newStatus = ShowStatusTransitionDialog(doc.Status);
                if (!newStatus.HasValue) return;

                var comment = PromptForComment("Submit to Workflow", "Enter comment for status change:");
                if (string.IsNullOrEmpty(comment)) return;

                await _apiClient.ChangeStatusAsync(objectId, newStatus.Value, comment);

                MessageBox.Show(
                    $"Status changed successfully!\n\nOld Status: {doc.Status}\nNew Status: {newStatus.Value}",
                    "Status Changed",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);

                await RefreshDocumentInfo();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void ChangeOwnerButton_Click(IRibbonControl control)
        {
            MessageBox.Show(
                "Change Owner: Transfer document ownership to another user.\n\n" +
                "You'll be prompted for the new owner's email/username.\n\n" +
                "Feature coming in backend implementation!",
                "Change Owner",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
        }

        public void ViewHistoryButton_Click(IRibbonControl control)
        {
            MessageBox.Show(
                "View History: Display all revisions with:\n" +
                "- Revision number\n" +
                "- Date/Time\n" +
                "- User\n" +
                "- Comment\n" +
                "- File size\n\n" +
                "You can also open any previous revision to view.\n\n" +
                "Feature coming in backend implementation!",
                "Revision History",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
        }

        public async void SearchButton_Click(IRibbonControl control)
        {
            try
            {
                // Check if service is available
                if (!await _apiClient.IsServiceAvailableAsync())
                {
                    MessageBox.Show(
                        "Cannot connect to FreePLM service.\n\nPlease ensure the FreePLM service is running on localhost:5000",
                        "Service Unavailable",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                    return;
                }

                // Show WPF search dialog via the service
                var result = await _apiClient.SearchUIAsync();

                if (result != null && !string.IsNullOrEmpty(result.ObjectId))
                {
                    // User selected a document to open
                    await OpenDocumentFromPLM(result.ObjectId);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"Error opening search dialog: {ex.Message}",
                    "Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        public async void NewRevisionButton_Click(IRibbonControl control)
        {
            try
            {
                var objectId = GetObjectIdFromDocument();
                if (string.IsNullOrEmpty(objectId))
                {
                    MessageBox.Show(
                        "No PLM document is open.",
                        "No Document",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                    return;
                }

                // Get current document info
                var doc = await _apiClient.GetDocumentAsync(objectId);

                var confirm = MessageBox.Show(
                    $"Create a new MAJOR revision from current document?\n\n" +
                    $"Current Revision: {doc.CurrentRevision}\n" +
                    $"New Revision will be: {GetNextMajorRevision(doc.CurrentRevision)}\n\n" +
                    $"This will create a new major revision branch.",
                    "Create New Major Revision",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question);

                if (confirm != DialogResult.Yes) return;

                var comment = PromptForComment("New Revision", "Enter comment for new major revision:");
                if (string.IsNullOrEmpty(comment)) return;

                // Save current document
                Globals.ThisAddIn.Application.ActiveDocument.Save();

                // Read file content
                var filePath = Globals.ThisAddIn.Application.ActiveDocument.FullName;
                var fileBytes = File.ReadAllBytes(filePath);

                // Check in with major revision flag
                var result = await _apiClient.CheckInAsync(
                    objectId,
                    fileBytes,
                    comment,
                    createMajorRevision: true);

                if (result.Success)
                {
                    MessageBox.Show(
                        $"New major revision created successfully!\n\n" +
                        $"Old Revision: {result.PreviousRevision}\n" +
                        $"New Revision: {result.NewRevision}",
                        "New Revision Created",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);

                    await RefreshDocumentInfo();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"Error creating new revision:\n\n{ex.Message}",
                    "Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        #endregion

        #region Helper Methods

        private string GetObjectIdFromDocument()
        {
            try
            {
                if (Globals.ThisAddIn.Application.Documents.Count > 0)
                {
                    var doc = Globals.ThisAddIn.Application.ActiveDocument;

                    // First try to find .dat file by scanning directory
                    var filePath = doc.FullName;
                    var objectId = DatFileHelper.FindObjectIdByFilePath(filePath);

                    if (!string.IsNullOrEmpty(objectId))
                    {
                        return objectId;
                    }

                    // Fallback to custom properties (for backwards compatibility)
                    return GetDocumentProperty(doc, "PLM_ObjectId");
                }
            }
            catch { }
            return null;
        }

        /// <summary>
        /// Get the .dat sidecar file for the active document
        /// </summary>
        private DatFileSidecar GetDatFileForActiveDocument()
        {
            try
            {
                if (Globals.ThisAddIn.Application.Documents.Count > 0)
                {
                    var doc = Globals.ThisAddIn.Application.ActiveDocument;
                    var filePath = doc.FullName;
                    var directory = Path.GetDirectoryName(filePath);

                    // Find ObjectId
                    var objectId = DatFileHelper.FindObjectIdByFilePath(filePath);

                    if (!string.IsNullOrEmpty(objectId))
                    {
                        return DatFileHelper.ReadDatFile(directory, objectId);
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error reading .dat file: {ex.Message}");
            }

            return null;
        }

        private string GetDocumentProperty(Word.Document doc, string propertyName)
        {
            try
            {
                // Access by name directly to avoid COM casting issues
                var prop = doc.CustomDocumentProperties[propertyName];
                return prop?.Value?.ToString();
            }
            catch
            {
                return null;
            }
        }

        private void SetDocumentProperty(Word.Document doc, string propertyName, string value)
        {
            try
            {
                var properties = doc.CustomDocumentProperties;
                try
                {
                    properties[propertyName].Delete();
                }
                catch { }

                properties.Add(propertyName, false,
                    Microsoft.Office.Core.MsoDocProperties.msoPropertyTypeString, value);
            }
            catch { }
        }

        /// <summary>
        /// Create a .dat sidecar file for a PLM-managed file
        /// </summary>
        private void CreateDatFile(string filePath, DocumentReadDto docInfo, bool isCheckedOut = false)
        {
            try
            {
                var directory = Path.GetDirectoryName(filePath);
                if (string.IsNullOrWhiteSpace(directory))
                {
                    return;
                }

                var fileExtension = Path.GetExtension(docInfo.FileName);

                var sidecar = new DatFileSidecar
                {
                    ObjectId = docInfo.ObjectId,
                    FileName = docInfo.FileName,
                    FileExtension = fileExtension,
                    CurrentRevision = docInfo.CurrentRevision,
                    Status = docInfo.Status.ToString(),
                    Owner = docInfo.Owner,
                    IsCheckedOut = isCheckedOut,
                    CheckedOutBy = isCheckedOut ? docInfo.CheckedOutBy : null,
                    CheckedOutDate = isCheckedOut ? docInfo.CheckedOutDate : null,
                    MachineName = isCheckedOut ? Environment.MachineName : null,
                    LastSyncDate = DateTime.UtcNow,
                    VaultPath = $"Vault\\{docInfo.ObjectId}",
                    Project = docInfo.Project,
                    Group = docInfo.Group,
                    Role = docInfo.Role
                };

                DatFileHelper.WriteDatFile(directory, sidecar);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error creating .dat file: {ex.Message}");
            }
        }

        /// <summary>
        /// Create a .dat sidecar file for a newly created PLM document
        /// </summary>
        private void CreateDatFile(string filePath, string objectId, string fileName, string revision, bool isCheckedOut = false)
        {
            try
            {
                var directory = Path.GetDirectoryName(filePath);
                if (string.IsNullOrWhiteSpace(directory))
                {
                    return;
                }

                var fileExtension = Path.GetExtension(fileName);

                var sidecar = new DatFileSidecar
                {
                    ObjectId = objectId,
                    FileName = fileName,
                    FileExtension = fileExtension,
                    CurrentRevision = revision,
                    Status = "Private",
                    Owner = Environment.UserName,
                    IsCheckedOut = isCheckedOut,
                    CheckedOutBy = isCheckedOut ? Environment.UserName : null,
                    CheckedOutDate = isCheckedOut ? (DateTime?)DateTime.UtcNow : null,
                    MachineName = isCheckedOut ? Environment.MachineName : null,
                    LastSyncDate = DateTime.UtcNow,
                    VaultPath = $"Vault\\{objectId}"
                };

                DatFileHelper.WriteDatFile(directory, sidecar);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error creating .dat file: {ex.Message}");
            }
        }

        private string PromptForObjectId()
        {
            var form = new Form
            {
                Text = "Enter ObjectId",
                Width = 400,
                Height = 150,
                StartPosition = FormStartPosition.CenterScreen
            };

            var label = new Label { Text = "ObjectId:", Left = 20, Top = 20, Width = 100 };
            var textBox = new TextBox { Left = 130, Top = 20, Width = 230 };
            var okButton = new Button { Text = "OK", Left = 200, Top = 60, Width = 75, DialogResult = DialogResult.OK };
            var cancelButton = new Button { Text = "Cancel", Left = 285, Top = 60, Width = 75, DialogResult = DialogResult.Cancel };

            form.Controls.AddRange(new Control[] { label, textBox, okButton, cancelButton });
            form.AcceptButton = okButton;
            form.CancelButton = cancelButton;

            return form.ShowDialog() == DialogResult.OK ? textBox.Text : null;
        }

        private string PromptForComment(string title, string prompt)
        {
            var form = new Form
            {
                Text = title,
                Width = 400,
                Height = 200,
                StartPosition = FormStartPosition.CenterScreen
            };

            var label = new Label { Text = prompt, Left = 20, Top = 20, Width = 340, Height = 40 };
            var textBox = new TextBox { Left = 20, Top = 65, Width = 340, Height = 60, Multiline = true };
            var okButton = new Button { Text = "OK", Left = 200, Top = 130, Width = 75, DialogResult = DialogResult.OK };
            var cancelButton = new Button { Text = "Cancel", Left = 285, Top = 130, Width = 75, DialogResult = DialogResult.Cancel };

            form.Controls.AddRange(new Control[] { label, textBox, okButton, cancelButton });
            form.AcceptButton = okButton;
            form.CancelButton = cancelButton;

            return form.ShowDialog() == DialogResult.OK ? textBox.Text : null;
        }

        private bool? ShowCloseAfterCheckInDialog()
        {
            var form = new Form
            {
                Text = "Check In Document",
                Width = 420,
                Height = 180,
                StartPosition = FormStartPosition.CenterScreen,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                MaximizeBox = false,
                MinimizeBox = false
            };

            var label = new Label
            {
                Text = "The document will be checked in and marked as read-only.\n\nWhat would you like to do after check-in?",
                Left = 20,
                Top = 20,
                Width = 370,
                Height = 50
            };

            var checkBox = new CheckBox
            {
                Text = "Close document after check-in",
                Left = 20,
                Top = 80,
                Width = 250,
                Checked = false  // Default: keep document open
            };

            var okButton = new Button
            {
                Text = "Check In",
                Left = 210,
                Top = 110,
                Width = 90,
                DialogResult = DialogResult.OK
            };

            var cancelButton = new Button
            {
                Text = "Cancel",
                Left = 310,
                Top = 110,
                Width = 75,
                DialogResult = DialogResult.Cancel
            };

            form.Controls.AddRange(new Control[] { label, checkBox, okButton, cancelButton });
            form.AcceptButton = okButton;
            form.CancelButton = cancelButton;

            var result = form.ShowDialog();

            if (result == DialogResult.OK)
            {
                return checkBox.Checked;
            }

            return null; // User cancelled
        }

        private async System.Threading.Tasks.Task OpenDocumentFromPLM(string objectId)
        {
            try
            {
                // Download document from PLM
                var openResult = await _apiClient.OpenDocumentAsync(objectId);

                // Save to temp location
                var tempPath = Path.Combine(Path.GetTempPath(), "FreePLM", openResult.ObjectId);
                Directory.CreateDirectory(tempPath);
                var filePath = Path.Combine(tempPath, openResult.FileName);
                File.WriteAllBytes(filePath, openResult.FileContent);

                // Get full document info for .dat file creation
                var docInfo = await _apiClient.GetDocumentAsync(openResult.ObjectId);

                // Create .dat sidecar file (file may or may not be checked out)
                CreateDatFile(filePath, docInfo, isCheckedOut: openResult.IsCheckedOut);

                // Check if this document is already open
                bool alreadyOpen = false;
                if (Globals.ThisAddIn.Application.Documents.Count > 0)
                {
                    foreach (Word.Document doc in Globals.ThisAddIn.Application.Documents)
                    {
                        var existingObjectId = GetDocumentProperty(doc, "PLM_ObjectId");
                        if (existingObjectId == openResult.ObjectId)
                        {
                            // Document already open, just activate it
                            doc.Activate();
                            alreadyOpen = true;
                            Globals.ThisAddIn.Application.StatusBar = $"PLM: {openResult.ObjectId} is already open";
                            break;
                        }
                    }
                }

                if (alreadyOpen)
                {
                    return;
                }

                // Open the file in Word (alongside any other open documents)
                var wordDoc = Globals.ThisAddIn.Application.Documents.Open(filePath);

                // Store PLM properties in document
                SetDocumentProperty(wordDoc, "PLM_ObjectId", openResult.ObjectId);
                SetDocumentProperty(wordDoc, "PLM_Revision", openResult.Revision);
                SetDocumentProperty(wordDoc, "PLM_CheckedOut", openResult.IsCheckedOut ? "true" : "false");
                SetDocumentProperty(wordDoc, "PLM_Status", openResult.Status.ToString());

                // Save the document to persist custom properties
                wordDoc.Save();

                // Update current document tracking
                _currentObjectId = openResult.ObjectId;
                _currentDocument = await _apiClient.GetDocumentAsync(openResult.ObjectId);

                // Update ribbon state
                UpdateRibbonState();

                // Show status in Word status bar instead of popup
                var checkedOutStatus = openResult.IsCheckedOut ? $"Checked out by {openResult.CheckedOutBy}" : "Available";
                Globals.ThisAddIn.Application.StatusBar = $"PLM: {openResult.ObjectId} Rev {openResult.Revision} - {openResult.Status} - {checkedOutStatus}";
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"Error opening document from PLM: {ex.Message}",
                    "Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        private DocumentStatus? ShowStatusTransitionDialog(DocumentStatus currentStatus)
        {
            var form = new Form
            {
                Text = "Submit to Workflow",
                Width = 400,
                Height = 250,
                StartPosition = FormStartPosition.CenterScreen
            };

            var label = new Label
            {
                Text = $"Current Status: {currentStatus}\n\nSelect new status:",
                Left = 20,
                Top = 20,
                Width = 340,
                Height = 40
            };

            var listBox = new ListBox { Left = 20, Top = 65, Width = 340, Height = 100 };

            // Add valid transitions based on current status
            switch (currentStatus)
            {
                case DocumentStatus.Private:
                    listBox.Items.Add(DocumentStatus.InWork);
                    break;
                case DocumentStatus.InWork:
                    listBox.Items.Add(DocumentStatus.Frozen);
                    listBox.Items.Add(DocumentStatus.Private);
                    break;
                case DocumentStatus.Frozen:
                    listBox.Items.Add(DocumentStatus.Released);
                    listBox.Items.Add(DocumentStatus.InWork);
                    break;
                case DocumentStatus.Released:
                    listBox.Items.Add(DocumentStatus.Obsolete);
                    break;
                case DocumentStatus.Obsolete:
                    // No transitions from obsolete
                    MessageBox.Show("Obsolete documents cannot change status.", "Invalid Operation",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return null;
            }

            if (listBox.Items.Count == 0)
            {
                MessageBox.Show("No valid status transitions available.", "No Transitions",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return null;
            }

            listBox.SelectedIndex = 0;

            var okButton = new Button { Text = "OK", Left = 200, Top = 175, Width = 75, DialogResult = DialogResult.OK };
            var cancelButton = new Button { Text = "Cancel", Left = 285, Top = 175, Width = 75, DialogResult = DialogResult.Cancel };

            form.Controls.AddRange(new Control[] { label, listBox, okButton, cancelButton });
            form.AcceptButton = okButton;
            form.CancelButton = cancelButton;

            if (form.ShowDialog() == DialogResult.OK && listBox.SelectedItem != null)
            {
                return (DocumentStatus)listBox.SelectedItem;
            }

            return null;
        }

        private async System.Threading.Tasks.Task RefreshDocumentInfo()
        {
            try
            {
                var objectId = GetObjectIdFromDocument();
                if (!string.IsNullOrEmpty(objectId))
                {
                    _currentDocument = await _apiClient.GetDocumentAsync(objectId);
                    _currentObjectId = objectId;
                }
                UpdateRibbonState();
            }
            catch { }
        }

        private void UpdateRibbonState()
        {
            // This will be called to update ribbon button states
            // Ribbon controls will check document state
            _ribbon?.Invalidate();
        }

        private string GetNextMajorRevision(string currentRevision)
        {
            // Parse current revision (e.g., "A.03" -> "A")
            if (string.IsNullOrEmpty(currentRevision)) return "B.01";

            var parts = currentRevision.Split('.');
            if (parts.Length < 1) return "B.01";

            var majorLetter = parts[0];
            if (string.IsNullOrEmpty(majorLetter)) return "B.01";

            // Increment the letter (A->B, B->C, etc.)
            var nextLetter = (char)(majorLetter[0] + 1);
            return $"{nextLetter}.01";
        }

        private static string GetResourceText(string resourceName)
        {
            var asm = Assembly.GetExecutingAssembly();
            var resourceNames = asm.GetManifestResourceNames();
            foreach (var name in resourceNames)
            {
                if (string.Compare(resourceName, name, StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (var resourceReader = new StreamReader(asm.GetManifestResourceStream(name)))
                    {
                        return resourceReader.ReadToEnd();
                    }
                }
            }
            return null;
        }

        #endregion
    }
}
