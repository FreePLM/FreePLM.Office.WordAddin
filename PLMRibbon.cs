using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using FreePLM.Office.WordAddin.ApiClient;
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

                // Get ObjectId from user or document custom property
                var objectId = GetObjectIdFromDocument();
                if (string.IsNullOrEmpty(objectId))
                {
                    objectId = PromptForObjectId();
                    if (string.IsNullOrEmpty(objectId)) return;
                }

                // Prompt for comment
                var comment = PromptForComment("Check Out", "Enter reason for checking out this document:");

                // Perform checkout
                var result = await _apiClient.CheckOutAsync(objectId, comment);

                if (result.Success)
                {
                    // Download the file
                    var fileBytes = await _apiClient.DownloadFileAsync(result.ObjectId);

                    // Save to temp location
                    var tempPath = Path.Combine(Path.GetTempPath(), "FreePLM", result.ObjectId);
                    Directory.CreateDirectory(tempPath);

                    var doc = await _apiClient.GetDocumentAsync(result.ObjectId);
                    var filePath = Path.Combine(tempPath, doc.FileName);
                    File.WriteAllBytes(filePath, fileBytes);

                    // Close current document if open
                    if (Globals.ThisAddIn.Application.Documents.Count > 0)
                    {
                        Globals.ThisAddIn.Application.ActiveDocument.Close(false);
                    }

                    // Open the file in Word
                    var wordDoc = Globals.ThisAddIn.Application.Documents.Open(filePath);

                    // Store ObjectId in document custom properties
                    SetDocumentProperty(wordDoc, "PLM_ObjectId", result.ObjectId);
                    SetDocumentProperty(wordDoc, "PLM_Revision", result.Revision);
                    SetDocumentProperty(wordDoc, "PLM_CheckedOut", "true");

                    _currentObjectId = result.ObjectId;
                    _currentDocument = doc;

                    UpdateRibbonState();

                    MessageBox.Show(
                        $"Document checked out successfully!\n\nObjectId: {result.ObjectId}\nRevision: {result.Revision}",
                        "Check Out Successful",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                }
            }
            catch (PLMApiException ex)
            {
                if (ex.IsDocumentLocked)
                {
                    MessageBox.Show(
                        $"Document is already checked out.\n\n{ex.Message}",
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

                // Prompt for comment
                var comment = PromptForComment("Check In", "Enter description of changes:");
                if (string.IsNullOrEmpty(comment)) return;

                // Ask if major revision
                var createMajor = MessageBox.Show(
                    "Create a major revision?\n\nYes = Major revision (A→B)\nNo = Minor revision (.01→.02)",
                    "Revision Type",
                    MessageBoxButtons.YesNoCancel,
                    MessageBoxIcon.Question);

                if (createMajor == DialogResult.Cancel) return;

                // Save the document
                var doc = Globals.ThisAddIn.Application.ActiveDocument;
                doc.Save();

                // Read file content
                var filePath = doc.FullName;
                var fileBytes = File.ReadAllBytes(filePath);

                // Perform checkin
                var result = await _apiClient.CheckInAsync(
                    objectId,
                    fileBytes,
                    comment,
                    createMajor == DialogResult.Yes);

                if (result.Success)
                {
                    // Update document properties
                    SetDocumentProperty(doc, "PLM_Revision", result.NewRevision);
                    SetDocumentProperty(doc, "PLM_CheckedOut", "false");

                    UpdateRibbonState();

                    MessageBox.Show(
                        $"Document checked in successfully!\n\nNew Revision: {result.NewRevision}\nPrevious: {result.PreviousRevision}",
                        "Check In Successful",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);

                    // Close the document
                    doc.Close(false);
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

        public void NewButton_Click(IRibbonControl control)
        {
            MessageBox.Show(
                "New Document: This will create a blank Word document and register it in FreePLM.\n\n" +
                "You'll be prompted for: Group, Role, Project, and initial comment.\n\n" +
                "Feature coming in backend implementation!",
                "New PLM Document",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
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

        public void OpenButton_Click(IRibbonControl control)
        {
            MessageBox.Show(
                "Open: This will show a dialog to browse and open existing PLM documents.\n\n" +
                "You can search by: ObjectId, File Name, Project, Owner, Status.\n\n" +
                "Feature coming in backend implementation!",
                "Open PLM Document",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
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

                // Check if checked out
                var doc = await _apiClient.GetDocumentAsync(objectId);
                if (!doc.IsCheckedOut)
                {
                    MessageBox.Show(
                        "Document must be checked out before you can save changes.\n\nPlease check out the document first.",
                        "Not Checked Out",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                    return;
                }

                // Save the Word document
                Globals.ThisAddIn.Application.ActiveDocument.Save();

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

        public void SaveAsButton_Click(IRibbonControl control)
        {
            MessageBox.Show(
                "Save As: This will save the current document as a NEW PLM document with a new ObjectId.\n\n" +
                "The original document will remain unchanged.\n\n" +
                "You'll be prompted for: Group, Role, Project, and comment.\n\n" +
                "Feature coming in backend implementation!",
                "Save As New Document",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
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

        public void SearchButton_Click(IRibbonControl control)
        {
            MessageBox.Show(
                "Search: Find documents in FreePLM by:\n" +
                "- ObjectId or File Name\n" +
                "- Project, Group, Role\n" +
                "- Owner\n" +
                "- Status\n" +
                "- Date range\n\n" +
                "Results will be displayed in a searchable grid.\n\n" +
                "Feature coming in backend implementation!",
                "Search Documents",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
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
                    return GetDocumentProperty(doc, "PLM_ObjectId");
                }
            }
            catch { }
            return null;
        }

        private string GetDocumentProperty(Word.Document doc, string propertyName)
        {
            try
            {
                foreach (Microsoft.Office.Core.DocumentProperty prop in doc.CustomDocumentProperties)
                {
                    if (prop.Name == propertyName)
                    {
                        return prop.Value.ToString();
                    }
                }
            }
            catch { }
            return null;
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
