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
