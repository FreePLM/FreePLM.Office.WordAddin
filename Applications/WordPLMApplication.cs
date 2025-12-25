using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;

namespace FreePLM.Office.WordAddin.Applications
{
    /// <summary>
    /// Word implementation of the IPLMApplication interface.
    /// Provides Word-specific integration with the FreePLM system.
    /// </summary>
    public class WordPLMApplication : IPLMApplication
    {
        public string ApplicationName => "Microsoft Word";

        public IEnumerable<string> SupportedExtensions => new[]
        {
            ".docx", ".doc",   // Word documents
            ".dotx", ".dot",   // Word templates
            ".docm", ".dotm"   // Macro-enabled documents and templates
        };

        public Task<bool> IsAvailableAsync()
        {
            try
            {
                // Check if Word is installed and accessible
                return Task.FromResult(Globals.ThisAddIn?.Application != null);
            }
            catch
            {
                return Task.FromResult(false);
            }
        }

        public Task<bool> OpenFileAsync(string filePath, string objectId, bool readOnly)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(filePath) || !File.Exists(filePath))
                {
                    return Task.FromResult(false);
                }

                var wordApp = Globals.ThisAddIn.Application;
                var doc = wordApp.Documents.Open(filePath, ReadOnly: readOnly);

                // Note: The PLM ribbon automatically reads the .dat file for metadata display
                return Task.FromResult(doc != null);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error opening Word file: {ex.Message}");
                return Task.FromResult(false);
            }
        }

        public Task<bool> CreateNewFileAsync(string filePath, string objectId)
        {
            try
            {
                var wordApp = Globals.ThisAddIn.Application;
                var doc = wordApp.Documents.Add();

                // Save the new document
                doc.SaveAs2(filePath);

                return Task.FromResult(true);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error creating Word file: {ex.Message}");
                return Task.FromResult(false);
            }
        }

        public Task<bool> SaveActiveDocumentAsync()
        {
            try
            {
                var wordApp = Globals.ThisAddIn.Application;

                if (wordApp.Documents.Count > 0)
                {
                    wordApp.ActiveDocument.Save();
                    return Task.FromResult(true);
                }

                return Task.FromResult(false);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error saving Word document: {ex.Message}");
                return Task.FromResult(false);
            }
        }

        public Task<string> GetActiveDocumentPathAsync()
        {
            try
            {
                var wordApp = Globals.ThisAddIn.Application;

                if (wordApp.Documents.Count > 0)
                {
                    return Task.FromResult(wordApp.ActiveDocument.FullName);
                }

                return Task.FromResult<string>(null);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error getting active document path: {ex.Message}");
                return Task.FromResult<string>(null);
            }
        }

        public Task<bool> CloseFileAsync(string filePath, bool saveChanges)
        {
            try
            {
                var wordApp = Globals.ThisAddIn.Application;

                foreach (Word.Document doc in wordApp.Documents)
                {
                    if (doc.FullName.Equals(filePath, StringComparison.OrdinalIgnoreCase))
                    {
                        doc.Close(saveChanges);
                        return Task.FromResult(true);
                    }
                }

                return Task.FromResult(false);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error closing Word document: {ex.Message}");
                return Task.FromResult(false);
            }
        }

        public Task<bool> IsFileOpenAsync(string filePath)
        {
            try
            {
                var wordApp = Globals.ThisAddIn.Application;

                foreach (Word.Document doc in wordApp.Documents)
                {
                    if (doc.FullName.Equals(filePath, StringComparison.OrdinalIgnoreCase))
                    {
                        return Task.FromResult(true);
                    }
                }

                return Task.FromResult(false);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error checking if Word document is open: {ex.Message}");
                return Task.FromResult(false);
            }
        }

        public Task<bool> ExportFileAsync(string sourcePath, string targetPath, string targetFormat)
        {
            try
            {
                var wordApp = Globals.ThisAddIn.Application;
                var doc = wordApp.Documents.Open(sourcePath);

                try
                {
                    switch (targetFormat.ToLower())
                    {
                        case "pdf":
                            doc.ExportAsFixedFormat(
                                targetPath,
                                Word.WdExportFormat.wdExportFormatPDF,
                                OpenAfterExport: false);
                            break;

                        case "html":
                        case "htm":
                            doc.SaveAs2(
                                targetPath,
                                FileFormat: Word.WdSaveFormat.wdFormatHTML);
                            break;

                        case "txt":
                            doc.SaveAs2(
                                targetPath,
                                FileFormat: Word.WdSaveFormat.wdFormatText);
                            break;

                        default:
                            doc.Close(false);
                            return Task.FromResult(false);
                    }

                    doc.Close(false);
                    return Task.FromResult(true);
                }
                finally
                {
                    // Ensure document is closed
                    if (doc != null)
                    {
                        try { doc.Close(false); } catch { }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error exporting Word document: {ex.Message}");
                return Task.FromResult(false);
            }
        }
    }
}
