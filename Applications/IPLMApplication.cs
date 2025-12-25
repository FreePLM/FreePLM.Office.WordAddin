using System.Collections.Generic;
using System.Threading.Tasks;

namespace FreePLM.Office.WordAddin.Applications
{
    /// <summary>
    /// Interface that any application (Word, Excel, AutoCAD, etc.) must implement
    /// to integrate with the FreePLM system. This enables a pluggable architecture
    /// where new application support can be added by implementing this interface.
    /// </summary>
    public interface IPLMApplication
    {
        /// <summary>
        /// Application identifier (e.g., "Word", "Excel", "AutoCAD", "SolidWorks")
        /// </summary>
        string ApplicationName { get; }

        /// <summary>
        /// File extensions supported by this application (e.g., [".docx", ".doc"] for Word)
        /// </summary>
        IEnumerable<string> SupportedExtensions { get; }

        /// <summary>
        /// Check if the application is installed and available on this machine
        /// </summary>
        Task<bool> IsAvailableAsync();

        /// <summary>
        /// Open a PLM-managed file in the application
        /// </summary>
        /// <param name="filePath">Full path to the file</param>
        /// <param name="objectId">PLM ObjectId</param>
        /// <param name="readOnly">Whether to open as read-only</param>
        Task<bool> OpenFileAsync(string filePath, string objectId, bool readOnly);

        /// <summary>
        /// Create a new file from a template (if supported)
        /// </summary>
        /// <param name="filePath">Full path where the new file should be created</param>
        /// <param name="objectId">PLM ObjectId for the new document</param>
        Task<bool> CreateNewFileAsync(string filePath, string objectId);

        /// <summary>
        /// Save the currently active document (if applicable)
        /// </summary>
        Task<bool> SaveActiveDocumentAsync();

        /// <summary>
        /// Get the file path of the currently active document in this application
        /// </summary>
        /// <returns>Full path to active document, or null if none</returns>
        Task<string> GetActiveDocumentPathAsync();

        /// <summary>
        /// Close a file in the application
        /// </summary>
        /// <param name="filePath">Full path to the file to close</param>
        /// <param name="saveChanges">Whether to save changes before closing</param>
        Task<bool> CloseFileAsync(string filePath, bool saveChanges);

        /// <summary>
        /// Check if a specific file is currently open in the application
        /// </summary>
        /// <param name="filePath">Full path to the file</param>
        Task<bool> IsFileOpenAsync(string filePath);

        /// <summary>
        /// Export or convert a file to a different format (if supported)
        /// </summary>
        /// <param name="sourcePath">Source file path</param>
        /// <param name="targetPath">Target file path</param>
        /// <param name="targetFormat">Target format (e.g., "PDF", "HTML")</param>
        Task<bool> ExportFileAsync(string sourcePath, string targetPath, string targetFormat);
    }
}
