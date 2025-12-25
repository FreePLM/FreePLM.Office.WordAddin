using System;
using System.IO;
using FreePLM.Office.WordAddin.Models;
using Newtonsoft.Json;

namespace FreePLM.Office.WordAddin.Helpers
{
    /// <summary>
    /// Helper class for reading and writing .dat sidecar files
    /// </summary>
    public static class DatFileHelper
    {
        private static readonly JsonSerializerSettings JsonSettings = new JsonSerializerSettings
        {
            Formatting = Formatting.Indented,
            NullValueHandling = NullValueHandling.Ignore
        };

        /// <summary>
        /// Write a .dat sidecar file to the specified directory
        /// </summary>
        /// <param name="directory">Directory where the file is located</param>
        /// <param name="sidecar">Sidecar metadata</param>
        /// <returns>True if successful</returns>
        public static bool WriteDatFile(string directory, DatFileSidecar sidecar)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(directory) || sidecar == null || string.IsNullOrWhiteSpace(sidecar.ObjectId))
                {
                    return false;
                }

                // Ensure directory exists
                if (!Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                var datFilePath = Path.Combine(directory, $"{sidecar.ObjectId}.dat");
                var json = JsonConvert.SerializeObject(sidecar, JsonSettings);

                File.WriteAllText(datFilePath, json);
                return true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error writing .dat file: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Read a .dat sidecar file from the specified directory
        /// </summary>
        /// <param name="directory">Directory where the file is located</param>
        /// <param name="objectId">ObjectId to construct the .dat filename</param>
        /// <returns>Sidecar metadata, or null if not found</returns>
        public static DatFileSidecar ReadDatFile(string directory, string objectId)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(directory) || string.IsNullOrWhiteSpace(objectId))
                {
                    return null;
                }

                var datFilePath = Path.Combine(directory, $"{objectId}.dat");

                if (!File.Exists(datFilePath))
                {
                    return null;
                }

                var json = File.ReadAllText(datFilePath);
                return JsonConvert.DeserializeObject<DatFileSidecar>(json);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error reading .dat file: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// Check if a .dat file exists for the given ObjectId
        /// </summary>
        public static bool DatFileExists(string directory, string objectId)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(directory) || string.IsNullOrWhiteSpace(objectId))
                {
                    return false;
                }

                var datFilePath = Path.Combine(directory, $"{objectId}.dat");
                return File.Exists(datFilePath);
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Find ObjectId by scanning for .dat files in a directory
        /// Useful when you have a file but don't know its ObjectId
        /// </summary>
        /// <param name="filePath">Full path to the managed file</param>
        /// <returns>ObjectId if found, or null</returns>
        public static string FindObjectIdByFilePath(string filePath)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(filePath))
                {
                    return null;
                }

                var directory = Path.GetDirectoryName(filePath);
                if (string.IsNullOrWhiteSpace(directory) || !Directory.Exists(directory))
                {
                    return null;
                }

                var fileName = Path.GetFileName(filePath);

                // Find all .dat files in the directory
                var datFiles = Directory.GetFiles(directory, "*.dat");

                foreach (var datFile in datFiles)
                {
                    try
                    {
                        var objectId = Path.GetFileNameWithoutExtension(datFile);
                        var sidecar = ReadDatFile(directory, objectId);

                        if (sidecar != null)
                        {
                            // Check if this .dat file references our file
                            var sidecarFileName = $"{sidecar.ObjectId}_{sidecar.CurrentRevision.Replace(".", "_")}{sidecar.FileExtension}";

                            if (fileName.Equals(sidecarFileName, StringComparison.OrdinalIgnoreCase) ||
                                fileName.StartsWith(sidecar.ObjectId, StringComparison.OrdinalIgnoreCase))
                            {
                                return objectId;
                            }
                        }
                    }
                    catch
                    {
                        // Skip this .dat file and continue
                        continue;
                    }
                }

                return null;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error finding ObjectId: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// Delete a .dat file
        /// </summary>
        public static bool DeleteDatFile(string directory, string objectId)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(directory) || string.IsNullOrWhiteSpace(objectId))
                {
                    return false;
                }

                var datFilePath = Path.Combine(directory, $"{objectId}.dat");

                if (File.Exists(datFilePath))
                {
                    File.Delete(datFilePath);
                    return true;
                }

                return false;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error deleting .dat file: {ex.Message}");
                return false;
            }
        }
    }
}
