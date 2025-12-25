using System;
using System.Collections.Generic;
using Newtonsoft.Json;

namespace FreePLM.Office.WordAddin.Models
{
    /// <summary>
    /// Represents the local .dat sidecar file that accompanies any PLM-managed file.
    /// This file makes the managed file self-describing without needing to embed
    /// custom properties in application-specific formats.
    /// </summary>
    public class DatFileSidecar
    {
        [JsonProperty("objectId")]
        public string ObjectId { get; set; }

        [JsonProperty("fileName")]
        public string FileName { get; set; }

        [JsonProperty("fileExtension")]
        public string FileExtension { get; set; }

        [JsonProperty("currentRevision")]
        public string CurrentRevision { get; set; }

        [JsonProperty("status")]
        public string Status { get; set; }

        [JsonProperty("owner")]
        public string Owner { get; set; }

        [JsonProperty("isCheckedOut")]
        public bool IsCheckedOut { get; set; }

        [JsonProperty("checkedOutBy")]
        public string CheckedOutBy { get; set; }

        [JsonProperty("checkedOutDate")]
        public DateTime? CheckedOutDate { get; set; }

        [JsonProperty("machineName")]
        public string MachineName { get; set; }

        [JsonProperty("lastSyncDate")]
        public DateTime LastSyncDate { get; set; }

        [JsonProperty("vaultPath")]
        public string VaultPath { get; set; }

        [JsonProperty("project")]
        public string Project { get; set; }

        [JsonProperty("group")]
        public string Group { get; set; }

        [JsonProperty("role")]
        public string Role { get; set; }

        [JsonProperty("customAttributes")]
        public Dictionary<string, string> CustomAttributes { get; set; }
    }
}
