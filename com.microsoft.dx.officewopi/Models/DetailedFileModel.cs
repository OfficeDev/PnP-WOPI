using com.microsoft.dx.officewopi.Models.Wopi;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace com.microsoft.dx.officewopi.Models
{
    /// <summary>
    /// This class contains additional file properties that are used in 
    /// CheckFileInfo requests, but not persisted in the database.
    /// 
    /// Note: many of the properties are hard-coded for this WOPI Host and
    /// might be more dynamic in other implementations
    /// </summary>
    public class DetailedFileModel : FileModel
    {
        [JsonProperty(PropertyName = "UserId")]
        public string UserId { get; set; }

        [JsonProperty(PropertyName = "CloseUrl")]
        public string CloseUrl { get; set; }

        [JsonProperty(PropertyName = "HostEditUrl")]
        public string HostEditUrl { get; set; }

        [JsonProperty(PropertyName = "HostViewUrl")]
        public string HostViewUrl { get; set; }

        [JsonProperty(PropertyName = "SupportsCoauth")]
        public bool SupportsCoauth
        {
            get { return false; }
        }

        [JsonProperty(PropertyName = "SupportsExtendedLockLength")]
        public bool SupportsExtendedLockLength
        {
            get { return false; }
        }

        [JsonProperty(PropertyName = "SupportsFileCreation")]
        public bool SupportsFileCreation
        {
            get { return false; }
        }

        [JsonProperty(PropertyName = "SupportsFolders")]
        public bool SupportsFolders
        {
            get { return false; }
        }

        [JsonProperty(PropertyName = "SupportsGetLock")]
        public bool SupportsGetLock
        {
            get { return true; }
        }

        [JsonProperty(PropertyName = "SupportsLocks")]
        public bool SupportsLocks
        {
            get { return true; }
        }

        [JsonProperty(PropertyName = "SupportsRename")]
        public bool SupportsRename
        {
            get { return true; }
        }

        [JsonProperty(PropertyName = "SupportsScenarioLinks")]
        public bool SupportsScenarioLinks
        {
            get { return false; }
        }

        [JsonProperty(PropertyName = "SupportsSecureStore")]
        public bool SupportsSecureStore
        {
            get { return false; }
        }

        [JsonProperty(PropertyName = "SupportsUpdate")]
        public bool SupportsUpdate
        {
            get { return true; }
        }

        [JsonProperty(PropertyName = "SupportsUserInfo")]
        public bool SupportsUserInfo
        {
            get { return true; }
        }

        [JsonProperty(PropertyName = "LicensesCheckForEditIsEnabled")]
        public bool LicensesCheckForEditIsEnabled
        {
            get { return true; }
        }

        /// <summary>
        /// Permissions for documents
        /// </summary>
        [JsonProperty(PropertyName = "ReadOnly")]
        public bool ReadOnly
        {
            get { return false; }
        }

        [JsonProperty(PropertyName = "RestrictedWebViewOnly")]
        public bool RestrictedWebViewOnly
        {
            get { return false; }
        }

        [JsonProperty(PropertyName = "UserCanAttend")] //Broadcast only
        public bool UserCanAttend
        {
            get { return false; }
        }

        [JsonProperty(PropertyName = "UserCanNotWriteRelative")]
        public bool UserCanNotWriteRelative
        {
            get { return false; }
        }

        [JsonProperty(PropertyName = "UserCanPresent")] //Broadcast only
        public bool UserCanPresent
        {
            get { return false; }
        }

        [JsonProperty(PropertyName = "UserCanRename")]
        public bool UserCanRename
        {
            get { return true; }
        }

        [JsonProperty(PropertyName = "UserCanWrite")]
        public bool UserCanWrite
        {
            get { return true; }
        }

        [JsonProperty(PropertyName = "WebEditingDisabled")]
        public bool WebEditingDisabled
        {
            get { return false; }
        }

        [JsonProperty(PropertyName = "Actions")]
        public List<WopiAction> Actions { get; set; }
    }
}
