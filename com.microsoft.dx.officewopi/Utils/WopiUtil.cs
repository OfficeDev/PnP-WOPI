using com.microsoft.dx.officewopi.Models;
using com.microsoft.dx.officewopi.Models.Wopi;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Net.Http;
using System.Runtime.Caching;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using System.Web;
using System.Xml.Linq;

namespace com.microsoft.dx.officewopi.Utils
{
    public static class WopiUtil
    {
        //WOPI protocol constants
        public const string WOPI_BASE_PATH = @"/wopi/";
        public const string WOPI_CHILDREN_PATH = @"/children";
        public const string WOPI_CONTENTS_PATH = @"/contents";
        public const string WOPI_FILES_PATH = @"files/";
        public const string WOPI_FOLDERS_PATH = @"folders/";

        /// <summary>
        /// Populates a list of files with action details from WOPI discovery
        /// </summary>
        public async static Task PopulateActions(this IEnumerable<DetailedFileModel> files)
        {
            if (files.Count() > 0)
            {
                foreach (var file in files)
                {
                    await file.PopulateActions();
                }
            }
        }

        /// <summary>
        /// Populates a file with action details from WOPI discovery based on the file extension
        /// </summary>
        public async static Task PopulateActions(this DetailedFileModel file)
        {
            // Get the discovery informations
            var actions = await GetDiscoveryInfo();
            var fileExt = file.BaseFileName.Substring(file.BaseFileName.LastIndexOf('.') + 1).ToLower();
            file.Actions = actions.Where(i => i.ext == fileExt).OrderBy(i => i.isDefault).ToList();
        }

        /// <summary>
        /// Gets the discovery information from WOPI discovery and caches it appropriately
        /// </summary>
        public async static Task<List<WopiAction>> GetDiscoveryInfo()
        {
            List<WopiAction> actions = new List<WopiAction>();

            // Determine if the discovery data is cached
            MemoryCache memoryCache = MemoryCache.Default;
            if (memoryCache.Contains("DiscoData"))
                actions = (List<WopiAction>)memoryCache["DiscoData"];
            else
            {
                // Data isn't cached, so we will use the Wopi Discovery endpoint to get the data
                HttpClient client = new HttpClient();
                using (HttpResponseMessage response = await client.GetAsync(ConfigurationManager.AppSettings["WopiDiscovery"]))
                {
                    if (response.IsSuccessStatusCode)
                    {
                        // Read the xml string from the response
                        string xmlString = await response.Content.ReadAsStringAsync();

                        // Parse the xml string into Xml
                        var discoXml = XDocument.Parse(xmlString);

                        // Convert the discovery xml into list of WopiApp
                        var xapps = discoXml.Descendants("app");
                        foreach (var xapp in xapps)
                        {
                            // Parse the actions for the app
                            var xactions = xapp.Descendants("action");
                            foreach (var xaction in xactions)
                            {
                                actions.Add(new WopiAction()
                                {
                                    app = xapp.Attribute("name").Value,
                                    favIconUrl = xapp.Attribute("favIconUrl").Value,
                                    checkLicense = Convert.ToBoolean(xapp.Attribute("checkLicense").Value),
                                    name = xaction.Attribute("name").Value,
                                    ext = (xaction.Attribute("ext") != null) ? xaction.Attribute("ext").Value : String.Empty,
                                    progid = (xaction.Attribute("progid") != null) ? xaction.Attribute("progid").Value : String.Empty,
                                    isDefault = (xaction.Attribute("default") != null) ? true : false,
                                    urlsrc = xaction.Attribute("urlsrc").Value,
                                    requires = (xaction.Attribute("requires") != null) ? xaction.Attribute("requires").Value : String.Empty
                                });
                            }

                            // Cache the discovey data for an hour
                            memoryCache.Add("DiscoData", actions, DateTimeOffset.Now.AddHours(1));
                        }
                    }
                }
            }

            return actions;
        }

        /// <summary>
        /// Forms the correct action url for the file and host
        /// </summary>
        public static string GetActionUrl(WopiAction action, FileModel file, string authority)
        {
            // Initialize the urlsrc
            var urlsrc = action.urlsrc;

            // Look through the action placeholders
            var phCnt = 0;
            foreach (var p in WopiUrlPlaceholders.Placeholders)
            {
                if (urlsrc.Contains(p))
                {
                    // Replace the placeholder value accordingly
                    var ph = WopiUrlPlaceholders.GetPlaceholderValue(p);
                    if (!String.IsNullOrEmpty(ph))
                    {
                        urlsrc = urlsrc.Replace(p, ph + "&");
                        phCnt++;
                    }
                    else
                        urlsrc = urlsrc.Replace(p, ph);
                }
            }

            // Add the WOPISrc to the end of the request
            urlsrc += ((phCnt > 0) ? "" : "?") + String.Format("WOPISrc=https://{0}/wopi/files/{1}", authority, file.id.ToString());
            return urlsrc;
        }

        /// <summary>
        /// Validates the WOPI Proof on an incoming WOPI request
        /// </summary>
        public async static Task<bool> ValidateWopiProof(HttpContext context)
        {
            // Make sure the request has the correct headers
            if (context.Request.Headers[WopiRequestHeaders.PROOF] == null ||
                context.Request.Headers[WopiRequestHeaders.TIME_STAMP] == null)
                return false;

            // Set the requested proof values
            var requestProof = context.Request.Headers[WopiRequestHeaders.PROOF];
            var requestProofOld = String.Empty;
            if (context.Request.Headers[WopiRequestHeaders.PROOF_OLD] != null)
                requestProofOld = context.Request.Headers[WopiRequestHeaders.PROOF_OLD];

            // Get the WOPI proof info from discovery
            var discoProof = await getWopiProof(context);

            // Encode the values into bytes
            var accessTokenBytes = Encoding.UTF8.GetBytes(context.Request.QueryString["access_token"]);
            var hostUrl = context.Request.Url.OriginalString.Replace(":44300", "").Replace(":443", "");
            var hostUrlBytes = Encoding.UTF8.GetBytes(hostUrl.ToUpperInvariant());
            var timeStampBytes = BitConverter.GetBytes(Convert.ToInt64(context.Request.Headers[WopiRequestHeaders.TIME_STAMP])).Reverse().ToArray();

            // Build expected proof
            List<byte> expected = new List<byte>(
                4 + accessTokenBytes.Length +
                4 + hostUrlBytes.Length +
                4 + timeStampBytes.Length);

            // Add the values to the expected variable
            expected.AddRange(BitConverter.GetBytes(accessTokenBytes.Length).Reverse().ToArray());
            expected.AddRange(accessTokenBytes);
            expected.AddRange(BitConverter.GetBytes(hostUrlBytes.Length).Reverse().ToArray());
            expected.AddRange(hostUrlBytes);
            expected.AddRange(BitConverter.GetBytes(timeStampBytes.Length).Reverse().ToArray());
            expected.AddRange(timeStampBytes);
            byte[] expectedBytes = expected.ToArray();

            return (verifyProof(expectedBytes, requestProof, discoProof.value) ||
                verifyProof(expectedBytes, requestProof, discoProof.oldvalue) ||
                verifyProof(expectedBytes, requestProofOld, discoProof.value));
        }

        /// <summary>
        /// Verifies the proof against a specified key
        /// </summary>
        private static bool verifyProof(byte[] expectedProof, string proofFromRequest, string proofFromDiscovery)
        {
            using (RSACryptoServiceProvider rsaProvider = new RSACryptoServiceProvider())
            {
                try
                {
                    rsaProvider.ImportCspBlob(Convert.FromBase64String(proofFromDiscovery));
                    return rsaProvider.VerifyData(expectedProof, "SHA256", Convert.FromBase64String(proofFromRequest));
                }
                catch (FormatException)
                {
                    return false;
                }
                catch (CryptographicException)
                {
                    return false;
                }
            }
        }

        /// <summary>
        /// Gets the WOPI proof details from the WOPI discovery endpoint and caches it appropriately
        /// </summary>
        internal async static Task<WopiProof> getWopiProof(HttpContext context)
        {
            WopiProof wopiProof = null;

            // Check cache for this data
            MemoryCache memoryCache = MemoryCache.Default;
            if (memoryCache.Contains("WopiProof"))
                wopiProof = (WopiProof)memoryCache["WopiProof"];
            else
            {
                HttpClient client = new HttpClient();
                using (HttpResponseMessage response = await client.GetAsync(ConfigurationManager.AppSettings["WopiDiscovery"]))
                {
                    if (response.IsSuccessStatusCode)
                    {
                        // Read the xml string from the response
                        string xmlString = await response.Content.ReadAsStringAsync();

                        // Parse the xml string into Xml
                        var discoXml = XDocument.Parse(xmlString);

                        // Convert the discovery xml into list of WopiApp
                        var proof = discoXml.Descendants("proof-key").FirstOrDefault();
                        wopiProof = new WopiProof()
                        {
                            value = proof.Attribute("value").Value,
                            modulus = proof.Attribute("modulus").Value,
                            exponent = proof.Attribute("exponent").Value,
                            oldvalue = proof.Attribute("oldvalue").Value,
                            oldmodulus = proof.Attribute("oldmodulus").Value,
                            oldexponent = proof.Attribute("oldexponent").Value
                        };

                        // Add to cache for 20min
                        memoryCache.Add("WopiProof", wopiProof, DateTimeOffset.Now.AddMinutes(20));
                    }
                }
            }

            return wopiProof;
        }
    }

    /// <summary>
    /// Contains valid WOPI response headers
    /// </summary>
    public class WopiResponseHeaders
    {
        //WOPI Header Consts
        public const string HOST_ENDPOINT = "X-WOPI-HostEndpoint";
        public const string INVALID_FILE_NAME_ERROR = "X-WOPI-InvalidFileNameError";
        public const string LOCK = "X-WOPI-Lock";
        public const string LOCK_FAILURE_REASON = "X-WOPI-LockFailureReason";
        public const string LOCKED_BY_OTHER_INTERFACE = "X-WOPI-LockedByOtherInterface";
        public const string MACHINE_NAME = "X-WOPI-MachineName";
        public const string PREF_TRACE = "X-WOPI-PerfTrace";
        public const string SERVER_ERROR = "X-WOPI-ServerError";
        public const string SERVER_VERSION = "X-WOPI-ServerVersion";
        public const string VALID_RELATIVE_TARGET = "X-WOPI-ValidRelativeTarget";
    }

    /// <summary>
    /// Contains valid WOPI request headers
    /// </summary>
    public class WopiRequestHeaders
    {
        //WOPI Header Consts
        public const string APP_ENDPOINT = "X-WOPI-AppEndpoint";
        public const string CLIENT_VERSION = "X-WOPI-ClientVersion";
        public const string CORRELATION_ID = "X-WOPI-CorrelationId";
        public const string LOCK = "X-WOPI-Lock";
        public const string MACHINE_NAME = "X-WOPI-MachineName";
        public const string MAX_EXPECTED_SIZE = "X-WOPI-MaxExpectedSize";
        public const string OLD_LOCK = "X-WOPI-OldLock";
        public const string OVERRIDE = "X-WOPI-Override";
        public const string OVERWRITE_RELATIVE_TARGET = "X-WOPI-OverwriteRelativeTarget";
        public const string PREF_TRACE_REQUESTED = "X-WOPI-PerfTraceRequested";
        public const string PROOF = "X-WOPI-Proof";
        public const string PROOF_OLD = "X-WOPI-ProofOld";
        public const string RELATIVE_TARGET = "X-WOPI-RelativeTarget";
        public const string REQUESTED_NAME = "X-WOPI-RequestedName";
        public const string SESSION_CONTEXT = "X-WOPI-SessionContext";
        public const string SIZE = "X-WOPI-Size";
        public const string SUGGESTED_TARGET = "X-WOPI-SuggestedTarget";
        public const string TIME_STAMP = "X-WOPI-TimeStamp";
    }

    /// <summary>
    /// Contains all valid URL placeholders for different WOPI actions
    /// </summary>
    public class WopiUrlPlaceholders
    {
        public static List<string> Placeholders = new List<string>() { BUSINESS_USER,
            DC_LLCC, DISABLE_ASYNC, DISABLE_CHAT, DISABLE_BROADCAST,
            EMBDDED, FULLSCREEN, PERFSTATS, RECORDING, THEME_ID, UI_LLCC, VALIDATOR_TEST_CATEGORY
        };
        public const string BUSINESS_USER = "<IsLicensedUser=BUSINESS_USER&>";
        public const string DC_LLCC = "<rs=DC_LLCC&>";
        public const string DISABLE_ASYNC = "<na=DISABLE_ASYNC&>";
        public const string DISABLE_CHAT = "<dchat=DISABLE_CHAT&>";
        public const string DISABLE_BROADCAST = "<vp=DISABLE_BROADCAST&>";
        public const string EMBDDED = "<e=EMBEDDED&>";
        public const string FULLSCREEN = "<fs=FULLSCREEN&>";
        public const string PERFSTATS = "<showpagestats=PERFSTATS&>";
        public const string RECORDING = "<rec=RECORDING&>";
        public const string THEME_ID = "<thm=THEME_ID&>";
        public const string UI_LLCC = "<ui=UI_LLCC&>";
        public const string VALIDATOR_TEST_CATEGORY = "<testcategory=VALIDATOR_TEST_CATEGORY>";

        /// <summary>
        /// Sets a specific WOPI URL placeholder with the correct value
        /// Most of these are hard-coded in this WOPI implementation
        /// </summary>
        public static string GetPlaceholderValue(string placeholder)
        {
            var ph = placeholder.Substring(1, placeholder.IndexOf("="));
            string result = "";
            switch (placeholder)
            {
                case BUSINESS_USER:
                    result = ph + "1";
                    break;
                case DC_LLCC:
                case UI_LLCC:
                    result = ph + "1033";
                    break;
                case DISABLE_ASYNC:
                case DISABLE_BROADCAST:
                case EMBDDED:
                case FULLSCREEN:
                case RECORDING:
                case THEME_ID:
                    // These are all broadcast related actions
                    result = ph + "true";
                    break;
                case DISABLE_CHAT:
                    result = ph + "false";
                    break;
                case PERFSTATS:
                    result = ""; // No documentation
                    break;
                case VALIDATOR_TEST_CATEGORY:
                    result = ph + "OfficeOnline"; //This value can be set to All, OfficeOnline or OfficeNativeClient to activate tests specific to Office Online and Office for iOS. If omitted, the default value is All.
                    break;
                default:
                    result = "";
                    break;

            }

            return result;
        }
    }
}
