using com.microsoft.dx.officewopi.Models;
using com.microsoft.dx.officewopi.Models.Wopi;
using com.microsoft.dx.officewopi.Security;
using com.microsoft.dx.officewopi.Utils;
using Microsoft.Owin.Security.OAuth;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Caching;
using System.Threading.Tasks;
using System.Web;
using System.Web.Mvc;

namespace com.microsoft.dx.officewopi.Controllers
{
    [Authorize]
    public class HomeController : Controller
    {
        /// <summary>
        /// Index view displays all files for the signed in user
        /// </summary>
        [Authorize]
        public async Task<ActionResult> Index()
        {
            // Get files for the user
            var files = DocumentDBRepository<DetailedFileModel>.GetItems("Files", i => i.OwnerId == User.Identity.Name.ToLower()).ToList();

            // Populate valid actions for each of the files
            await files.PopulateActions();

            // Return the view with the files
            return View(files);
        }

        /// <summary>
        /// Detail view hosts the WOPI host frame and loads the appropriate action view from Office Online
        /// </summary>
        [Authorize]
        [Route("Home/Detail/{id}")]
        public async Task<ActionResult> Detail(Guid id)
        {
            // Make sure an action was passed in
            if (String.IsNullOrEmpty(Request["action"]))
                return RedirectToAction("Error", "Home", new { error = "No action provided" });

            // Get the specific file from DocumentDB
            var file = DocumentDBRepository<FileModel>.GetItem("Files",
                i => i.OwnerId == User.Identity.Name.ToLower() && i.id == id);

            // Check for null file
            if (file == null)
                return RedirectToAction("Error", "Home", new { error = "Files does not exist" });

            // Use discovery to determine endpoint to leverage
            List<WopiAction> discoData = await WopiUtil.GetDiscoveryInfo();
            var fileExt = file.BaseFileName.Substring(file.BaseFileName.LastIndexOf('.') + 1).ToLower();
            var action = discoData.FirstOrDefault(i => i.name == Request["action"] && i.ext == fileExt);

            // Make sure the action isn't null
            if (action != null)
            {
                string urlsrc = WopiUtil.GetActionUrl(action, file, Request.Url.Authority);

                // Generate JWT token for the user/document
                WopiSecurity wopiSecurity = new WopiSecurity();
                var token = wopiSecurity.GenerateToken(User.Identity.Name.ToLower(), getUserContainer(), id.ToString());
                ViewData["access_token"] = wopiSecurity.WriteToken(token);
                ViewData["access_token_ttl"] = token.ValidTo.Subtract(new DateTime(1970, 1, 1)).TotalMilliseconds;
                ViewData["wopi_urlsrc"] = urlsrc;
                return View();
            }
            else
            {
                // This will only hit if the extension isn't supported by WOPI
                return RedirectToAction("Error", "Home", new { error = "File is not a supported WOPI extension" });
            }
        }

        /// <summary>
        /// Adds the submitted files for Azure Blob Storage and metadata into DocumentDB
        /// </summary>
        [HttpPost]
        [Authorize]
        public async Task<ActionResult> Add()
        {
            try
            {
                // Create the file entity
                DetailedFileModel file = new DetailedFileModel()
                {
                    id = Guid.NewGuid(),
                    OwnerId = User.Identity.Name.ToLower(),
                    BaseFileName = HttpUtility.UrlDecode(Request["HTTP_X_FILE_NAME"]),
                    Size = Convert.ToInt32(Request["HTTP_X_FILE_SIZE"]),
                    Container = getUserContainer(),
                    Version = 1
                };

                // Populate valid actions for each of the files
                await file.PopulateActions();

                // First stream the file into blob storage
                var stream = Request.InputStream;
                var bytes = new byte[stream.Length];
                await stream.ReadAsync(bytes, 0, (int)stream.Length);
                var id = await Utils.AzureStorageUtil.UploadFile(file.id.ToString(), file.Container, bytes);

                // Write the details into documentDB
                await DocumentDBRepository<FileModel>.CreateItemAsync("Files", (FileModel)file);

                // Return json representation of information
                return Json(new { success = true, file = file });
            }
            catch (Exception)
            {
                // Something failed...return false
                return Json(new { success = false });
            }
        }

        /// <summary>
        /// Deletes the file from Azure Blob Storage and metadata into DocumentDB
        /// </summary>
        [HttpDelete]
        [Authorize]
        [Route("Home/Delete/{id}")]
        public async Task<ActionResult> Delete(Guid id)
        {
            try
            {
                // Get the file from DocumentDB
                var file = DocumentDBRepository<FileModel>.GetItem("Files", i => i.id == id);

                // Delete the record from DocumentDB
                await DocumentDBRepository<FileModel>.DeleteItemAsync("Files", file.id.ToString(), file);

                // Delete the blob
                await Utils.AzureStorageUtil.DeleteFile(file.id.ToString(), file.Container);

                //return json representation of information
                return Json(new { success = true, id = id.ToString() });
            }
            catch (Exception)
            {
                // Something failed...return false
                return Json(new { success = false, id = id.ToString() });
            }
        }

        /// <summary>
        /// Error view displays error messages passed from other controllers
        /// </summary>
        public ActionResult Error(string error)
        {
            ViewData["Error"] = error;
            return View();
        }

        /// <summary>
        /// Simple function to correctly form a container name based on the signed in user
        /// </summary>
        private string getUserContainer()
        {
            return User.Identity.Name.Replace("@", "-").Replace(".", "-");
        }
    }
}