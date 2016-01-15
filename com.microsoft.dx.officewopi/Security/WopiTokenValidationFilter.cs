using com.microsoft.dx.officewopi.Models;
using com.microsoft.dx.officewopi.Models.Wopi;
using com.microsoft.dx.officewopi.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web.Http;
using System.Web.Http.Controllers;
using System.Web.Http.Filters;

namespace com.microsoft.dx.officewopi.Security
{
    public class WopiTokenValidationFilter : AuthorizeAttribute
    {
        /// <summary>
        /// Determines if the user is authorized to access the WebAPI endpoint based on the bearer token
        /// </summary>
        protected override bool IsAuthorized(HttpActionContext actionContext)
        {
            try
            {
                // Parse the query string and ensure there is an access_token
                var queryParams = parseRequestParams(actionContext.Request.RequestUri.Query);
                if (!queryParams.ContainsKey("access_token"))
                    return false;

                // Get the details of the WOPI request
                WopiRequest requestData = new WopiRequest()
                {
                    RequestType = WopiRequestType.None,
                    AccessToken = queryParams["access_token"],
                    Id = actionContext.RequestContext.RouteData.Values["id"].ToString()
                };
                
                // Get the requested file from Document DB
                var itemId = new Guid(requestData.Id);
                var file = DocumentDBRepository<DetailedFileModel>.GetItem("Files", i => i.id == itemId);

                // Check for missing file
                if (file == null)
                    return false;

                // Validate the access token
                return WopiSecurity.ValidateToken(requestData.AccessToken, file.Container, file.id.ToString());
            }
            catch (Exception)
            {
                // Any exception will return false, but should probably return an alternate status codes
                return false;
            }
        }

        /// <summary>
        /// Simple method to parse a URL query string into a dictionary of parameters
        /// </summary>
        private Dictionary<string, string> parseRequestParams(string query)
        {
            Dictionary<string, string> queryParams = new Dictionary<string, string>();
            query = query.Substring(1);
            var parts = query.Split('&');
            foreach (var part in parts)
            {
                var parts2 = part.Split('=');
                queryParams.Add(parts2[0], Uri.UnescapeDataString(parts2[1]));
            }
            return queryParams;
        }
    }
}
