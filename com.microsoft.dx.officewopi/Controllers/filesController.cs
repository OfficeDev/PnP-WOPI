using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web;
using System.Web.Http;
using com.microsoft.dx.officewopi.Utils;
using System.Threading.Tasks;
using com.microsoft.dx.officewopi.Security;

namespace com.microsoft.dx.officewopi.Controllers
{
    [WopiTokenValidationFilter]
    public class filesController : ApiController
    {
        [WopiTokenValidationFilter]
        [HttpGet]
        [Route("wopi/files/{id}")]
        public async Task<HttpResponseMessage> Get(Guid id)
        {
            //Handles CheckFileInfo
            return await HttpContext.Current.ProcessWopiRequest();
        }

        [WopiTokenValidationFilter]
        [HttpGet]
        [Route("wopi/files/{id}/contents")]
        public async Task<HttpResponseMessage> Contents(Guid id)
        {
            //Handles GetFile
            return await HttpContext.Current.ProcessWopiRequest();
        }

        [WopiTokenValidationFilter]
        [HttpPost]
        [Route("wopi/files/{id}")]
        public async Task<HttpResponseMessage> Post(Guid id)
        {
            //Handles Lock, GetLock, RefreshLock, Unlock, UnlockAndRelock, PutRelativeFile, RenameFile, PutUserInfo
            return await HttpContext.Current.ProcessWopiRequest();
        }

        [WopiTokenValidationFilter]
        [HttpPost]
        [Route("wopi/files/{id}/contents")]
        public async Task<HttpResponseMessage> PostContents(Guid id)
        {
            //Handles PutFile
            return await HttpContext.Current.ProcessWopiRequest();
        }
    }
}
