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
    public class foldersController : ApiController
    {
        [WopiTokenValidationFilter]
        [HttpGet]
        [Route("wopi/folders/{id}")]
        public async Task<HttpResponseMessage> Get(Guid id)
        {
            return await HttpContext.Current.ProcessWopiRequest();
        }

        [WopiTokenValidationFilter]
        [HttpGet]
        [Route("wopi/folders/{id}/contents")]
        public async Task<HttpResponseMessage> Contents(Guid id)
        {
            return await HttpContext.Current.ProcessWopiRequest();
        }

        [WopiTokenValidationFilter]
        [HttpPost]
        [Route("wopi/folders/{id}")]
        public async Task<HttpResponseMessage> Post(Guid id)
        {
            return await HttpContext.Current.ProcessWopiRequest();
        }

        [WopiTokenValidationFilter]
        [HttpPost]
        [Route("wopi/folders/{id}/contents")]
        public async Task<HttpResponseMessage> PostContents(Guid id)
        {
            return await HttpContext.Current.ProcessWopiRequest();
        }
    }
}
