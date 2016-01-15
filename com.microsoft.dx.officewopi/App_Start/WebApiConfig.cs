using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web.Http;

namespace com.microsoft.dx.officewopi
{
    public static class WebApiConfig
    {
        public static void Register(HttpConfiguration config)
        {
            // Ignore AAD Auth for WebAPI...will be handled by WopiTokenValidationFilter class
            config.SuppressDefaultHostAuthentication();

            // Web API routes
            config.MapHttpAttributeRoutes();

            config.Routes.MapHttpRoute(
                name: "WopiDefault",
                routeTemplate: "wopi/{controller}/{id}/{action}",
                defaults: new { id = RouteParameter.Optional, action = RouteParameter.Optional }
            );
        }
    }
}
