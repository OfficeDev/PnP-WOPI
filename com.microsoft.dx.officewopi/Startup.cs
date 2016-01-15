using System;
using System.Threading.Tasks;
using Microsoft.Owin;
using Owin;

namespace com.microsoft.dx.officewopi
{
    public partial class Startup
    {
        public void Configuration(IAppBuilder app)
        {
            ConfigureAuth(app);
        }
    }
}
