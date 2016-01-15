using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace com.microsoft.dx.officewopi.Models.Wopi
{
    /// <summary>
    /// Represents the base information determined at the beginning of a WOPI request
    /// </summary>
    public class WopiRequest
    {
        public string Id { get; set; }
        public WopiRequestType RequestType { get; set; }
        public string AccessToken { get; set; }
    }
}
