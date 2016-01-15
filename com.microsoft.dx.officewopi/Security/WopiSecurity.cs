using com.microsoft.dx.officewopi.Utils;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IdentityModel.Protocols.WSTrust;
using System.IdentityModel.Tokens;
using System.IO;
using System.Linq;
using System.Security.Claims;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading.Tasks;

namespace com.microsoft.dx.officewopi.Security
{
    /// <summary>
    /// Class handles token generation and validation for the WOPI host
    /// </summary>
    public class WopiSecurity
    {
        /// <summary>
        /// Generates an access token specific to a user and file
        /// </summary>
        public static bool ValidateToken(string tokenString, string container, string docId)
        {
            // Initialize the token handler and validation parameters
            var tokenHandler = new JwtSecurityTokenHandler();
            var tokenValidation = new TokenValidationParameters()
            {
                ValidAudience = SettingsHelper.Audience,
                ValidIssuer = SettingsHelper.Audience,
                IssuerSigningToken = new X509SecurityToken(getCert())
            };

            try
            {
                // Try to validate the token
                SecurityToken token = null;
                var principal = tokenHandler.ValidateToken(tokenString, tokenValidation, out token);
                return (principal.HasClaim("container", container) &&
                    principal.HasClaim("docid", docId));
            }
            catch (Exception)
            {
                return false;
            }
        }

        /// <summary>
        /// Extracts the user information from a provided access token
        /// </summary>
        public static string GetUserFromToken(string tokenString)
        {
            // Initialize the token handler and validation parameters
            var tokenHandler = new JwtSecurityTokenHandler();
            var tokenValidation = new TokenValidationParameters()
            {
                ValidAudience = SettingsHelper.Audience,
                ValidIssuer = SettingsHelper.Audience,
                IssuerSigningToken = new X509SecurityToken(getCert())
            };

            try
            {
                // Try to extract the user principal from the token
                SecurityToken token = null;
                var principal = tokenHandler.ValidateToken(tokenString, tokenValidation, out token);
                return principal.Identity.Name;
            }
            catch (Exception)
            {
                return String.Empty;
            }
        }

        /// <summary>
        /// Private token handler used in instance classes
        /// </summary>
        private JwtSecurityTokenHandler tokenHandler;

        /// <summary>
        /// Generates an access token for the user and file
        /// </summary>
        public JwtSecurityToken GenerateToken(string user, string container, string docId)
        {
            var now = DateTime.UtcNow;
            tokenHandler = new JwtSecurityTokenHandler();
            var signingCert = getCert();
            var tokenDescriptor = new SecurityTokenDescriptor
            {
                Subject = new ClaimsIdentity(new[]
                    {
                        new Claim(ClaimTypes.Name, user),
                        new Claim("container", container),
                        new Claim("docid", docId)
                }),
                TokenIssuerName = SettingsHelper.Audience,
                AppliesToAddress = SettingsHelper.Audience,
                Lifetime = new Lifetime(now, now.AddHours(1)),
                EncryptingCredentials = new X509EncryptingCredentials(signingCert),
                SigningCredentials = new X509SigningCredentials(signingCert)
            };
            var token = (JwtSecurityToken)tokenHandler.CreateToken(tokenDescriptor);
            return token;
        }

        /// <summary>
        /// Converts the JwtSecurityToken to a Base64 string that can be used by the Host
        /// </summary>
        public string WriteToken(JwtSecurityToken token)
        {
            return tokenHandler.WriteToken(token);
        }

        /// <summary>
        /// Gets the self-signed certificate used to sign the access tokens
        /// </summary>
        private static X509Certificate2 getCert()
        {
            var certPath = Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~"), "OfficeWopi.pfx");
            var certfile = System.IO.File.OpenRead(certPath);
            var certificateBytes = new byte[certfile.Length];
            certfile.Read(certificateBytes, 0, (int)certfile.Length);
            var cert = new X509Certificate2(
                certificateBytes,
                ConfigurationManager.AppSettings["CertPassword"],
                X509KeyStorageFlags.Exportable |
                X509KeyStorageFlags.MachineKeySet |
                X509KeyStorageFlags.PersistKeySet);

            return cert;
        }
    }
}
