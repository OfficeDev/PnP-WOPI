# PnP-WOPI
This repository contains an application that integrates with Office Online for viewing/editing Office documents. This type of integration classifies this application as a WOPI host. WOPI (Web Application Open Platform Interface) is a protocol for integrating with Office Online and is documented in detail at [https://wopi.readthedocs.org](https://wopi.readthedocs.org "https://wopi.readthedocs.org"). This sample illustrates many important patterns and practices for implementing a WOPI host, a number of which are outlined in this readme. 

This WOPI host implementation is deployed to [https://officewopi.azurewebsites.net](https://officewopi.azurewebsites.net "https://officewopi.azurewebsites.net") and can be used/tested by anyone with an organization/student account registered with Microsoft (read: Office 365 logins). It is provided for testing/experimenting purposes and offered with no service level agreement.

> NOTE: You cannot simply clone and run this sample locally. Integrating with Office Online requires that your host domain is white-listed by Microsoft. The first step to white-listing a domain is to join the Cloud Storage Provider Program detail [HERE](http://dev.office.com/programs/officecloudstorage "HERE"). Additionally, a WOPI host must expose endpoints to the internet that Office Online can communicate with (read: localhost probably won't work).

## Solution Overview ##
This WOPI host sample is written in ASP.NET/C# with MVC for the user interface and Web API for the WOPI endpoints. Although it uses Azure AD for user identity, Azure AD has NOTHING to do with the WOPI integration. A WOPI host can use virtually any identity provider (or function anonymously). The sample stores files in Azure Blob Storage and file metadata in Azure DocumentDB (a NoSQL platform service similar to MongoDB). There a number of configuration values that should be updated in the web.config to support the identity and storage providers:

    <!-- These are Azure AD specific properties-->
    <add key="ida:ClientId" value="CLIENT ID FROM AZURE AD" />
    <add key="ida:ClientSecret" value="CLIENT SECRET FROM AZURE AD" />
    
    <!-- This is the private key to the self-signed cert...probably shouldn't be in config file -->
    <add key="CertPassword" value="CERT PRIVATE KEY/PASSWORD"/>
    
    <!-- These are properties for Azure Blob Storage -->
    <add key="abs:Protocol" value="AZURE BLOB STORAGE PROTOCOL...LIKELY https" />
    <add key="abs:AccountName" value="AZURE BLOB STORAGE ACCOUNT NAME" />
    <add key="abs:AccountKey" value="AZURE BLOB STORAGE ACCOUNT KEY" />
    
    <!-- These are properties for DocumentDB -->
    <add key="ddb:endpoint" value="DOCUMENTDB ENDPOINT" />
    <add key="ddb:authKey" value="DOCUMENTDB AUTH KEY" />
    <add key="ddb:database" value="DOCUMENTDB DATABASE NAME" />

A WOPI host is composed of two primary components...a frame to host the Office Online renderings and endpoints that Office Online can call into to perform specific operations (ex: GetFile, PutFile, etc). Both of these components and their unique considerations are detailed below.

## WOPI Host Page ##
The WOPI host page for this sample is found in the **Home/Detail** view with logic in the **Detail** method of the **HomeController.cs**. This page can only be invoked with a WOPI action and a file id. The WOPI action includes details on how to reach Office Online for the desired action (ex: view, edit, etc). The file id is used to look up file details which placed in a form in the Detail view that is posted to the WOPI action URL. When invoked, the controller will also generate a user and file specific access token that is part of the POST to the WOPI action URL.


    <form id="office_form" name="office_form" target="office_frame" action='@ViewData["wopi_urlsrc"]' method="post">
        <input name="access_token" value='@ViewData["access_token"]' type="hidden" />
        <input name="access_token_ttl" value='@ViewData["access_token_ttl"]' type="hidden" />
    </form>
    <span id="frameholder"></span>
    <script type="text/javascript">
        var frameholder = document.getElementById("frameholder");
        var office_frame = document.createElement("iframe");
        office_frame.name = "office_frame";
        office_frame.id ="office_frame";
        frameholder.appendChild(office_frame);

		//Submit the form the WOPI action URL
        document.getElementById("office_form").submit();
    </script>

Essentially, the WOPI Host Page hosts and posts data to a big IFRAME that renders Office Online.

## WOPI Endpoints ##
The WOPI endpoints should not use the default auth that is configured in Startup.Auth.cs. Remember, Office Online is what calls into these endpoints and it has no dependency on Azure AD. Office Online will pass an access token in the header of all WOPI requests (using the Authorization header). This is the exact same access token that the WOPI Host Page generated a posted to the WOPI action URL. To accomplish this from the same web application, the **WebApiConfig.cs** needs to ignore the default authentication:

	// Ignore AAD Auth for WebAPI...will be handled by WopiTokenValidationFilter class
	config.SuppressDefaultHostAuthentication();

The application also needs a **AuthorizeAttribute** to validate the access token on requests. This sample implements this in the **WopiTokenValidationFilter** class. The WebAPI routes are all configured with this filter as seen below. The **WopiSecurity** class contains methods to generate and validate our custom access tokens.

	[WopiTokenValidationFilter]
    [HttpGet]
    [Route("wopi/files/{id}")]
    public async Task<HttpResponseMessage> Get(Guid id)
    {
        //Handles CheckFileInfo
        return await HttpContext.Current.ProcessWopiRequest();
    }

One of the challenges of implementing the WOPI endpoints with WebAPI is that most of the WOPI operations use the same few routes. Operations are instead determined by the header details included on requests. As such, the **filesController** has four generic endpoints that simply call a **ProcessWopiRequest** extension on the HttpContext:

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

Most of the WOPI logic exists in the **WOPIExtensions.cs** and **WOPIUtils.cs** files. The **WOPIExtensions.cs** file contains extension methods for each WOPI operation and the **WOPIUtils.cs** contains utility methods for doing things such as WOPI discovery (which lists all the actions and proof keys for the WOPI integration), validating WOPI proof (ie - proving that the WOPI request actually came from Office Online), and a number of important WOPI constants (like the numerous custom headers WOPI uses).