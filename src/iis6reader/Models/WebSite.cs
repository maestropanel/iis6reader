namespace maestropanel.iis6reader.Models
{
    public class WebSite
    {
        public string Name { get; set; }
        public string MetaName { get; set; }
        public string State { get; set; }
        public bool EnableSSL { get; set; }
        public bool EnableDirBrowsing { get; set; }
        public string Path { get; set; }

        public WebSiteCustomHeader[] Headers { get; set; }
        public WebSiteMimeType[] MimeTypes { get; set; }
        public WebSiteCustomError[] HttpErrors { get; set; }
        public WebSiteBinding[] Bindings { get; set; }
        public WebSiteVirtualDirectory[] VirtualDirectories { get; set; }

        public string UserName { get; set; }
        public string Password { get; set; }
        public string SSLCertHash { get; set; }
    }

    public class WebSiteBinding
    {
        public bool isSecure { get; set; } //SecureBindings
        public string Hostname { get; set; }
        public string Port { get; set; }
        public string IpAddr { get; set; }
    }

    public class WebSiteCustomHeader
    {
        public string Name { get; set; }
        public string Value { get; set; }
    }
    
    public class WebSiteMimeType
    {
        public string Extension { get; set; }
        public string MType { get; set; }
    }

    public class WebSiteCustomError
    {
        public string HandlerLocation { get; set; }
        public string HandlerType { get; set; }
        public string HttpErrorCode { get; set; }
        public string HttpErrorSubcode { get; set; }
    }

    public class WebSiteVirtualDirectory
    {        
        public string Name { get; set; }
        public string Path { get; set; }
        public bool isApplication { get; set; }
    }
}
