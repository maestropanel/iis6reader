﻿namespace maestropanel.iis6reader
{
    using maestropanel.iis6reader.Models;
    using System;
    using System.Collections.Generic;
    using System.DirectoryServices;

    public class ReadWebSitesWithAdsi
    {
        public List<WebSite> GetAllDomains()
        {
            var list = new List<WebSite>();

            using (DirectoryEntry w3svc = new DirectoryEntry("IIS://localhost/W3SVC"))
            {
                foreach (DirectoryEntry de in w3svc.Children)
                {
                    if (de.SchemaClassName == "IIsWebServer")
                    {                        
                        var d = new WebSite();
                        d.Name = GetValue<string>(de.Properties, "ServerComment");
                        d.MetaName = de.Name;
                        d.UserName = GetValue<string>(de.Properties, "AnonymousUserName");
                        d.Password = GetValue<string>(de.Properties, "AnonymousUserPass");                        
                        d.Path = GetDomainPath(d.MetaName);

                        var serverState = GetValue<int>(de.Properties, "ServerState");
                        var serverBindings = GetWebSiteBindings(de.Properties["ServerBindings"]);
                        var secureBindings = GetWebSiteBindings(de.Properties["SecureBindings"], secureBinding: true);

                        var bindings = new List<WebSiteBinding>();
                        
                        if (serverBindings.Length > 0)
                            bindings.AddRange(serverBindings);

                        if (secureBindings.Length > 0)
                        {
                            bindings.AddRange(secureBindings);
                            d.EnableSSL = true;
                        }

                        d.Headers = GetWebSiteCustomHeader(de.Properties["HttpCustomHeaders"]);
                        d.HttpErrors = GetWebSiteCustomErrors(de.Properties["HttpErrors"]);
                        d.SSLCertHash = GetCertificateHash(de.Properties["SSLCertHash"]);
                        d.State = serverState.ToString();
                        
                        list.Add(d);
                    }
                }
            }

            return list;
        }

        private T GetValue<T>(PropertyCollection source, string name)
        {
            var def = default(T);

            if (String.IsNullOrEmpty(name))
                return def;

            if (source == null)
                return def;

            if (!isExists(source, name))
                return def;

            if (source[name].Count == 0)
                return def;

            def = (T)source[name][0];

            return def;
        }

        private bool isExists(PropertyCollection source, string name)
        {
            var exists = false;

            foreach (var item in source.PropertyNames)
            {
                if (item.ToString() == name)
                {
                    exists = true;
                    break;
                }
            }

            return exists;
        }

        private string GetDomainPath(string metaname)
        {
            var path = String.Empty;

            using (DirectoryEntry w3svc = new DirectoryEntry(String.Format("IIS://localhost/W3SVC/{0}", metaname)))
            {
                foreach (DirectoryEntry de in w3svc.Children)
                {                    
                    path = GetValue<string>(de.Properties, "Path");

                    if (de.Name == "Path")
                        break;
                }
            }

            return path;
        }

        private WebSiteBinding[] GetWebSiteBindings(PropertyValueCollection bindings, bool secureBinding = false)
        {
            var list = new List<WebSiteBinding>();

            if (bindings == null)
                return list.ToArray();

            foreach (var item in bindings)
            {                
                var itemStr = item.ToString();
                var itemArray = itemStr.Split(':');

                if (itemArray.Length == 3)
                {
                    var b = new WebSiteBinding();
                    b.IpAddr = itemArray[0];
                    b.Port = itemArray[1];
                    b.Hostname = itemArray[2];
                    b.isSecure = secureBinding;

                    list.Add(b);
                }

            }

            return list.ToArray();
        }

        private WebSiteCustomHeader[] GetWebSiteCustomHeader(PropertyValueCollection headers)
        {

            var list = new List<WebSiteCustomHeader>();

            if (headers == null)
                return list.ToArray();

            foreach (var item in headers)
            {
                var keyvalue = item.ToString().Split(':');
                if (keyvalue.Length == 2)
                {
                    var ch = new WebSiteCustomHeader();
                    ch.Name = keyvalue[0].Trim();
                    ch.Value = keyvalue[1].Trim();

                    list.Add(ch);
                }                
            }

            return list.ToArray();

        }

        private WebSiteCustomError[] GetWebSiteCustomErrors(PropertyValueCollection errors)
        {
            var list = new List<WebSiteCustomError>();

            if (errors == null)
                return list.ToArray();

            //404,*,URL,/404.asp
            foreach (var item in errors)
            {
                var customPage = item.ToString().Split(',');
                if (customPage.Length == 4)
                {
                    var err = new WebSiteCustomError();
                    err.HandlerLocation = customPage[3];
                    err.HandlerType = customPage[2];
                    err.HttpErrorCode = customPage[0];

                    if (err.HandlerType == "URL")
                    {
                        err.HandlerType = "ExecuteURL";
                        list.Add(err);
                    }                    
                }
            }

            return list.ToArray();
        }

        private string GetCertificateHash(PropertyValueCollection property)
        {
            if (property == null)
                return String.Empty;

            if (property.Value == null)
                return String.Empty;

            var values = (Object[])property.Value;
            var certHash = String.Empty;

            foreach (var item in values)            
                certHash += item.ToString();
            
            return certHash;
        }
    }    
}
