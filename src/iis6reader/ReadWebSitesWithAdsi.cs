﻿namespace maestropanel.iis6reader
{
    using maestropanel.iis6reader.Models;
    using Microsoft.Win32;
    using System;
    using System.Collections.Generic;
    using System.DirectoryServices;

    internal class ReadWebSitesWithAdsi : IReadWebSite
    {

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

            if (source == null)
            {
                return exists;
            }

            if(String.IsNullOrEmpty(name)) {
                return exists;
            }

            try
            {
                foreach (var item in source.PropertyNames)
                {
                    if (item.ToString() == name)
                    {
                        exists = true;
                        break;
                    }
                }

            }
            catch
            {}

            return exists;
        }

        private string GetDomainPath(string metaname)
        {
            var path = String.Empty;

            using (DirectoryEntry w3svc = new DirectoryEntry(String.Format("IIS://localhost/W3SVC/{0}", metaname)))
            {
                foreach (DirectoryEntry de in w3svc.Children)
                {
                    if (de.Name == "Path")
                    {
                        path = GetValue<string>(de.Properties, "Path");
                        break;
                    }                    
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
            {
                var itemstr = item.ToString();

                if (itemstr.Length == 1)
                    itemstr = String.Format("0{0}", itemstr);

                certHash += itemstr.ToString();
            }
            
            return certHash;
        }
        
        public List<WebSite> GetAllDomains(string where = "")
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

                        if (!String.IsNullOrEmpty(where))
                        {
                            if (d.Name != where)
                                continue;
                        }
                                                
                        d.MetaName = de.Name;
                        d.UserName = GetValue<string>(de.Properties, "AnonymousUserName");
                        d.Password = GetValue<string>(de.Properties, "AnonymousUserPass");
                        d.Path = GetDomainPath(d.MetaName);
                        d.VirtualDirectories = new List<WebSiteVirtualDirectory>().ToArray();

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

                        var rootApp = GetRoot(de.Children);

                        if (rootApp != null)
                        {
                            var vlist = GetVirtualDirectories("", rootApp.Children, "");

                            if (vlist.Count > 0)
                                d.VirtualDirectories = vlist.ToArray();

                            var items = GetWebSiteItems(rootApp);

                            d.ApplicationPoolName = items.ApplicationPoolId;
                            d.DefaultDocs = items.DefaultDocs;
                            d.DotNetRuntime = items.DotNetRuntime;

                            d.HttpErrors = GetWebSiteCustomErrors(rootApp.Properties["HttpErrors"]);
                        }
                        else
                        {
                            d.DefaultDocs = new List<string>().ToArray();
                            d.HttpErrors = GetWebSiteCustomErrors(de.Properties["HttpErrors"]);
                        }

                        d.Bindings = bindings.ToArray();
                        d.Headers = GetWebSiteCustomHeader(de.Properties["HttpCustomHeaders"]);

                        d.SSLCertHash = GetCertificateHash(de.Properties["SSLCertHash"]);
                        d.State = serverState.ToString();                        


                        list.Add(d);
                    }
                }
            }

            return list;
        }

        public List<WebSite> GetAllDomains()
        {
            return GetAllDomains(where: "");
        }

        private DirectoryEntry GetRoot(DirectoryEntries children)
        {
            DirectoryEntry _root = null;

            foreach (DirectoryEntry item in children)
            {

                if (item.Name == "ROOT" && item.SchemaClassName == "IIsWebVirtualDir")
                {
                    _root = item;
                    break;
                }
            }

            return _root;
        }

        private WebSiteItems GetWebSiteItems(DirectoryEntry children)
        {
            var witem = new WebSiteItems();

            witem.ApplicationPoolId = GetValue<string>(children.Properties, "AppPoolId");
            witem.DotNetRuntime = GetDotNetVersion(children);
            witem.DefaultDocs = GetDefaultDocs(children);

            return witem;
        }

        private List<WebSiteVirtualDirectory> GetVirtualDirectories(string parentPath, DirectoryEntries children, string parentKey)
        {
            var list = new List<WebSiteVirtualDirectory>();

            foreach (DirectoryEntry item in children)
            {
                if (item.SchemaClassName != "IIsWebVirtualDir")
                    continue;

                if (item.Name.StartsWith("_vti"))
                    continue;

                if (item.Name.Equals("aspnet_client"))
                    continue;

                if (item.Name.Equals("_private"))
                    continue;

                var virtualPath = String.Format("{0}/{1}", parentPath, item.Name);
                                
                var l = new WebSiteVirtualDirectory();
                l.Name = virtualPath;
                l.Path = GetValue<string>(item.Properties, "Path");
                l.isApplication = (GetValue<uint>(item.Properties, "AccessFlasg") == 513);

                list.Add(l);

                list.AddRange(GetVirtualDirectories(virtualPath, item.Children, item.Name));
            }

            return list;
        }

        private string[] GetDefaultDocs(DirectoryEntry children)
        {
            var list = new List<string>();
            var defaultDocsTemplate = DefaultDocs();

            var default_docs = children.Properties["DefaultDoc"];

            if (default_docs == null)
                return list.ToArray();
            
            foreach (var item in default_docs)
            {
                var docs = item.ToString();
                if (!String.IsNullOrEmpty(docs))
                {
                    list.AddRange(docs.Split(','));
                    break;
                }
            }

            return DefaultDocsDiff(list).ToArray();
        }

        private List<string> DefaultDocsDiff(List<string> list2)
        {
            var diffList = new List<string>();
            var list1 = DefaultDocs();

            foreach (var item in list2)
            {
                var litem = item.ToLower();
                var isExists = list1.Contains(litem);
                if (!isExists)
                {
                    diffList.Add(litem);
                }
            }

            return diffList;
        }

        private List<string> DefaultDocs()
        {
            var l = new List<string>();
            l.Add("default.htm");
            l.Add("default.html");
            l.Add("default.asp");
            l.Add("default.aspx");
            l.Add("default.php");
            l.Add("index.htm");
            l.Add("index.html");
            l.Add("index.asp");
            l.Add("index.aspx");
            l.Add("index.php");

            return l;
        }

        private string GetDotNetVersion(DirectoryEntry children)
        {
            var runtime = "unknown";            

            var scriptMapes = children.Properties["ScriptMaps"];

            if (scriptMapes == null)
                return runtime;

            foreach (var item in scriptMapes)
            {
                var scriptItem = item.ToString();

                //.aspx,c:\windows\microsoft.net\framework\v4.0.30319\aspnet_isapi.dll,1,GET,HEAD,POST,DEBUG
                if (scriptItem.StartsWith(".aspx"))
                {
                    
                    if (scriptItem.IndexOf("v1.0") != -1)
                    {
                        runtime = "v1.0";
                        break;
                    }

                    if (scriptItem.IndexOf("v1.1") != -1)
                    {
                        runtime = "v1.1";
                        break;
                    }

                    if (scriptItem.IndexOf("v2.0") != -1)
                    {
                        runtime = "v2.0";
                        break;
                    }

                    if (scriptItem.IndexOf("v4.0") != -1)
                    {
                        runtime = "v4.0";
                        break;
                    }

                    break;
                }
            }

            return runtime;
        }


    }

    struct WebSiteItems
    {
        public string DotNetRuntime { get; set; }
        public string ApplicationPoolId { get; set; }
        public string[] DefaultDocs { get; set; }
    }
}
