namespace maestropanel.iis6reader
{
    using maestropanel.iis6reader.Models;
    using Microsoft.Win32;
    using System.Collections.Generic;

    public class ReadWebSites
    {
        private string _provider;

        /// <summary>
        /// 
        /// </summary>
        /// <param name="provider">wmi or adsi. Default wmi</param>
        public ReadWebSites(string provider)
        {
            if (string.IsNullOrEmpty(provider))
                provider = "wmi";

            this._provider = provider;

        }

        public ReadWebSites()
        {
            this._provider = "wmi";
        }
        
        public List<WebSite> GetAllDomains(string where = "")
        {
            IReadWebSite provider = CreateObject();

            if (string.IsNullOrEmpty(where))
                return provider.GetAllDomains();
            else
                return provider.GetAllDomains(where);
        }

        private IReadWebSite CreateObject()
        {
            if (this._provider == "wmi")
            {
                return new ReadWebSitesWithWMI();
            }
            else
            {
                return new ReadWebSitesWithAdsi();
            }
        }

        public bool IsInstalled()
        {
            return Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\InetStp", false) == null ? false : true;
        }
    }
}
