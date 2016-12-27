namespace maestropanel.iis6reader
{
    using maestropanel.iis6reader.Models;
    using System.Collections.Generic;

    internal interface IReadWebSite
    {
        List<WebSite> GetAllDomains();
        List<WebSite> GetAllDomains(string where = "");
    }
}
