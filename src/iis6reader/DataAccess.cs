namespace maestropanel.iis6reader
{
    using System;
    using System.Management;

    public class DataAccess
    {

        public ManagementObjectCollection GetProperties(string query)
        {
            using (var s = GetSearcher(query))
                return s.Get();
        }

        private ManagementObjectSearcher GetSearcher(string query)
        {
            return new ManagementObjectSearcher("root\\MicrosoftIISv2", query);
        }

        public T GetValue<T>(ManagementObject source, string name)
        {
            var def = default(T);

            if (String.IsNullOrEmpty(name))
                return def;

            if (source == null)
                return def;

            if (!isExists(source, name))
                return def;

            def = (T)source[name];

            return def;
        }

        public T GetValue<T>(ManagementBaseObject source, string name)
        {
            var def = default(T);

            if (String.IsNullOrEmpty(name))
                return def;

            if (source == null)
                return def;

            if (!isExists(source, name))
                return def;

            def = (T)source[name];

            return def;
        }

        private bool isExists(ManagementObject source, string name)
        {
            var exists = false;

            foreach (var item in source.Properties)
            {
                if (item.Name == name)
                {
                    exists = true;
                    break;
                }
            }

            return exists;
        }

        private bool isExists(ManagementBaseObject source, string name)
        {
            var exists = false;

            foreach (var item in source.Properties)
            {
                if (item.Name == name)
                {
                    exists = true;
                    break;
                }
            }

            return exists;
        }
    }
}
