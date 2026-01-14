using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Win32;
using System.Diagnostics;

namespace BDOpenOffice
{
    static class RegistryHelper
    {
     
        public static bool RegisterProtocol(string scheme)
        {
            try
            {
                string exePath = Process.GetCurrentProcess().MainModule.FileName;
              
                using (var key = Registry.CurrentUser.CreateSubKey($@"Software\Classes\{scheme}"))
                {
                    if (key == null) return false;
                    key.SetValue("", $"URL:{scheme} Protocol");
                    key.SetValue("URL Protocol", ""); 

                    using (var defaultIcon = key.CreateSubKey("DefaultIcon"))
                    {
                        defaultIcon.SetValue("", $"\"{exePath}\",1");
                    }

                    using (var commandKey = key.CreateSubKey(@"shell\open\command"))
                    {
                      
                        commandKey.SetValue("", $"\"{exePath}\" \"%1\"");
                    }
                }
                return true;
            }
            catch
            {
                return false;
            }
        }

        public static bool UnregisterProtocol(string scheme)
        {
            try
            {
                Registry.CurrentUser.DeleteSubKeyTree($@"Software\Classes\{scheme}", false);
                return true;
            }
            catch
            {
                return false;
            }
        }
    }
}
