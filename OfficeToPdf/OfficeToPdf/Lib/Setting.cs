using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;

namespace OfficeToPdf.Lib
{
    public class Setting
    {
        const string SETTING_FILE = "setting.json";

        public int Left { get; set; }
        public int Top { get; set; }

        public void Init()
        {
            Left = 100;
            Top = 100;
        }

        public static Setting Load()
        {
            Setting setting = null;
            try
            {
                setting = JsonSerializer.Deserialize<Setting>(File.ReadAllText(
                    Path.Combine(
                        Path.GetDirectoryName(Process.GetCurrentProcess().MainModule.FileName),
                        SETTING_FILE)));
            }
            catch { }
            if (setting == null)
            {
                setting = new Setting();
            }
            return setting;
        }

        public void Save()
        {
            File.WriteAllText(
                Path.Combine(
                    Path.GetDirectoryName(Process.GetCurrentProcess().MainModule.FileName),
                    SETTING_FILE),
                JsonSerializer.Serialize(this));
        }
    }
}
