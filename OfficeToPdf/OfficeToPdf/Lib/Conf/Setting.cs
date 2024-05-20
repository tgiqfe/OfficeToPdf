﻿using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;

namespace OfficeToPdf.Lib.Conf
{
    internal class Setting
    {
        public double Width { get; set; }
        public double Height { get; set; }
        public double X { get; set; }
        public double Y { get; set; }

        const string _fileName = "setting.json";

        public static Setting Load()
        {
            var path = Path.Combine(Item.WorkingDirectory, _fileName);
            try
            {
                using (var sr = new StreamReader(path, Encoding.UTF8))
                {
                    var setting = JsonSerializer.Deserialize<Setting>(sr.ReadToEnd());
                    return setting;
                }
            }
            catch
            {
                return new Setting()
                {
                    Width = 800,
                    Height = 600,
                    X = 100,
                    Y = 100,
                };
            }
        }

        public void Save()
        {
            var path = Path.Combine(Item.WorkingDirectory, _fileName);
            using (var sw = new StreamWriter(path, false, Encoding.UTF8))
            {
                var json = JsonSerializer.Serialize(this, new JsonSerializerOptions()
                {
                    WriteIndented = true,
                });
                sw.Write(json);
            }
        }
    }
}
