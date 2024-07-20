using Microsoft.Win32;
using Newtonsoft.Json;
using System;
using System.Configuration;
using System.IO;
using System.Windows.Forms;

namespace FilterAddin
{
    public class Helper
    {
        public static bool LoadSettings()
        {
            try
            {
                string path = (string)Registry.GetValue(CONFIGS.keyName, "", "");
                if (!string.IsNullOrEmpty(path))
                {
                    CONFIGS.SettingFilePath = path;
                }

                if (!File.Exists(CONFIGS.SettingFilePath))
                {
                    MessageBox.Show($"File doesn't exisits or not accessible: {CONFIGS.SettingFilePath}. Please add file path from Add-ins menu.", "Alert", MessageBoxButtons.OK, MessageBoxIcon.Warning);

                    if (CONFIGS.Settings == null)
                    {
                        string jsonData = (string)Registry.GetValue(CONFIGS.keyDataName, "", "");
                        if (!string.IsNullOrEmpty(jsonData))
                        {
                            CONFIGS.Settings = JsonConvert.DeserializeObject<Settings>(jsonData);
                        }
                    }
                }
                else
                {
                    try
                    {
                        Registry.SetValue(CONFIGS.keyName, "", CONFIGS.SettingFilePath);

                        //MessageBox.Show(tExpand, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        CONFIGS.Settings = JsonConvert.DeserializeObject<Settings>(File.ReadAllText(CONFIGS.SettingFilePath));
                        string jsonSetting = JsonConvert.SerializeObject(CONFIGS.Settings);
                        Registry.SetValue(CONFIGS.keyDataName, "", jsonSetting);
                    }
                    catch (System.Exception ex)
                    {
                        MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    return true;
                }

                
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            return false;
        }
    }
}
