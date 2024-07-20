using Microsoft.Office.Tools.Ribbon;
using Microsoft.Win32;
using System.Timers;

namespace FilterAddin
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            string path = (string)Registry.GetValue(CONFIGS.keyName, "", "");
            filepath.Text = string.IsNullOrEmpty(CONFIGS.SettingFilePath) ? path : CONFIGS.SettingFilePath;
            StartLoadingSettings();
        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            Helper.LoadSettings();
        }
        private void Filepath_TextChanged(object sender, Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs e)
        {
            CONFIGS.SettingFilePath = filepath.Text;
            Registry.SetValue(CONFIGS.keyName, "", CONFIGS.SettingFilePath);
        }

        private static Timer _timer;

        public static void StartLoadingSettings()
        {
            _timer = new Timer(60 * 60 * 1000); // 1 hour in milliseconds
            _timer.Elapsed += LoadSettingsElapsed;
            _timer.Start();
        }

        private static void LoadSettingsElapsed(object sender, ElapsedEventArgs e)
        {
            Helper.LoadSettings();
        }
    }
}
