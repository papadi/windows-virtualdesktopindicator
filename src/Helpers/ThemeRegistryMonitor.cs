using Microsoft.Win32;
using System;
using System.IO;

namespace VirtualDesktopIndicator.Helpers
{
    public sealed class ThemeRegistryMonitor: IDisposable
    {
        private const string RegistryThemeDataPath = @"HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize";

        private RegistryMonitor themePathRegistryMonitor;

        public bool IsDark { get; private set; }

        public ThemeRegistryMonitor()
        {
            IsDark = GetIsDark();
            
            themePathRegistryMonitor = new RegistryMonitor(RegistryThemeDataPath);
            themePathRegistryMonitor.RegChanged += OnThemeRegistryChanged;
            themePathRegistryMonitor.Error += OnThemeRegistryError;
        }

        private bool GetIsDark()
        {
            return (int)Registry.GetValue(RegistryThemeDataPath, "SystemUsesLightTheme", 0) != 1;
        }

        private void OnThemeRegistryChanged(object sender, EventArgs e)
        {
            try
            {
                IsDark = GetIsDark();
            }
            catch (Exception)
            {
                // TODO: LOG ERROR
            }
        }

        private void OnThemeRegistryError(object sender, ErrorEventArgs e)
        {
            // TODO: LOG ERRRO
        }

        public void Dispose()
        {
            if (themePathRegistryMonitor != null)
            {
                themePathRegistryMonitor.Stop();
                themePathRegistryMonitor.RegChanged -= OnThemeRegistryChanged;
                themePathRegistryMonitor.Error -= OnThemeRegistryError;
                themePathRegistryMonitor = null;
            }
        }
    }
}
