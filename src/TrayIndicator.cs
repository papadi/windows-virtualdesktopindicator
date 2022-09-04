using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Text;
using System.IO;
using System.Reflection;
using System.Windows.Forms;
using VirtualDesktopIndicator.Api;
using VirtualDesktopIndicator.Helpers;

namespace VirtualDesktopIndicator
{
    class TrayIndicator : IDisposable
    {
        private static string AppName =>
            Assembly.GetExecutingAssembly().GetName().Name;

        private IVirtualDesktopApi virtualDesktopProxy;

        private const uint DesktopErrorIndex = 0;

        private NotifyIcon trayIcon;
        private Timer timer;

        #region Virtual Desktops

        private uint previewVirtualDesktop = 0;

        #endregion

        #region Theme

        private const string RegistryThemeDataPath =
            @"HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize";

        private enum Theme { Light, Dark }

        private static readonly Dictionary<Theme, Color[]> ThemesColors = new Dictionary<Theme, Color[]>()
        {
            { Theme.Dark, new[] { Color.White, Color.Gold, Color.LightGreen, Color.LightSkyBlue } },
            { Theme.Light, new[] { Color.Black, Color.Gold, Color.Blue, Color.DarkGreen } },
        };

        private Color GetCurrentThemeColor(uint desktopIndex)
        {
            if (desktopIndex == 0) return Color.Red;    // failed to get desktop index
            var set = ThemesColors[systemTheme];
            return set[Math.Min(desktopIndex, set.Length) - 1];
        }

        private Theme cachedSystemTheme;
        private Theme systemTheme;

        private RegistryMonitor themePathRegistryMonitor;

        #endregion

        #region Drawing data 

        // Default windows tray icon size
        private const int BaseHeight = 16;
        private const int BaseWidth = 16;

        // We use half the size, because otherwise the image is rendered with incorrect anti-aliasing
        private int Height
        {
            get
            {
                var height = SystemMetricsApi.GetSystemMetrics(SystemMetric.SM_CYICON) / 2;
                return height < BaseHeight ? BaseHeight : height;
            }
        }
        private int Width
        {
            get
            {
                var width = SystemMetricsApi.GetSystemMetrics(SystemMetric.SM_CXICON) / 2;
                return width < BaseWidth ? BaseWidth : width;
            }
        }

        private int BorderThinkness => Width / BaseWidth;

        private int FontSize => (int)Math.Ceiling(Width / 1.5);
        private FontStyle FontStyle = FontStyle.Bold;

        #endregion

        public TrayIndicator()
        {
            trayIcon = new NotifyIcon { ContextMenuStrip = CreateContextMenu() };
            trayIcon.Click += TrayIconClick;

            timer = new Timer { Enabled = false, Interval = 500 };
            timer.Tick += TimerTick;

            InitRegistryMonitor();

            cachedSystemTheme = systemTheme = GetSystemTheme();
        }

        private static IVirtualDesktopApi GetActualDesktopApi()
        {
            if (Environment.OSVersion.Version.Build >= 22000)
            {
                return new Latest();
            }
            else
            {
                return new Previous();
            }
        }

        #region Events

        private void TimerTick(object sender, EventArgs e)
        {
            try
            {
                if (virtualDesktopProxy == null) virtualDesktopProxy = GetActualDesktopApi();
                var currentVirtualDesktop = virtualDesktopProxy.GetCurrent();
                if (currentVirtualDesktop == previewVirtualDesktop) return;

                previewVirtualDesktop = currentVirtualDesktop;

                GenerateIcon(currentVirtualDesktop);
            }
            catch (Exception ex)
            {
                previewVirtualDesktop = DesktopErrorIndex;
                GenerateIcon(DesktopErrorIndex);
                virtualDesktopProxy = null; // failed - set to null to re-initialize in case explorer crashed
                // TODO: LOG THE ERROR
            }
        }

        private void TrayIconClick(object sender, EventArgs e)
        {
            /*
            MouseEventArgs me = e as MouseEventArgs;

            if (me.Button == MouseButtons.Left)
                ShowTaskView();
            */
        }

        #endregion

        public void Display()
        {
            themePathRegistryMonitor.Start();

            trayIcon.Visible = true;
            timer.Enabled = true;
        }

        public void Dispose()
        {
            StopRegistryMonitor();

            trayIcon.Dispose();
            timer.Dispose();
        }

        private void InitRegistryMonitor()
        {
            themePathRegistryMonitor = new RegistryMonitor(RegistryThemeDataPath);

            themePathRegistryMonitor.RegChanged += OnThemeRegistryChanged;
            themePathRegistryMonitor.Error += OnThemeRegistryError;
        }

        private void StopRegistryMonitor()
        {
            if (themePathRegistryMonitor == null) return;

            themePathRegistryMonitor.Stop();
            themePathRegistryMonitor.RegChanged -= OnThemeRegistryChanged;
            themePathRegistryMonitor.Error -= OnThemeRegistryError;
            themePathRegistryMonitor = null;
        }

        public static void ShowTaskView()
        {
            /*
             * Unimplemented!
             * I didn't find a efficient way to launch task viewer.
             * Each of them has problems, but here are some solutions:
             *   1. Run "explorer shell:::{3080F90E-D7AD-11D9-BD98-0000947B0257}"
             *   2. Simulating <Win + Tab>
             */
        }

        private ContextMenuStrip CreateContextMenu()
        {
            var menu = new ContextMenuStrip();

            var autorunItem = new ToolStripMenuItem("Start application at Windows startup")
            {
                Checked = AutorunManager.GetAutorunStatus(AppName, Application.ExecutablePath)
            };
            autorunItem.Click += (sender, e) =>
            {
                autorunItem.Checked = !autorunItem.Checked;

                if (autorunItem.Checked)
                {
                    AutorunManager.AddApplicationToAutorun(AppName, Application.ExecutablePath);
                }
                else
                {
                    AutorunManager.RemoveApplicationFromAutorun(AppName);
                }
            };

            var exitItem = new ToolStripMenuItem("Exit");
            exitItem.Click += (sender, e) => Application.Exit();

            menu.Items.Add(autorunItem);
            menu.Items.Add(new ToolStripSeparator());
            menu.Items.Add(exitItem);

            return menu;
        }

        private void GenerateIcon(uint desktopIndex)
        {
            var color = GetCurrentThemeColor(desktopIndex);
            var font = new Font("", FontSize, FontStyle, GraphicsUnit.Pixel);   // use default font
            var brush = new SolidBrush(color);
            var bitmap = new Bitmap(Width, Height);
            var displayText = desktopIndex == DesktopErrorIndex ? "?" : (desktopIndex < 100 ? desktopIndex.ToString() : "++");

            var g = Graphics.FromImage(bitmap);

            g.SmoothingMode = SmoothingMode.HighSpeed;
            g.TextRenderingHint = TextRenderingHint.SingleBitPerPixelGridFit;

            g.Clear(Color.Transparent);

            // Draw border
            // The g.DrawRectangle always uses anti-aliasing and border looks very poor at such small resolutions
            // Implement own hack!
            var pen = new Pen(color, 1);
            for (int o = 0; o < BorderThinkness; o++)
            {
                // Top
                g.DrawLine(pen, 0, o, Width - 1, o);

                // Right
                g.DrawLine(pen, o, 0, o, Height - 1);

                // Left
                g.DrawLine(pen, Width - 1 - o, 0, Width - 1 - o, Height - 1);

                // Bottom
                g.DrawLine(pen, 0, Height - 1 - o, Width - 1, Height - 1 - o);
            }

            // Draw text
            var textSize = g.MeasureString(displayText, font);

            // Сalculate padding to center the text
            // We can't assume that g.DrawString will round the coordinates correctly, so we do it manually
            var offsetX = (float)Math.Ceiling((Width - textSize.Width) / 2);
            var offsetY = (float)Math.Ceiling((Height - textSize.Height - 2) / 2);

            g.DrawString(displayText, font, brush, offsetX, offsetY);

            // Create icon from bitmap and return it
            // bitmapText.GetHicon() can throw exception
            try
            {
                trayIcon.Icon = Icon.FromHandle(bitmap.GetHicon());
            }
            catch
            {
                trayIcon.Icon = null;
            }
        }

        private Theme GetSystemTheme()
        {
            return (int)Registry.GetValue(RegistryThemeDataPath, "SystemUsesLightTheme", 0) == 1 ?
                     Theme.Light :
                     Theme.Dark;
        }

        private void OnThemeRegistryChanged(object sender, EventArgs e)
        {
            systemTheme = GetSystemTheme();
            if (systemTheme == cachedSystemTheme) return;
            cachedSystemTheme = systemTheme;
        }

        private void OnThemeRegistryError(object sender, ErrorEventArgs e) => StopRegistryMonitor();
    }
}
