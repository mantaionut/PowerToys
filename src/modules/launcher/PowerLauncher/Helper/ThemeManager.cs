// Copyright (c) Microsoft Corporation
// The Microsoft Corporation licenses this file to you under the MIT license.
// See the LICENSE file in the project root for more information.

using System;
using System.IO;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Threading;
using ControlzEx.Theming;
using ManagedCommon;
using Microsoft.Office.Interop.OneNote;
using Microsoft.Win32;
using UnitsNet;
using Wox.Infrastructure.Image;
using Wox.Infrastructure.UserSettings;
using static PowerLauncher.Helper.WindowsInteropHelper;

namespace PowerLauncher.Helper
{
    public class ThemeManager : IDisposable
    {
        private readonly PowerToysRunSettings _settings;
        private readonly MainWindow _mainWindow;
        private ManagedCommon.Theme _currentTheme;
        private bool _disposed;

        public ManagedCommon.Theme CurrentTheme => _currentTheme;

        public event Common.UI.ThemeChangedHandler ThemeChanged;

        public ThemeManager(PowerToysRunSettings settings, MainWindow mainWindow)
        {
            _settings = settings;
            _mainWindow = mainWindow;
            SystemEvents.UserPreferenceChanged += OnUserPreferenceChanged;
        }

        private void OnUserPreferenceChanged(object sender, UserPreferenceChangedEventArgs e)
        {
            ManagedCommon.Theme theme = ThemeExtensions.GetCurrentTheme();
            if (e.Category == UserPreferenceCategory.General)
            {
                UpdateTheme();
            }
            else if (e.Category == UserPreferenceCategory.Color)
            {
                if (_currentTheme is ManagedCommon.Theme.Dark or ManagedCommon.Theme.Light)
                {
                    UpdateTheme();
                }
            }
        }

        public void PrintKeysToFile(ResourceDictionary resourceDictionary, string filePath)
        {
            try
            {
                // Open the file for writing
                using (StreamWriter writer = new(filePath, false))
                {
                    // Iterate through all the keys in the resource dictionary
                    foreach (object key in resourceDictionary.Keys)
                    {
                        writer.WriteLine(key.ToString());  // Write each key to the file
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
        }

        private void SetSystemTheme(ManagedCommon.Theme theme)
        {
            _mainWindow.Resources.MergedDictionaries.Clear();
            string themeString = theme switch
            {
                ManagedCommon.Theme.Light => "Themes/Light.xaml",
                ManagedCommon.Theme.Dark => "Themes/Dark.xaml",
                ManagedCommon.Theme.HighContrastOne => "Themes/HighContrast1.xaml",
                ManagedCommon.Theme.HighContrastTwo => "Themes/HighContrast2.xaml",
                ManagedCommon.Theme.HighContrastWhite => "Themes/HighContrastWhite.xaml",
                _ => "Themes/HighContrastBlack.xaml",
            };

            if (theme is ManagedCommon.Theme.Dark or ManagedCommon.Theme.Light)
            {
                // Step 2: Create a new ResourceDictionary pointing to Fluent.xaml
                ResourceDictionary fluentThemeDictionary = new()
                {
                    Source = new Uri("pack://application:,,,/PresentationFramework.Fluent;component/Themes/Fluent.xaml", UriKind.Absolute),
                };
                _mainWindow.Resources.MergedDictionaries.Add(fluentThemeDictionary);
            }
            else
            {
                _mainWindow.Resources.MergedDictionaries.Add(new ResourceDictionary
                {
                    Source = new Uri("Styles/FluentHC.xaml", UriKind.Relative),
                });
            }

            _mainWindow.Resources.MergedDictionaries.Add(new ResourceDictionary
            {
                Source = new Uri("Styles/Styles.xaml", UriKind.Relative),
            });
            ImageLoader.UpdateIconPath(theme);
            ThemeChanged(theme, _currentTheme);
            _currentTheme = theme;
        }

        public void UpdateTheme()
        {
            ManagedCommon.Theme newTheme = _settings.Theme;
            ManagedCommon.Theme theme = ThemeExtensions.GetHighContrastBaseType();
            if (theme != ManagedCommon.Theme.Light)
            {
                newTheme = theme;
            }
            else if (_settings.Theme == ManagedCommon.Theme.System)
            {
                newTheme = ThemeExtensions.GetCurrentTheme();
            }

            _mainWindow.Dispatcher.Invoke(() =>
            {
                SetSystemTheme(newTheme);
            });
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (_disposed)
            {
                return;
            }

            if (disposing)
            {
                SystemEvents.UserPreferenceChanged -= OnUserPreferenceChanged;
            }

            _disposed = true;
        }
    }
}
