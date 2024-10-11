using System;
using System.Threading.Tasks;
using System.Windows;

namespace StudentID
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : Application
    {
        private LicenseManager _licenseManager;

        protected override async void OnStartup(StartupEventArgs e)
        {
            base.OnStartup(e);

            // Initialize the LicenseManager
            _licenseManager = new LicenseManager();

            // Check if the license is valid online when the app starts
            bool isLicenseValid = await _licenseManager.VerifyLicenseOnline();
            if (!isLicenseValid)
            {
                MessageBox.Show("The license is invalid. The application will now close.", "License Error", MessageBoxButton.OK, MessageBoxImage.Error);
                Shutdown(); // Close the app if the license is invalid
                return; // Ensure no further execution
            }

            // Only show the main window if the license is valid
            MainWindow mainWindow = new MainWindow();
            mainWindow.Show();

            // Start the periodic license verification (e.g., every 24 hours)
            Task.Run(() => _licenseManager.PeriodicLicenseVerification(24));
        }

        protected override void OnExit(ExitEventArgs e)
        {
            base.OnExit(e);

            // Perform any necessary cleanup (if needed)
        }
    }
}
