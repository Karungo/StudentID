using System;
using System.IO;
using System.Net.Http;
using System.Threading.Tasks;

namespace StudentID
{
    public class LicenseManager
    {
        // Get the user's desktop path and set the license file path
        private string _licenseFilePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "licenses.txt");

        // License verification that checks only online
        public async Task<bool> VerifyLicenseOnline()
        {
            try
            {
                // Example of how to check the license validity using an online GitHub repo or API
                HttpClient client = new HttpClient();
                string url = "https://raw.githubusercontent.com/Karungo/License/master/licenses.json"; // Replace with your actual API endpoint

                HttpResponseMessage response = await client.GetAsync(url);
                string result = await response.Content.ReadAsStringAsync();

                string localLicenseKey = File.ReadAllText(_licenseFilePath);

                // Check if the result contains a valid license key
                return result.Contains(localLicenseKey);
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred while verifying the license online: " + ex.Message);
                return false; // If there's an error, assume the license is invalid
            }
        }

        // Periodic license verification that only checks online
        public async Task PeriodicLicenseVerification(int intervalInHours)
        {
            while (true)
            {
                bool isLicenseValid = await VerifyLicenseOnline();

                if (!isLicenseValid)
                {
                    Console.WriteLine("License is invalid. Application will be disabled.");
                    // Implement application-disable logic here
                    DisableApplication();
                    break;
                }

                // Wait for the next verification interval
                await Task.Delay(TimeSpan.FromHours(intervalInHours));
            }
        }

        // Method to disable the application if the license is invalid
        private void DisableApplication()
        {
            // Example of how you could disable the app - you can implement this based on your app's requirements
            Environment.Exit(0); // Exits the application
        }
    }
}
