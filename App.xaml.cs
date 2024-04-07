using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;

namespace LSS
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : Application
    {
        private void Application_Startup(object sender, StartupEventArgs e)
        {
            // Set the DataDirectory to the app's base directory
            //AppDomain.CurrentDomain.SetData("DataDirectory", System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, ""));

            //string dataDirectory = AppDomain.CurrentDomain.GetData("DataDirectory").ToString();
            //MessageBox.Show("DataDirectory is set to: " + dataDirectory);


            SplashScreen splashScreen = new SplashScreen();
            splashScreen.Show();

            // Perform any loading tasks...
            Thread.Sleep(3000);

            MainWindow mainWindow = new MainWindow();
            mainWindow.Show();

            splashScreen.Close();

        }
    }
}
