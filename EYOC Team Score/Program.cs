using EYOC_Team_Score.Model;
using Microsoft.Extensions.Configuration;

namespace EYOC_Team_Score
{
    internal static class Program
    {
        /// <summary>
        ///  The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            // Initialize Winform application
            ApplicationConfiguration.Initialize();
            Application.Run(new EYOCTeamScore());
        }
    }
}