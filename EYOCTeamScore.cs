using EYOC_Team_Score.Model;
using EYOC_Team_Score.Util;
using Microsoft.Extensions.Configuration;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace EYOC_Team_Score
{
    public partial class EYOCTeamScore : Form
    {
        // Point tables for calculation of team scores
        Scores[]? Scores;
        // Parser for IOF XML 3.0 file
        EventFileParser EventFileParser;
        // Events
        List<Event> ActiveEvents;
        BindingSource bindingSource_Events = new BindingSource();
        // Total scores
        Dictionary<string, int> CountryScores = new Dictionary<string, int>();

        // Fonts used
        private static Font fontNormal = new Font("Segoe UI", 9F, FontStyle.Regular);
        private static Font fontBold = new Font("Segoe UI", 9F, FontStyle.Bold);

        public EYOCTeamScore()
        {
            ReadPointTables();
            InitializeComponent();
            ActiveEvents = new List<Event>();
            bindingSource_Events.DataSource = ActiveEvents;
            listbox_ResultFiles.DataSource = bindingSource_Events;
            EventFileParser = new EventFileParser();
        }

        //
        // Helper functions
        //
        private void ReadPointTables()
        {
            try
            {
                IConfigurationRoot config = new ConfigurationBuilder()
                .AddJsonFile("appsettings.json")
                .Build();

                // Get Point tables for score calculation
                Scores = config.GetRequiredSection("PointTables").Get<Scores[]>();
                if (Scores == null || Scores.Length == 0)
                {
                    DialogResult result = MessageBox.Show("Point tables for score calculation couldn't be loaded from 'appsettings.json'.", "Failed to load point tables", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    Close();
                    throw new Exception();
                }
            }
            catch (FileNotFoundException)
            {
                DialogResult result = MessageBox.Show("File 'appsettings.json' is not found.", "Failed to load point tables", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Close();
                throw;
            }
            catch (InvalidOperationException)
            {
                DialogResult result = MessageBox.Show("Point tables for score calculation couldn't be found in 'appsettings.json'.", "Failed to load point tables", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Close();
                throw;
            }
        }

        private void calculateTeamScores()
        {
            foreach (Event evt in ActiveEvents)
            {
                Scores scoreTable = Scores.Where(s => s.Type == evt.Type).FirstOrDefault();
                foreach (Clazz cls in evt.Clazzes)
                {
                    cls.PersonsOrTeams.Sort((a, b) => { return a.Time.CompareTo(b.Time); });
                    foreach (PersonOrTeam pot in cls.PersonsOrTeams)
                    {
                        pot.Score = pot.Place > scoreTable.Points.Length ? scoreTable.Points[scoreTable.Points.Length - 1] : scoreTable.Points[pot.Place - 1];
                    }
                }
            }
        }

        private void calculateTotalScores()
        {
            CountryScores = new Dictionary<string, int>();
            foreach (Event evt in ActiveEvents)
            {
                int countingLimit = evt.Type == EventType.Relay ? 1 : 2;  // The first two counts in individual events, but only the first relay team counts.
                foreach (Clazz cls in evt.Clazzes)
                {
                    var countingPerCountry = new Dictionary<string, int>();  // Number of persons/teams are counting at this time
                    cls.PersonsOrTeams.Sort((a, b) => { return a.Time.CompareTo(b.Time); });
                    foreach (PersonOrTeam pot in cls.PersonsOrTeams)
                    {
                        // Find potential score
                        var score = 0;
                        if (countingPerCountry.ContainsKey(pot.Country))
                        {
                            if (countingPerCountry.GetValueOrDefault(pot.Country) < countingLimit)
                            {
                                score = pot.Score;
                                countingPerCountry[pot.Country]++;
                            }
                        }
                        else
                        {
                            score = pot.Score;
                            countingPerCountry[pot.Country] = 1;
                        }

                        // Add to country's total
                        if (CountryScores.ContainsKey(pot.Country))
                        {
                            CountryScores[pot.Country] += score;
                        }
                        else
                        {
                            CountryScores.Add(pot.Country, score);
                        }
                    }
                }
            }
        }

        private void updateIndividualReport()
        {
            StringBuilder report = new StringBuilder($@"
<html>
<head>
  <link rel=""stylesheet"" href=""teamscore.css"">
</head>
<body>
<h1 class='o-h1'>EYOC Team Score</h1>
");
            foreach (Event evt in ActiveEvents)
            {
                report.AppendLine($"<h2 class='o-h2'>{evt.Name}</h2>");
                foreach (Clazz cls in evt.Clazzes)
                {
                    report.AppendLine($"<h3 class='o-h3'>{cls.Name}</h3>");
                    report.AppendLine("<table class='o-table'><tr class='o-tr'><th class='o-th'></th><th class='o-th'>Name</th><th class='o-th'>Country</th><th class='o-th'>Time</th><th class='o-th'>Points</th></tr>");
                    cls.PersonsOrTeams.Sort((a, b) => { return a.Time.CompareTo(b.Time); });
                    foreach (PersonOrTeam pot in cls.PersonsOrTeams)
                    {
                        report.AppendLine($"<tr class='o-tr'><td class='o-td'>{pot.Place}</td><td class='o-td'>{pot.Name}</td><td class='o-td'>{pot.Country}</td><td class='o-td'>{secondsToHHMMSS(pot.Time)}</td><td class='o-td'>{pot.Score}</td></tr>");
                    }
                    report.AppendLine("</table>");
                }
            }
            report.AppendLine("</body></html>");

            htmlPanel_Individual.Text = report.ToString();
        }

        private void updateTotalReport()
        {
            StringBuilder report = new StringBuilder($@"
<html>
<head>
  <link rel=""stylesheet"" href=""teamscore.css"">
</head>
<body>
<h1 class='o-h1'>EYOC Team Score - Result</h1>
<p class='o-p'>Based on
<ul class='o-ul'>
");
            foreach (Event evt in ActiveEvents)
            {
                report.AppendLine($"<li class='o-li'>{evt.Name}</li>");
            }
            report.AppendLine("</ul>");

            report.AppendLine("<table class='o-table'><tr class='o-tr'><th class='o-th'></th><th class='o-th'>Country</th><th class='o-th'>Points</th></tr>");
            var sortedDict = from country in CountryScores orderby country.Value descending select country;
            int place = 0;
            int sameplace = 1;
            int lastCountrysScore = 0;
            foreach (var country in sortedDict)
            {
                // Calculate place with same points getting same place.
                if (country.Value != lastCountrysScore)
                {
                    place += sameplace;
                    sameplace = 1;
                }
                else
                {
                    sameplace++;
                }
                report.AppendLine($"<tr class='o-tr'><td class='o-td'>{place}</td><td class='o-td'>{country.Key}</td><td class='o-td'>{country.Value}</td></tr>");
                lastCountrysScore = country.Value;
            }
            report.AppendLine("</table>");
            report.AppendLine("</body></html>");

            htmlPanel_Total.Text = report.ToString();
        }

        private string secondsToHHMMSS(int seconds)
        {
            TimeSpan time = TimeSpan.FromSeconds(seconds);
            return time.Hours > 0 ? time.ToString(@"h\:mm\:ss") : time.ToString(@"m\:ss");
        }

        //
        // GUI Events
        //
        private void ImportFile_Click(object sender, EventArgs e)
        {
            if (openResultFileDialog.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    Event @event = EventFileParser.Parse(openResultFileDialog.FileName);
                    if (ActiveEvents.Exists(e => e.Id == @event.Id))
                    {
                        MessageBox.Show("This event is already loaded.", "Duplicate", MessageBoxButtons.OK);
                        return;
                    }
                    ActiveEvents.Add(@event);
                    bindingSource_Events.ResetBindings(false);
                    btn_CalculateTeamScores.Font = fontBold;
                }
                catch (IofXmlParseException ex)
                {
                    MessageBox.Show("The file doesn't seem to contain a valid IOF XML 3.0 file.", ex.Message, MessageBoxButtons.OK);
                }
            }
        }

        private void ExportFile_Click(object sender, EventArgs e)
        {
            if (exportTeamScoresIndividualDialog.ShowDialog() == DialogResult.OK)
            {
                var filename = exportTeamScoresIndividualDialog.FileName;
                File.WriteAllText(filename, htmlPanel_Individual.Text);
            }
            if (exportTeamScoresTotalDialog.ShowDialog() == DialogResult.OK)
            {
                var filename = exportTeamScoresTotalDialog.FileName;
                File.WriteAllText(filename, htmlPanel_Total.Text);
            }
        }

        private void btn_CalculateTeamScores_Click(object sender, EventArgs e)
        {
            calculateTeamScores();
            updateIndividualReport();
            calculateTotalScores();
            updateTotalReport();
            btn_CalculateTeamScores.Font = fontNormal;
        }

        private void ctxmenu_DeleteResultFile_Click(object sender, EventArgs e)
        {
            // Delete the selected file from the EventList
            ActiveEvents.RemoveAt(listbox_ResultFiles.SelectedIndex);
            bindingSource_Events.ResetBindings(false);
            btn_CalculateTeamScores.Font = fontBold;
        }

        private void listbox_ResultFiles_MouseUp(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                MouseEventArgs myMEArs = new MouseEventArgs(MouseButtons.Left, e.Clicks, e.X, e.Y, e.Delta);
                var p = new Point(e.X, e.Y);
                int selectedIndx = this.listbox_ResultFiles.IndexFromPoint(p);
                if (selectedIndx != ListBox.NoMatches)
                {
                    listbox_ResultFiles.SelectedIndex = selectedIndx;
                    ctxmenu_DeleteResultFile.Show((ListBox)sender, p);
                }
            }
        }

    }
}
