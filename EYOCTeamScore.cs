using EYOC_Team_Score.Model;
using EYOC_Team_Score.Util;
using Microsoft.Extensions.Configuration;
using OfficeOpenXml;
using OfficeOpenXml.ConditionalFormatting;
using OfficeOpenXml.Style;
using OfficeOpenXml.Style.XmlAccess;
using System.Diagnostics.Metrics;
using System.Globalization;
using System.IO;
using System.Text;
using System.Windows.Forms;
using static System.Formats.Asn1.AsnWriter;

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
        List<Country> CountryScores = [];

        // Fonts used
        private static Font fontNormal = new Font("Segoe UI", 9F, FontStyle.Regular);
        private static Font fontBold = new Font("Segoe UI", 9F, FontStyle.Bold);

        // Excel styles
        ExcelNamedStyleXml style;
        ExcelNamedStyleXml countryHeader;

        public EYOCTeamScore()
        {
            ReadPointTables();
            InitializeComponent();
            ActiveEvents = new List<Event>();
            bindingSource_Events.DataSource = ActiveEvents;
            listbox_ResultFiles.DataSource = bindingSource_Events;
            EventFileParser = new EventFileParser();
        }


        #region Score calculation

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

        private void calculateScores()
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
            CountryScores = [];  // Reset
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
                                pot.Counting = true;  // This person/team counts towards the country's total
                            }
                        }
                        else
                        {
                            score = pot.Score;
                            countingPerCountry[pot.Country] = 1;
                            pot.Counting = true;  // This person/team counts towards the country's total
                        }

                        // Add to country's total
                        if (CountryScores.Any(cs => cs.Name == pot.Country))
                        {
                            // Country exists
                            Country country = CountryScores.First(cs => cs.Name == pot.Country);
                            Tuple<string, string> scoreTuple = new Tuple<string, string>(evt.Name, cls.Name);
                            if (country.Scores.ContainsKey(scoreTuple))
                            {
                                country.Scores[scoreTuple] += score;
                            }
                            else
                            {
                                country.Scores.Add(new Tuple<string, string>(evt.Name, cls.Name), score);
                            }
                        }
                        else
                        {
                            // Country doesn't exist
                            CountryScores.Add(new Country()
                            {
                                Name = pot.Country,
                                Scores = new Dictionary<Tuple<string, string>, int>() { { new Tuple<string, string>(evt.Name, cls.Name), score } }
                            });
                        }
                    }
                }
            }
        }

        #endregion

        #region HTML handling

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
                        report.AppendLine($"<tr class='o-tr'><td class='o-td'>{pot.Place}</td><td class='o-td o-l'>{pot.Name}</td><td class='o-td o-l'>{pot.Country}</td><td class='o-td'>{secondsToHHMMSS(pot.Time)}</td><td class='o-td'>{pot.Score}</td></tr>");
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
<table class='o-table'>
<tr class='o-tr'><th class='o-th'></th><th class='o-th'></th> <!-- Event Header -->
");
            var classheader = new StringBuilder("<tr class='o-tr'><th class='o-th'></th><th class='o-th'>Country</th>");
            for (int i = 0; i < ActiveEvents.Count; i++)
            {
                Event evt = ActiveEvents[i];
                foreach (Clazz cls in evt.Clazzes)
                {
                    classheader.Append($"<th class='o-th-{i % 3}'>{cls.Name}</th>");
                }
                classheader.Append($"<th class='o-th-{i % 3}'>Total</th>");
                report.AppendLine($"<th class='o-th-{i % 3}' colspan='{evt.Clazzes.Count + 1}'>{evt.Name}</th>");  // Event header
            }
            report.AppendLine($"</tr>");  // Event header
            report.AppendLine($"{classheader}<th class='o-th-total'>TOTAL</th></tr>");

            // Country scores
            var sortedDict = from country in CountryScores orderby country.TotalScore() descending select country;
            int place = 0;
            int sameplace = 1;
            int lastCountrysScore = 0;
            foreach (var country in sortedDict)
            {
                var totalScore = country.TotalScore();
                // Calculate place with same points getting same place.
                if (totalScore != lastCountrysScore)
                {
                    place += sameplace;
                    sameplace = 1;
                }
                else
                {
                    sameplace++;
                }
                report.Append($"<tr class='o-tr'><td align='right' class='o-td'>{place}</td><td align='center' class='o-td o-l'>{country.Name}</td>");
                foreach (Event evt in ActiveEvents)
                {
                    foreach (Clazz cls in evt.Clazzes)
                    {
                        Tuple<string, string> scoreTuple = new Tuple<string, string>(evt.Name, cls.Name);
                        object score = country.Scores.ContainsKey(scoreTuple) ? country.Scores[scoreTuple] : "";
                        report.Append($"<td align='center' class='o-td'>{score}</td>");
                    }
                    report.Append($"<td align='center' class='o-td'>{country.TotalScoreForEvent(evt.Name)}</td>");
                }
                report.AppendLine($"<td align='center' class='o-td'>{totalScore}</td></tr>");
                lastCountrysScore = totalScore;
            }
            report.AppendLine("</table>");
            report.AppendLine("</body></html>");

            htmlPanel_Total.Text = report.ToString();
        }

        #endregion

        #region Excel handling

        private void initializeExcelStyles(ExcelWorkbook workBook)
        {
            style = workBook.Styles.CreateNamedStyle("Default");
            style.Style.Font.Name = "Calibri";
            style.Style.Font.Bold = false;
            style.Style.Font.Family = 2;
            style.Style.Font.Size = 12;
            style.Style.Font.Color.SetColor(Color.Black);

            style = workBook.Styles.CreateNamedStyle("Header");
            style.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            style.Style.Font.Name = "Calibri";
            style.Style.Font.Bold = true;
            style.Style.Font.Family = 2;
            style.Style.Font.Size = 11;
            style.Style.Font.Color.SetColor(Color.White);
            style.Style.Fill.PatternType = ExcelFillStyle.Solid;
            style.Style.Fill.BackgroundColor.SetColor(Color.DarkGreen);

            style = workBook.Styles.CreateNamedStyle("Country");
            style.Style.Font.Name = "Calibri";
            style.Style.Font.Bold = true;
            style.Style.Font.Family = 2;
            style.Style.Font.Size = 12;
            style.Style.Font.Color.SetColor(Color.Black);
            style.Style.Fill.PatternType = ExcelFillStyle.Solid;
            style.Style.Fill.BackgroundColor.SetColor(Color.LightGreen);

            for (int i = 0; i < 3; i++)
            {
                style = workBook.Styles.CreateNamedStyle("EventHeader" + i);
                style.Style.Font.Bold = true;
                style.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                style.Style.Fill.PatternType = ExcelFillStyle.Solid;
                style.Style.Fill.BackgroundColor.SetColor(GetColour(i, true));

                style = workBook.Styles.CreateNamedStyle("EventScores" + i);
                style.Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                style.Style.Fill.PatternType = ExcelFillStyle.Solid;
                style.Style.Fill.BackgroundColor.SetColor(GetColour(i, false));
            }
        }

        private void writeExcelSheet(string filename)
        {
            using (var package = new ExcelPackage())
            {
                initializeExcelStyles(package.Workbook);

                // Temp variable used to calculate placings
                int place;
                int sameplace;
                int lastCountrysScore;
                // Temp variables used to navigate worksheet
                ExcelWorksheet workSheet;
                int i;
                int j;

                // One sheet per event.
                foreach (Event evt in ActiveEvents)
                {
                    workSheet = package.Workbook.Worksheets.Add(evt.Name);
                    workSheet.DefaultRowHeight = 13;
                    workSheet.Rows.StyleName = "Default";
                    //evt.Clazzes.Sort();

                    // Worksheet header
                    workSheet.Row(1).Height = 20;
                    j = 2;  // column
                    workSheet.Cells[1, j++].Value = "Name";
                    foreach (Clazz cls in evt.Clazzes)
                    {
                        workSheet.Cells[1, j++].Value = cls.Name;
                    }
                    workSheet.Cells[1, j].Value = "Total";
                    workSheet.Cells[1, 1, 1, j].StyleName = "Header";  // Dark green background

                    // Worksheet content
                    var countries = from country in CountryScores orderby country.TotalScoreForEvent(evt.Name) descending select country;
                    place = 0;
                    sameplace = 1;
                    lastCountrysScore = 0;
                    i = 2;  // We start at row 2
                    foreach (var country in countries)
                    {
                        var totalScore = country.TotalScoreForEvent(evt.Name);
                        // Calculate place with same points getting same place.
                        if (totalScore != lastCountrysScore)
                        {
                            place += sameplace;
                            sameplace = 1;
                            lastCountrysScore = totalScore;
                        }
                        else
                        {
                            sameplace++;
                        }

                        // Row content - country
                        workSheet.Row(i).Height = 18;
                        j = 1;  // Column
                        workSheet.Cells[i, j++].Value = place;
                        workSheet.Cells[i, j++].Value = country.Name;
                        foreach (Clazz cls in evt.Clazzes)
                        {
                            // Look at each class for the country
                            Tuple<string, string> scoreTuple = new Tuple<string, string>(evt.Name, cls.Name);
                            int score = country.Scores.ContainsKey(scoreTuple) ? country.Scores[scoreTuple] : 0;
                            workSheet.Cells[i, j++].Value = score != 0 ? score : "";
                        }
                        workSheet.Cells[i, j++].Value = totalScore;
                        workSheet.Cells[i, 1, i, j-1].StyleName = "Country";  // Light green background
                        i++;

                        // Row content - runners
                        var runnerLists = from cls in evt.Clazzes select cls.PersonsOrTeams;
                        j = 3;  // Start column for scores
                        foreach(List<PersonOrTeam> runners in runnerLists)
                        {
                            List<PersonOrTeam> countingRunners = runners.FindAll(r => r.Country == country.Name && r.Counting).ToList();
                            foreach (PersonOrTeam runner in countingRunners)
                            {
                                workSheet.Cells[i, 2].Value = runner.Name;
                                workSheet.Cells[i, j].Value = runner.Score;
                                i++;
                            }
                            j++;
                        }
                    }
                    // Fit columns
                    workSheet.Cells[1, 1, i, j].AutoFitColumns();
                }

                // Sheet with total
                workSheet = package.Workbook.Worksheets.Add("Total");
                workSheet.DefaultRowHeight = 13;
                workSheet.Rows.StyleName = "Default";
                //evt.Clazzes.Sort();

                // Worksheet header
                workSheet.Row(1).Height = 20;
                j = 3;  // column
                for (int k = 0; k < ActiveEvents.Count; k++)
                {
                    int start = j;
                    Event evt = ActiveEvents[k];
                    workSheet.Cells[1, j, 1, j+evt.Clazzes.Count].Merge = true;  // Event header
                    workSheet.Cells[1, j].Value = evt.Name;
                    workSheet.Cells[1, j].StyleName = "EventHeader" + (k % 3);
                    foreach (Clazz cls in evt.Clazzes)
                    {
                        workSheet.Cells[2, j++].Value = cls.Name;
                    }
                    workSheet.Cells[2, j++].Value = "Total";
                    workSheet.Cells[2, start, 2, j - 1].StyleName = "EventScores" + (k % 3);
                    workSheet.Cells[2, start, 2, j - 1].Style.Font.Bold = true;
                }
                workSheet.Cells[2, j].Value = "TOTAL";
                workSheet.Cells[2, j].Style.Font.Bold = true;
                workSheet.Cells[2, j].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                // Worksheet content
                var totalCountries = from country in CountryScores orderby country.TotalScore() descending select country;
                place = 0;
                sameplace = 1;
                lastCountrysScore = 0;
                i = 3;  // We start at row 3
                foreach (var country in totalCountries)
                {
                    j = 1;  // Column
                    var totalScore = country.TotalScore();
                    // Calculate place with same points getting same place.
                    if (totalScore != lastCountrysScore)
                    {
                        place += sameplace;
                        sameplace = 1;
                        lastCountrysScore = totalScore;
                    }
                    else
                    {
                        sameplace++;
                    }

                    workSheet.Cells[i, j++].Value = place;
                    workSheet.Cells[i, j++].Value = country.Name;
                    for (int k = 0; k < ActiveEvents.Count; k++)
                    {
                        Event evt = ActiveEvents[k];
                        foreach (Clazz cls in evt.Clazzes)
                        {
                            Tuple<string, string> scoreTuple = new Tuple<string, string>(evt.Name, cls.Name);
                            workSheet.Cells[i, j].Value = country.Scores.ContainsKey(scoreTuple) ? country.Scores[scoreTuple] : "";
                            workSheet.Cells[i, j++].StyleName = "EventScores" + (k % 3);
                        }
                        workSheet.Cells[i, j].Value = country.TotalScoreForEvent(evt.Name);
                        workSheet.Cells[i, j++].StyleName = "EventScores" + (k % 3);
                    }
                    workSheet.Cells[i, j++].Value = country.TotalScore();

                    i++;
                }


                // Save to file
                if (File.Exists(filename))
                {
                    File.Delete(filename);
                }
                FileStream objFileStrm = File.Create(filename);
                objFileStrm.Close();
                File.WriteAllBytes(filename, package.GetAsByteArray());
            }
        }

        #endregion

        #region Utility methods

        private string secondsToHHMMSS(int seconds)
        {
            TimeSpan time = TimeSpan.FromSeconds(seconds);
            return time.Hours > 0 ? time.ToString(@"h\:mm\:ss") : time.ToString(@"m\:ss");
        }

        // Creates three different colours based on index and if it should be saturated
        private static readonly Color[] Colours =
        {
            ColorTranslator.FromHtml("#CCFFCC"), ColorTranslator.FromHtml("#FFCCCC"), ColorTranslator.FromHtml("#CCCCFF"),
            ColorTranslator.FromHtml("#66FF66"), ColorTranslator.FromHtml("#FF6666"), ColorTranslator.FromHtml("#6666FF")
        };
        private Color GetColour(int index, bool saturated)
        {
            index = index % 3;
            return Colours[index + (saturated ? 3 : 0)];
        }
        #endregion

        #region GUI Events

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

        private void ExportReport_Click(object sender, EventArgs e)
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

        private void ExportSheet_Click(object sender, EventArgs e)
        {
            if (exportExcelDialog.ShowDialog() == DialogResult.OK)
            {
                var filename = exportExcelDialog.FileName;
                writeExcelSheet(filename);
            }
        }

        private void btn_CalculateTeamScores_Click(object sender, EventArgs e)
        {
            calculateScores();
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

        #endregion
    }
}
