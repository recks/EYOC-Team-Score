# EYOC Team Score

This is a Windows application for calculating Team Scores at European Youth Orenteering Championships.
It uses IOF XML 3.0 result files as input and can export as either HTML or Excel file.
The scores are calculated according to IOF EYOC Manual v1.4 2023.

## Quick guide

- You load a result file into the application using the Import Result File button. Multiple files can be imported by repeating. A file can be removed by right-clicking the file in the list of files.
- When you are satisfied, you click the "Calculate Team Score" and the calculated score can be seen - both as a total and in a detailed view.
- The scores can be exported both as HTML and spreadsheet using the File menu. You can change whether CSS should be exported with the HTML file using the Tools -> Options menu.

## How it works

Calculation is based on the application being able to find the *three letter country code* (for each competitor) in the result file. The scores of each competitor is then summmed for each country and the result is shown using the three letter code for the country.

It tries to find the code by looking at three different places in the `<Organisation>` element of the result file:

- Organisation Name
- Country "code" attribute
- Country Name

If it can't find any country code for the competitor, he/she will be attached to the glory country of Unknown (UNK).

### How to ensure this

In your Event software you should use the three letter code as either the "club" name or the club's nationality (as either the country name or specific country code if that is suppported).

### Caveats

- It doesn't matter whether you use ISO 3166 or IOC country codes, just be consistent and don't mix. It is important that the same country has the same code in all result files.
- Rememeber that AUT is Austria and AUS is Australia.

