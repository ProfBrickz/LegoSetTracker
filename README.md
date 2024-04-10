# LegoSetTracker
Have you ever taken a bunch of Lego sets apart and put them in the same bin?
Do you want to rebuild one of the sets?

This program will make an Excel spreadsheet so you can keep track of the parts you found and what parts are left to find.

## Installation
- Install the .Net SDK from [Microsoft](https://dotnet.microsoft.com/en-us/)
- Download this repository
- Before running it download the modules with `npm i`, copy the `settings-template.json` file, and name the copy `settings.json`.
- Google Chrome or Windows Defender may think the app.js file is a virus from the zip file. If this happens try having Windows Defender ignore the app.js file, temporarily turning off Windows Defender and turning it back on after the download, or downloading it with git bash instead.

## Usage
Start the program
- Open the terminal
- Run the following command `node app.js`

When the program runs you are prompted for the set number of the set you want to find.
After inputting a set number you can change some of the settings for the Excel file it makes.

In the Excel file, you should only change the green cells.

There is a second worksheet for settings.
You can change how it sorts the rows with the sort setting.
