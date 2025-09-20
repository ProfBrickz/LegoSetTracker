# LegoSetTracker
Have you ever taken a bunch of Lego sets apart and put them in the same bin?
Do you want to rebuild one of the sets?

This Electron application helps you track your Lego sets and pieces with an easy-to-use desktop interface.


## Installation
Download the latest release from the [Releases page](https://github.com/username/LegoSetTracker/releases) and:
- Run the installer (`.exe` or `.msi` file) for Windows
- Extract the archive (`.zip`, `.7z`, or `.tar.gz` file) and run the executable


## Usage
Start the application:
- Open the terminal in the project directory
- Run the following command: `npm start` or `electron .`

For development:
- Install [Node.js](https://nodejs.org/) from the official website
- Download this repository
- Install dependencies:
  - For npm use `npm install`
  - For pnpm use `pnpm install`

The application will open in a desktop window where you can:
- Navigate between different pages using the top navigation
- Manage your Lego sets
- Track total pieces
- Configure settings


## Development
This is an Electron application with the following structure:
- `src/main.js` - Main Electron process
- `src/views/` - HTML templates and pages
- `src/static/` - CSS styles and JavaScript files


## License

This project is licensed under the [GNU General Public License v2.0 only](https://www.gnu.org/licenses/old-licenses/gpl-2.0.html).

> [!IMPORTANT]
> The icons used in this application are not part of the main project license and are sourced from the [Electron Website](https://github.com/electron/website). They are licensed under the [Apache License 2.0](https://www.apache.org/licenses/LICENSE-2.0.html). These icons are provided as a temporary solution and may be replaced in future updates.
