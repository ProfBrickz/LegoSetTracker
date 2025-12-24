# LegoSetTracker
Have you ever taken a bunch of Lego sets apart and put them in the same bin?
Do you want to rebuild one of the sets?

This Electron application helps you track your Lego sets and pieces with an easy-to-use desktop interface.


## Installation
Download the latest release from the [Releases page](https://github.com/username/LegoSetTracker/releases) and:
- Run the installer (`.exe` or `.msi` file) for Windows
- Extract the archive (`.zip`, `.7z`, or `.tar.gz` file) and run the application

For development:
- Install [Node.js](https://nodejs.org/) from the official website
- Download or clone this repository
- Install dependencies:
  - For npm use `npm install`
  - For pnpm use `pnpm install`
- Start the application using `npm start` or `electron .`


## Development
This is an Electron application with the following structure:
- `src/main.js` - Main Electron process
- `src/views/` - HTML templates and pages
- `src/static/` - CSS styles and JavaScript files


## Versioning
This project uses a **custom versioning system** based on the `Major.Minor.Patch` format, with the following definitions:

- **Major**: Incremented for **large-scale changes**, **core feature additions**, **reaching a milestone** (e.g., `0.X.X` → `1.0.0`), or when **changes from a series of Minor versions accumulate into a significant update**.

- **Minor**: Incremented for **new features**, **major bug fixes**, or **refactoring** that may require changes to existing functionality. **Breaking changes** (e.g., API changes or removal of features) are allowed in Minor versions, which differs from standard Semantic Versioning (SemVer).

- **Patch**: Incremented for **small bug fixes**, **quality-of-life improvements**, or **minor feature additions** that probably do not break compatibility or require significant changes to the system.

> **Note**: This versioning system is tailored for this project and may differ from standard Semantic Versioning (SemVer), which prohibits breaking changes in Minor versions. This approach allows for more flexibility during development and feature additions.

> **Note**: I will try to mention in the patch notes which updates have breaking changes and what they are. If you find any breaking changes, please make an issue for it.


## License
This project is licensed under the [GNU General Public License v2.0 only](https://www.gnu.org/licenses/old-licenses/gpl-2.0.html).

> [!IMPORTANT]
> The icons used in this application are not part of the main project license and are sourced from the [Electron Website](https://github.com/electron/website). They are licensed under the [Apache License 2.0](https://www.apache.org/licenses/LICENSE-2.0.html). These icons are provided as a temporary solution and may be replaced in future updates.
