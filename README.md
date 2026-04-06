# LegoSetTracker
Have you ever taken a bunch of Lego sets apart and put them in the same bin?
Do you want to rebuild one of the sets?

This Electron application helps you track your Lego sets and pieces with an easy-to-use desktop interface.


## Installation
Download the latest release from the [Releases page](https://github.com/ProfBrickz/LegoSetTracker/releases) and:
- Run the installer (`.exe` or `.msi` file) for Windows
- Extract the archive (`.zip`, `.7z`, or `.tar.gz` file) and run the application

For development:
- Install [Node.js](https://nodejs.org/) from the official website
- Download or clone this repository
- Install dependencies using pnpm:
  - `pnpm install`
- Start the application using `pnpm dev` or `pnpm start`


## Development
This is an Electron application with the following structure:
- `src/main.js` - Main Electron process
- `src/renderer/views/pages/` - HTML templates and pages
- `src/renderer/views/layouts/` - Layout templates
- `src/static/` - Client side CSS styles and JavaScript files
- `src/database/` - Database related files
- `src/classes.js` - Custom classes
- `src/constants.js` - Application constants
- `src/preload/preload.js` - Preload script for Electron
- `tests/` - Unit test files and fixtures


> [!Note]
> Most of the UI design and styling has been developed with the assistance of AI. The application leverages AI-powered design tools to create an intuitive and visually appealing user interface.


## Versioning
This project uses a **custom versioning system** based on the `Major.Minor.Patch` format, with the following definitions:

- **Major**: Incremented for **large-scale changes**, **core feature additions**, **reaching a milestone** (e.g., `v0.X.X` → `v1.0.0`), or when **changes from a series of Minor versions accumulate into a significant update**.

- **Minor**: Incremented for **new features**, **major bug fixes**, or **refactoring** that may require changes to existing functionality. **Breaking changes** (e.g., API changes or removal of features) are allowed in Minor versions, which differs from standard Semantic Versioning (SemVer).

- **Patch**: Incremented for **small bug fixes**, **quality-of-life improvements**, or **minor feature additions** that probably do not break compatibility or require significant changes to the system.

> **Note**: This versioning system is tailored for this project and may differ from standard Semantic Versioning (SemVer), which prohibits breaking changes in Minor versions. This approach allows for more flexibility during development and feature additions.

> **Note**: I will try to mention in the patch notes which updates have breaking changes and what they are. If you find any breaking changes, please make an issue for it.


## License
This project is licensed under the [GNU General Public License v2.0 only](https://www.gnu.org/licenses/old-licenses/gpl-2.0.html).
