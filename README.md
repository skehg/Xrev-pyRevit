# PyRevit-Extensions
Kents Tools for PyRevit

## Xrev Internal Extension

This extension adds a custom tab called "Xrev Internal" to Autodesk Revit with the following panels:

### Favourites Panel
- **Sample Button** - A large button for quick access to favourite tools
- **Favourites Pulldown** - A dropdown menu with additional options:
  - Option A
  - Option B

### Management Panel
- **Manage Tool** - A large button for management functions
- **Management Pulldown** - A dropdown menu with admin and settings options:
  - Admin Option
  - Settings Option

## Installation

1. Copy the `Xrev Internal.extension` folder to your pyRevit extensions folder
2. Reload pyRevit or restart Revit
3. The "Xrev Internal" tab will appear in the Revit ribbon

## Structure

```
Xrev Internal.extension/
├── bundle.yaml
└── Xrev Internal.tab/
    ├── Favourites.panel/
    │   ├── Sample Button.pushbutton/
    │   └── Favourites Pulldown.pulldown/
    └── Management.panel/
        ├── Manage Tool.pushbutton/
        └── Management Pulldown.pulldown/
```
