# Penguins Walking in Lockstep - Excel Edition

## Quick Start Guide

### Method 1: Using VBA Macros (Recommended)

1. **Open Microsoft Excel** (works with Excel 2010 or later)

2. **Create a new workbook**

3. **Enable the Developer tab** (if not already visible):
   - File → Options → Customize Ribbon
   - Check "Developer" on the right side
   - Click OK

4. **Open VBA Editor**:
   - Press `Alt + F11` (Windows)
   - Or press `Option + F11` (Mac)

5. **Insert a new module**:
   - In the VBA Editor: Insert → Module

6. **Copy the code**:
   - Open `penguins.vba` file
   - Copy ALL the code (Ctrl+A, Ctrl+C)
   - Paste into the module window

7. **Return to Excel**:
   - Press `Alt + F11` again

8. **Run the animation**:
   - Press `Alt + F8` to open Macros dialog
   - Select `CreatePenguinShow` and click "Run"
   - Once penguins are created, press `Alt + F8` again
   - Select `AnimatePenguins` and click "Run"

9. **Watch the penguins march in perfect lockstep!**

### Features

- 🐧 5 adorable Excel penguins
- ⬆️ Synchronized movement (all penguins move together)
- 🎯 Waddle animation (up and down bobbing)
- 🎨 Color-coded with black bodies, white bellies, orange beaks and feet
- ♻️ Reset button to start over

### Macros Available

- `CreatePenguinShow` - Creates the penguins on the worksheet
- `AnimatePenguins` - Makes them walk across the screen in lockstep
- `ResetPenguins` - Deletes and recreates the penguins from the start

### Troubleshooting

**"Macros are disabled"**
- File → Options → Trust Center → Trust Center Settings
- Macro Settings → Enable all macros (or enable with notification)
- Restart Excel

**"Code doesn't run"**
- Make sure you pasted ALL the code from penguins.vba
- Check that you're using Excel (not Google Sheets or LibreOffice)

**"Penguins move too fast/slow"**
- Edit the VBA code
- Find this line: `Application.Wait Now + TimeValue("00:00:00.05")`
- Change `00:00:00.05` to increase (slower) or decrease (faster) the delay

### Method 2: View HTML Version in Excel

If you prefer the original HTML animation:

1. Open Excel
2. Insert → Object → Create from File
3. Browse to `penguins.html`
4. Check "Display as icon" if you just want to link to it
5. Or use Insert → Illustrations → Online Pictures to embed web content (Excel 365)

---

**Enjoy your marching penguins!** 🐧🐧🐧🐧🐧
