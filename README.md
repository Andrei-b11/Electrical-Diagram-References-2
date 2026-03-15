# 📊 PDF Reference Detector & Link Generator

> 🚀 **Transform your electrical diagrams into interactive PDFs with automatic cross-reference detection and clickable links!**

[![Python](https://img.shields.io/badge/Python-3.8+-blue.svg)](https://www.python.org/downloads/)
[![PyQt5](https://img.shields.io/badge/PyQt5-5.15+-green.svg)](https://pypi.org/project/PyQt5/)
[![License](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)

---

## 🎯 What Does This App Do?

This powerful desktop application **automatically detects cross-references** in your PDF electrical diagrams and **generates interactive PDFs** with clickable links that navigate directly to the referenced locations. Say goodbye to manual searching! 🎉

### ✨ Key Features

- 🔍 **Smart Reference Detection** - Automatically finds references like "25-A.0", "5.3-B", etc.
- 🎨 **Customizable Highlights** - Choose colors, animations, and effects for your links
- 📐 **Visual Grid Editor** - Draw your grid layout directly on the PDF
- 💾 **Grid Templates** - Save and reuse grid configurations
- ⚡ **Batch Processing** - Process multiple PDFs at once
- 🎭 **Animated Highlights** - Blink, fade, or pulse effects when clicking links
- 📊 **Statistics Dashboard** - View detailed analysis of detected references

---

## 📥 Installation

### Prerequisites

Make sure you have Python 3.8 or higher installed on your system.

### Step 1: Clone the Repository

```bash
git clone https://github.com/yourusername/pdf-reference-detector.git
cd pdf-reference-detector
```

### Step 2: Install Dependencies

```bash
pip install -r requirements.txt
```

### Step 3: Run the Application

```bash
python main.py
```

---

## 🎮 Quick Start Guide

### 1️⃣ Load Your PDFs

**Three ways to add files:**

- 🖱️ **Drag & Drop** - Simply drag PDF files into the drop zone
- 📂 **Browse Button** - Click "Select Files" to choose PDFs from your computer
- ➕ **Multiple Files** - Add as many PDFs as you need for batch processing

![Drop Zone](https://via.placeholder.com/800x200/FF6B2B/FFFFFF?text=Drop+PDF+Files+Here)

> 💡 **Tip:** You can remove individual files by selecting them and clicking "Remove", or clear all files with "Clear All"

---

### 2️⃣ Configure Reference Pattern

Navigate to the **Reference Pattern** tab to define how references appear in your PDFs.

#### 📝 Built-in Patterns

Choose from pre-configured patterns:

| Pattern | Example | Description |
|---------|---------|-------------|
| **Style 25-A.0** | `25-A.0`, `10-B.5` | Page-Row.Column format |
| **Style 5.3-A** | `5.3-A`, `12.7-C` | Page.Column-Row format |
| **Style /5.3-A** | `/5.3-A`, `/8.2-B` | With leading slash |
| **Style (5/A/3)** | `(5/A/3)`, `(12/C/7)` | Parentheses format |

#### 🎨 Custom Patterns

Create your own pattern using placeholders:

- `{P}` or `{PAG}` - Page number (e.g., 5, 25, 100)
- `{C}` or `{COL}` - Column number (e.g., 0, 3, 12)
- `{F}` or `{FILA}` - Row letter (e.g., A, B, Z)

**Example Custom Patterns:**

```
/{P}.{C}-{F}     →  /5.3-A
{P}-{F}.{C}      →  25-A.0
[{P}/{F}/{C}]    →  [5/A/3]
REF-{P}.{C}.{F}  →  REF-5.3.A
```

> ⚠️ **Important:** The pattern must match EXACTLY how references appear in your PDF!

---

### 3️⃣ Define Your Grid Layout

The grid defines how your diagram is divided into rows and columns.

#### 🎨 Visual Grid Editor (Recommended)

Click **"Grid Editor"** to open the visual editor:

1. **📄 Select Page** - Choose which page to use as reference
2. **🖱️ Click to Add Lines** - Click on the PDF to add column/row lines
3. **🗑️ Right-Click to Remove** - Right-click near a line to delete it
4. **🔍 Zoom & Pan** - Use mouse wheel to zoom, middle-click to pan
5. **💾 Save as Template** - Click "Save as Template" to reuse this grid

![Grid Editor](https://via.placeholder.com/800x400/4D1F0B/FF6B2B?text=Visual+Grid+Editor)

**Grid Editor Controls:**

| Control | Action |
|---------|--------|
| **Left Click** | Add column/row line |
| **Right Click** | Remove nearest line |
| **Mouse Wheel** | Zoom in/out |
| **Middle Click + Drag** | Pan around |
| **Page Selector** | Switch between pages |
| **Mode Toggle** | Switch between Column/Row mode |

#### 📐 Manual Grid Configuration

Alternatively, configure the grid manually in the **Grid Configuration** tab:

- **Columns** - Number of vertical divisions
- **Rows** - Number of horizontal divisions (A, B, C...)
- **Margins** - Left and top margins (%)
- **Column Widths** - Relative widths (e.g., `1,2,1` for narrow-wide-narrow)
- **Row Heights** - Relative heights (e.g., `1,1,1` for equal)

> 💡 **Pro Tip:** Use the visual editor first, then fine-tune with manual settings!

---

### 4️⃣ Customize Highlight Style

Make your links stand out with custom styling! Navigate to the **Highlight Style** tab.

#### 🎨 Basic Appearance

**Color Options:**
- 🔴 Red
- 🟢 Green  
- 🔵 Blue
- 🟡 Yellow
- 🟠 Orange
- 🟣 Magenta
- 🔵 Cyan

**Border Settings:**
- **Width** - Line thickness (1-10 px)
- **Style** - Solid, Dashed, or Dotted
- **Blink Speed** - Fast, Normal, Slow, or None

#### ✨ Animation & Effects

**Duration** - How long the highlight stays visible (0.1-30 seconds)

**Fill Style:**
- None - Transparent background
- Semi-transparent - Subtle fill
- Solid - Opaque fill

**Animation Type:**
- 🔄 Blink - Flash on/off
- 🌊 Fade - Smooth fade in/out
- 💓 Pulse - Pulsing effect
- ⏸️ None - Static highlight

**Opacity** - Transparency level (10-100%)

#### 🎯 Advanced Options

- **Fill Color** - Choose a different color for the fill
- **Corner Radius** - Rounded corners (0-20 px)
- **Margin** - Expand/shrink highlight area (-20 to +20 px)
- **Effect** - None, Soft shadow, or Glow

#### 👁️ Live Preview

See your changes in real-time with the preview box at the bottom!

![Style Preview](https://via.placeholder.com/200x80/EF4444/FFFFFF?text=REF-001)

---

### 5️⃣ Detect References

Ready to find those references? Let's go! 🚀

1. Click the **"🔍 Detect"** button
2. Watch the progress bar as the app analyzes your PDFs
3. View detected references in the table

**What You'll See:**

| Column | Description |
|--------|-------------|
| **Reference** | The detected reference text |
| **Page** | Page number |
| **Column** | Column number |
| **Row** | Row letter |
| **Context** | Surrounding text for verification |
| **PDF** | Source PDF (if multiple files) |

> 📊 **Statistics:** Click the **Statistics** button to see detailed analysis including:
> - Total references found
> - References per PDF
> - Distribution by page/column/row
> - Duplicate detection

---

### 6️⃣ Generate Interactive PDFs

Time to create your enhanced PDFs! 🎉

1. Click **"⚡ Generate"** button
2. Choose output location
3. Select options:
   - ✅ **Keep Original Name** - Preserve original filenames
   - ✅ **Disable Popups** - Suppress Adobe Reader warnings
   - ✅ **Clean PDF Links** - Remove existing links before adding new ones

4. Wait for generation to complete

**Success! 🎊**

Your new PDFs are ready with:
- ✅ Clickable cross-references
- ✅ Animated highlights
- ✅ Custom styling
- ✅ JavaScript-powered navigation

---

## 🎯 Advanced Features

### 💾 Grid Templates

Save time by creating reusable grid configurations!

**To Save a Template:**
1. Open the Visual Grid Editor
2. Configure your grid layout
3. Click **"💾 Save as Template"**
4. Enter a name (e.g., "Standard 10x6 Grid")

**To Load a Template:**
1. Click the grid template button in the sidebar (bottom left)
2. Select your saved template from the menu
3. Grid is instantly applied! ⚡

**To Manage Templates:**
1. Click the grid template button
2. Select **"⚙️ Manage Templates..."**
3. Delete unwanted templates

> 💡 **Use Case:** Create templates for different diagram types (control panels, schematics, layouts)

---

### 📊 Statistics Dashboard

Get insights into your reference detection:

**Metrics Displayed:**
- 📈 Total references detected
- 📄 References per PDF
- 🗺️ Distribution heatmap
- 🔢 Most common patterns
- ⚠️ Potential duplicates

**How to Access:**
1. Run detection first
2. Click the **📊 Statistics** button (dark red)
3. View comprehensive analysis

---

### 🎨 Style Presets

While the app doesn't have built-in presets, you can quickly recreate these popular styles:

**🔴 Classic Red Alert**
- Color: Red
- Width: 3px
- Style: Solid
- Animation: Blink (Fast)
- Duration: 3s

**🟢 Subtle Green**
- Color: Green
- Width: 2px
- Style: Dashed
- Fill: Semi-transparent
- Animation: Fade
- Duration: 5s

**🟡 High Visibility Yellow**
- Color: Yellow
- Width: 4px
- Style: Solid
- Fill: Semi-transparent
- Animation: Pulse
- Opacity: 80%
- Duration: 4s

---

## 🔧 Configuration Files

The app automatically saves your settings:

### 📁 File Locations

```
📂 Application Directory
├── 📄 grid_config.json          # Current grid configuration
├── 📄 grid_templates.json       # Saved grid templates
├── 📄 styles_config.json        # Highlight style settings
└── 📄 icons.json                # SVG icons (don't modify)
```

### 🔄 Backup & Restore

**To Backup Settings:**
1. Copy the JSON files above
2. Store them in a safe location

**To Restore Settings:**
1. Close the application
2. Replace JSON files with your backups
3. Restart the application

---

## 🎓 Tips & Best Practices

### ✅ Do's

- ✅ **Test with one PDF first** - Verify pattern and grid before batch processing
- ✅ **Use Visual Grid Editor** - More accurate than manual configuration
- ✅ **Save grid templates** - Reuse configurations for similar diagrams
- ✅ **Check statistics** - Verify detection accuracy before generating
- ✅ **Use descriptive template names** - "Panel_10x6", "Schematic_8x12", etc.
- ✅ **Adjust margins carefully** - Small changes can affect detection accuracy
- ✅ **Preview your style** - Use the live preview before generating

### ❌ Don'ts

- ❌ **Don't skip pattern verification** - Wrong pattern = no detection
- ❌ **Don't use extreme durations** - 0.1s too fast, 30s too long (3-5s is ideal)
- ❌ **Don't forget to save templates** - You'll thank yourself later!
- ❌ **Don't process scanned PDFs** - Text must be selectable, not images
- ❌ **Don't use too many grid lines** - More isn't always better
- ❌ **Don't ignore the context column** - Verify references are correct

---

## 🐛 Troubleshooting

### ❓ Common Issues

#### 🔴 No References Detected

**Possible Causes:**
- ❌ Pattern doesn't match your PDF format
- ❌ Grid configuration is incorrect
- ❌ PDF contains images instead of text

**Solutions:**
1. ✅ Verify pattern with "View References" after detection
2. ✅ Check if text is selectable in the PDF
3. ✅ Try the Visual Grid Editor for accurate grid placement
4. ✅ Use a custom pattern if built-in patterns don't work

---

#### 🔴 Links Don't Work in Generated PDF

**Possible Causes:**
- ❌ JavaScript disabled in PDF reader
- ❌ Using incompatible PDF viewer

**Solutions:**
1. ✅ Use Adobe Acrobat Reader DC (recommended)
2. ✅ Enable JavaScript in Reader:
   - Edit → Preferences → JavaScript
   - Check "Enable Acrobat JavaScript"
3. ✅ Try regenerating with "Clean PDF Links" enabled

---

#### 🔴 Highlights Don't Appear

**Possible Causes:**
- ❌ Duration too short
- ❌ Animation disabled
- ❌ Opacity too low

**Solutions:**
1. ✅ Increase duration to 3-5 seconds
2. ✅ Set animation to "Blink" or "Pulse"
3. ✅ Set opacity to 100%
4. ✅ Check preview to verify style

---

#### 🔴 Grid Lines Don't Match Diagram

**Possible Causes:**
- ❌ Wrong page selected as reference
- ❌ Margins not configured correctly
- ❌ PDF has different page sizes

**Solutions:**
1. ✅ Use Visual Grid Editor on the actual diagram page
2. ✅ Verify all pages have the same layout
3. ✅ Adjust margins in Grid Configuration tab
4. ✅ Check that column/row counts match your diagram

---

#### 🔴 Application Crashes or Freezes

**Possible Causes:**
- ❌ Very large PDF files
- ❌ Corrupted PDF
- ❌ Insufficient memory

**Solutions:**
1. ✅ Process PDFs one at a time
2. ✅ Close other applications to free memory
3. ✅ Try repairing the PDF with Adobe Acrobat
4. ✅ Check console for error messages

---

#### 🔴 Text Appears Cut Off in UI

**Possible Causes:**
- ❌ Window too small
- ❌ High DPI display scaling

**Solutions:**
1. ✅ Maximize the application window
2. ✅ Use scroll bars to view all content
3. ✅ Adjust Windows display scaling to 100%

---

## 🎨 Customization

### 🎭 Custom Icons

The app uses SVG icons defined in `icons.json`. You can customize them:

1. Open `icons.json`
2. Find the icon you want to change
3. Replace the SVG code with your own
4. Restart the application

**Available Icons:**
- `file_icon` - File/folder icon
- `pattern_icon` - Pattern configuration icon
- `style_icon` - Style customization icon
- `grid_editor` - Grid editor icon
- `stats_icon` - Statistics icon
- `pdf_icon` - PDF file icon in list

---

### 🎨 Color Scheme

Want to change the app's color scheme? Edit the `COLORS` dictionary in `main.py`:

```python
COLORS = {
    'accent': '#FF6B2B',  # Main accent color (orange)
    'accent_hover': '#FF8A4D',  # Hover state
    'accent_dim': '#4D1F0B',  # Dark accent
    # ... more colors
}
```

Popular color schemes:

**🔵 Blue Theme:**
```python
'accent': '#3B82F6',
'accent_hover': '#60A5FA',
'accent_dim': '#1E3A8A',
```

**🟣 Purple Theme:**
```python
'accent': '#8B5CF6',
'accent_hover': '#A78BFA',
'accent_dim': '#4C1D95',
```

**🟢 Green Theme:**
```python
'accent': '#10B981',
'accent_hover': '#34D399',
'accent_dim': '#065F46',
```

---

## 📚 Technical Details

### 🔧 How It Works

1. **PDF Analysis** - PyMuPDF extracts text and coordinates from each page
2. **Pattern Matching** - Regex patterns detect reference strings
3. **Grid Mapping** - References are mapped to grid coordinates
4. **Link Generation** - PyPDF2 creates clickable annotations
5. **JavaScript Injection** - Custom JS code handles animations
6. **PDF Assembly** - Modified PDF is saved with all enhancements

### 📦 Dependencies

| Library | Purpose |
|---------|---------|
| **PyQt5** | User interface framework |
| **PyMuPDF (fitz)** | PDF rendering and text extraction |
| **PyPDF2** | PDF manipulation and link creation |
| **Python 3.8+** | Core runtime |

### 🎯 JavaScript Implementation

The app injects JavaScript code into PDFs that:
- Creates highlight fields at target coordinates
- Implements animation effects (blink, pulse, fade)
- Handles timing and cleanup
- Automatically removes highlights after duration
- Works with Adobe Acrobat Reader and compatible viewers

**JavaScript Functions:**
- `highlight(page, coordinates)` - Main highlight function
- `blinker()` - Blink animation loop
- `finish()` - Cleanup and removal

---

## 🤝 Contributing

We welcome contributions! Here's how you can help:

### 🐛 Report Bugs

Found a bug? Please open an issue with:
- Detailed description
- Steps to reproduce
- Expected vs actual behavior
- Screenshots if applicable
- Your system info (OS, Python version)

### ✨ Suggest Features

Have an idea? Open an issue with:
- Feature description
- Use case / why it's needed
- Mockups or examples (if applicable)

### 💻 Submit Pull Requests

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'Add amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

---

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

---

## 🙏 Acknowledgments

- **PyQt5** - For the amazing GUI framework
- **PyMuPDF** - For powerful PDF processing capabilities
- **PyPDF2** - For PDF manipulation tools
- **Adobe** - For JavaScript API in PDF readers
- **Community** - For feedback and contributions

---

## 📞 Support

Need help? Here's how to get support:

- 📧 **Email:** support@example.com
- 💬 **Discord:** [Join our server](https://discord.gg/example)
- 🐛 **Issues:** [GitHub Issues](https://github.com/yourusername/pdf-reference-detector/issues)
- 📖 **Wiki:** [Documentation Wiki](https://github.com/yourusername/pdf-reference-detector/wiki)

---

## 🗺️ Roadmap

### 🚀 Upcoming Features

- [ ] **Multi-language support** - Spanish, French, German
- [ ] **Cloud sync** - Save templates to cloud
- [ ] **Batch templates** - Apply different templates to different PDFs
- [ ] **Export statistics** - CSV/Excel export
- [ ] **Dark mode** - Eye-friendly dark theme
- [ ] **Undo/Redo** - In visual grid editor
- [ ] **Auto-detection** - AI-powered grid detection
- [ ] **Mobile viewer** - Companion mobile app

### 🎯 Future Enhancements

- [ ] **Performance optimization** - Faster processing for large PDFs
- [ ] **Advanced patterns** - Support for complex reference formats
- [ ] **Collaboration** - Share templates with team
- [ ] **Version control** - Track changes to configurations
- [ ] **Plugins system** - Extend functionality with plugins

---

## 📊 Statistics

![GitHub stars](https://img.shields.io/github/stars/yourusername/pdf-reference-detector?style=social)
![GitHub forks](https://img.shields.io/github/forks/yourusername/pdf-reference-detector?style=social)
![GitHub watchers](https://img.shields.io/github/watchers/yourusername/pdf-reference-detector?style=social)

---

<div align="center">

### 🌟 Star this repo if you find it useful! 🌟

**Made with ❤️ by the PDF Reference Detector Team**

[⬆ Back to Top](#-pdf-reference-detector--link-generator)

</div>
