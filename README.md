# Word Macro Toolkit

A collection of useful VBA macros for Microsoft Word to enhance document formatting efficiency.

## Contents

1. **BlueFormatter.bas** - One-click blue color formatting for academic documents
2. **ThreeLineTable.bas** - Convert tables to professional three-line style

## BlueFormatter

### Features
- Recolor cross-references, headings, captions, abstract, and keywords to blue with one click
- Supports both English and Chinese style names
- Handles:
  - Built-in styles (Heading 1-9, Caption, Hyperlink, TOC 1-9)
  - Custom styles (Abstract, Keywords, 摘要, 关键词, etc.)
  - Cross-reference fields (Ref, PageRef, NoteRef, Hyperlink, etc.)
  - Text inside shapes and text boxes

### Usage

1. Open the VBA editor in Word (`Alt+F11`)
2. Go to `File` → `Import File` → Select `BlueFormatter.bas`
3. Run the macro `ChangeToBlue`

### Diagnostic Tool

Run `ListUsedStyles` to see all style names used in your document. This helps identify any missing styles that need to be added to the custom names list.

## ThreeLineTable

### Features
- **SelectAllTables**: Selects all tables in the document at once
- **ConvertToThreeLineTables**: Converts all tables to the professional three-line style format
  - Thick top border (1.5pt)
  - Thick bottom border (1.5pt)
  - Thin header separator (0.75pt)
  - Clean, minimalist design

### Usage

1. Open the VBA editor in Word (`Alt+F11`)
2. Insert a new module or import `ThreeLineTable.bas`
3. Run either:
   - `SelectAllTables` to select all tables
   - `ConvertToThreeLineTables` to format all tables

## Installation

### Method 1: Import Module
1. Open Word
2. Press `Alt+F11` to open VBA editor
3. Right-click your document in the Project Explorer
4. Select `Import File` and choose the `.bas` file

### Method 2: Manual Paste
1. Open Word
2. Press `Alt+F11` to open VBA editor
3. Insert a new module (`Insert` → `Module`)
4. Copy and paste the code from the `.bas` file
5. Save the document as a macro-enabled file (`.docm`)

## Notes

- These macros work with Microsoft Word 2010 and later versions
- Always save your document before running macros
- It's recommended to test macros on a copy of your document first
- Macros may need to be enabled in Word's security settings

## License

This project is in the public domain. Feel free to use and modify as needed.