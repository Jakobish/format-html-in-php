# 🎨 Format HTML in Classic ASP

<div align="center">

**The Ultimate HTML Formatting Solution for Classic ASP Development**

[![Version](https://img.shields.io/badge/version-1.7.0-blue.svg)](https://marketplace.visualstudio.com/items?itemName=rifi2k.format-html-in-classic-asp)
[![VS Code](https://img.shields.io/badge/VS_Code-%5E1.48.0-blue.svg)](https://code.visualstudio.com/)
[![License](https://img.shields.io/badge/license-Unlicense-green.svg)](LICENSE.md)
[![Downloads](https://img.shields.io/badge/downloads-10K+-brightgreen.svg)](https://marketplace.visualstudio.com/items?itemName=rifi2k.format-html-in-classic-asp)

*Format HTML code in Classic ASP files with precision and ease*

[📦 Install from VS Code Marketplace](https://marketplace.visualstudio.com/items?itemName=rifi2k.format-html-in-classic-asp) • [🐛 Report Issues](https://github.com/RiFi2k/format-html-in-classic-asp/issues) • [💡 Request Features](https://github.com/RiFi2k/format-html-in-classic-asp/issues)

</div>

---

## 📋 Table of Contents

- [✨ Features](#-features)
- [🚀 Quick Start](#-quick-start)
- [📖 Usage](#-usage)
- [⚙️ Configuration](#️-configuration)
- [🎯 Supported File Types](#-supported-file-types)
- [🛠️ Commands](#-commands)
- [🗺️ Roadmap](#-roadmap)
- [🤝 Contributing](#-contributing)
- [📄 License](#-license)

---

## ✨ Features

### 🎯 Core Functionality

- **Smart HTML Formatting**: Uses VS Code's native HTML formatting settings
- **Classic ASP-Compatible**: Preserves `<% %>` and `<%= %>` server-side code blocks
- **Format on Save**: Automatic formatting when saving Classic ASP files
- **Multi-Language Support**: Works with Classic ASP, VBScript, and JavaScript

### 🚀 Advanced Features

- **Selection Formatting**: Format only selected HTML code
- **Batch Processing**: Format all open Classic ASP files simultaneously
- **Error Handling**: Comprehensive error reporting and user feedback
- **Performance Optimized**: Efficient processing for large files
- **Configurable**: Extensive customization options

### 🎨 Formatting Capabilities

- ✅ Nested HTML structures
- ✅ Mixed HTML/Classic ASP code
- ✅ Server-side script blocks preservation
- ✅ VS Code HTML settings integration
- ✅ Custom formatting rules support

---

## 🚀 Quick Start

### Installation

1. **Via VS Code Marketplace** (Recommended):
   - Open VS Code
   - Go to Extensions (`Ctrl+Shift+X`)
   - Search for "Format HTML in Classic ASP"
   - Click Install

2. **Via Command Line**:

    ```bash
    code --install-extension rifi2k.format-html-in-classic-asp
    ```

### Basic Usage

1. **Open a Classic ASP file** (`.asp`, `.asa`, `.inc`)
2. **Format the entire file**:
   - Press `Ctrl+Alt+F`
   - Or use Command Palette: `Format HTML in Classic ASP`
   - Or right-click and select `Format HTML in Classic ASP`

3. **Format selected text only**:
   - Select HTML code
   - Use Command Palette: `Format HTML in Classic ASP (Selection)`

---

## 📖 Usage

### Format on Save

Enable automatic formatting when saving:

```json
{
  "editor.formatOnSave": false,
  "[asp]": {
    "editor.formatOnSave": true
  }
}
```

### Keyboard Shortcuts

| Command | Shortcut | Description |
|---------|----------|-------------|
| Format Document | `Ctrl+Alt+F` | Format entire Classic ASP file |
| Format Selection | `Ctrl+Alt+F` | Format selected text only |

### Context Menu

Right-click in any Classic ASP file to access formatting options:

- **Format HTML in Classic ASP**: Format entire file
- **Format HTML in Classic ASP (Selection)**: Format selected text (when text is selected)

### Command Palette

Access all features via `Ctrl+Shift+P`:

- `Format HTML in Classic ASP`: Format current file
- `Format HTML in Classic ASP (Selection)`: Format selection
- `Format HTML in All Open Classic ASP Files`: Batch format all open files

---

## ⚙️ Configuration

### Extension Settings

Configure the extension behavior:

```json
{
  // Enable/disable the extension
  "formatHtmlInAsp.enable": true,

  // Format on save for Classic ASP files
  "formatHtmlInAsp.formatOnSave": true,

  // Preserve Classic ASP server blocks during formatting
  "formatHtmlInAsp.preserveServerBlocks": true,

  // Supported language IDs
  "formatHtmlInAsp.supportedLanguages": [
    "asp",
    "vbscript",
    "javascript"
  ]
}
```

### HTML Formatting Settings

Leverage VS Code's native HTML formatting options:

```json
{
  "html.format.enable": true,
  "html.format.wrapLineLength": 120,
  "html.format.wrapAttributes": "auto",
  "html.format.indentInnerHtml": false,
  "html.format.preserveNewLines": true,
  "html.format.maxPreserveNewLines": 2,
  "html.format.endWithNewline": false,
  "html.format.extraLiners": "head, body, /html",
  "html.format.contentUnformatted": "pre,code,textarea"
}
```

---

## 🎯 Supported File Types

The extension intelligently detects and formats HTML in:

| File Type | Extension | Description |
|-----------|-----------|-------------|
| Classic ASP | `.asp` | Active Server Pages |
| ASP Application | `.asa` | Global.asa files |
| Include Files | `.inc` | Classic ASP include files |
| VBScript | `.asp` | VBScript in HTML |
| JavaScript | `.asp` | JavaScript in HTML |

### Language Detection

Automatically activates for:

- `asp` - Classic ASP files
- `vbscript` - VBScript embedded content
- `javascript` - JavaScript embedded content

---

## 🛠️ Commands

### Available Commands

| Command | Description | Context |
|---------|-------------|---------|
| `formatHtmlInAsp.format` | Format entire Classic ASP file | Classic ASP files |
| `formatHtmlInAsp.formatSelection` | Format selected text | Classic ASP files with selection |
| `formatHtmlInAsp.formatAllOpen` | Format all open Classic ASP files | Always available |

### Command Integration

- **Command Palette**: All commands available via `Ctrl+Shift+P`
- **Context Menu**: File and selection commands in right-click menu
- **Keyboard Shortcuts**: Quick access via customizable shortcuts

---

## 🗺️ Roadmap

### 🚀 Planned Features (Q1 2024)

#### Phase 1: Enhanced Formatting

- [ ] **Syntax Highlighting**: Improved ASP syntax highlighting
- [ ] **IntelliSense**: ASP-specific code completion
- [ ] **Code Folding**: Smart folding for ASP blocks
- [ ] **Bracket Matching**: Enhanced bracket matching for ASP tags

#### Phase 2: Advanced Features

- [ ] **Code Snippets**: Pre-built ASP code snippets
- [ ] **Refactoring Tools**: ASP-specific refactoring
- [ ] **Debug Integration**: Enhanced debugging support
- [ ] **Project Templates**: ASP project templates

#### Phase 3: Enterprise Features

- [ ] **Multi-root Workspace**: Support for complex ASP projects
- [ ] **Task Integration**: Build and deployment tasks
- [ ] **Testing Framework**: Unit testing support
- [ ] **Performance Monitoring**: Code performance analysis

### 🎯 Must-Have Features for Best ASP Extension

#### Core Requirements

- [x] HTML formatting in ASP files
- [x] Multiple file type support
- [x] Format on save
- [x] Selection formatting
- [x] Batch processing
- [x] Error handling

#### Advanced Requirements

- [ ] **ASP Code Formatting**: Format VBScript/JavaScript in ASP files
- [ ] **Database Integration**: SQL formatting in ASP
- [ ] **Component Recognition**: ASP component and COM object support
- [ ] **Session/State Management**: ASP session and application variable handling
- [ ] **Security Analysis**: ASP security vulnerability detection
- [ ] **Performance Optimization**: ASP performance suggestions

#### Developer Experience

- [ ] **Live Preview**: Real-time ASP page preview
- [ ] **Code Navigation**: Go to definition for ASP includes
- [ ] **Refactoring**: Rename variables across ASP files
- [ ] **Documentation**: Auto-generate ASP documentation
- [ ] **Version Control**: ASP-specific Git integration

### 📊 Feature Priority Matrix

| Feature Category | Current Status | Priority | Timeline |
|------------------|----------------|----------|----------|
| HTML Formatting | ✅ Complete | Critical | Done |
| Multi-file Support | ✅ Complete | High | Done |
| Error Handling | ✅ Complete | High | Done |
| ASP Code Formatting | ❌ Planned | Critical | Q2 2024 |
| IntelliSense | ❌ Planned | High | Q1 2024 |
| Debug Integration | ❌ Planned | Medium | Q2 2024 |
| Testing Framework | ❌ Planned | Low | Q3 2024 |

---

## 🤝 Contributing

We welcome contributions! Here's how you can help:

### Development Setup

1. **Clone the repository**:

   ```bash
   git clone https://github.com/RiFi2k/format-html-in-classic-asp.git
   cd format-html-in-classic-asp
   ```

2. **Install dependencies**:

   ```bash
   npm install
   ```

3. **Compile the extension**:

   ```bash
   npm run compile
   ```

4. **Test in VS Code**:
   - Open in VS Code
   - Press `F5` to launch extension development host

### Contribution Guidelines

- **Issues**: Report bugs and request features on [GitHub Issues](https://github.com/RiFi2k/format-html-in-classic-asp/issues)
- **Pull Requests**: Submit PRs with clear descriptions
- **Code Style**: Follow TypeScript best practices
- **Testing**: Include tests for new features

### Development Commands

```bash
# Compile TypeScript
npm run compile

# Watch for changes
npm run watch

# Lint code
npm run lint

# Run tests
npm test
```

---

## 📄 License

This project is licensed under the **Unlicense** - see the [LICENSE.md](LICENSE.md) file for details.

### Acknowledgments

- **JS Beautify**: HTML formatting engine
- **VS Code**: Extension platform
- **ASP Community**: Inspiration and feedback

---

<div align="center">

**Made with ❤️ for the Classic ASP development community**

[⭐ Star us on GitHub](https://github.com/RiFi2k/format-html-in-classic-asp) • [🐛 Report Issues](https://github.com/RiFi2k/format-html-in-classic-asp/issues) • [💡 Feature Requests](https://github.com/RiFi2k/format-html-in-classic-asp/issues)

</div>
