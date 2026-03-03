# Quill
**Fantastic Web WordPad**

Quill is a modern, production-ready rich text editor built with pure HTML, CSS, and Vanilla JavaScript.

It delivers a refined, user-friendly interface with advanced editing tools, theme support, autosave, zoom control, file export, and more, all running fully in the browser.

---

## Features

### Rich Text Editing
- Bold, Italic, Underline, Strikethrough
- Text alignment (Left, Center, Right, Justify)
- Ordered & Unordered lists
- Insert links, images, tables
- Line spacing control
- Clear formatting

### Advanced Font System
- 50+ Google Fonts integrated
- Live preview font picker
- Font size selection
- Text color & highlight color

### Theme System
- Light mode
- Dark mode
- Sepia mode
- System theme detection

### Storage & Autosave
- Automatic local autosave
- Save documents to browser storage
- Recent documents list
- Word & character count
- Zoom persistence

### File Operations
- New document
- Open `.html` and `.txt` files
- Save as HTML
- Save as TXT
- Export as PDF (via browser print)

### Productivity Tools
- Find & Replace with live highlighting
- Floating contextual toolbar
- Keyboard shortcuts
- Undo / Redo
- Page breaks
- Modal-based insert dialogs

---

## Tech Stack

- **HTML** – Semantic structure
- **CSS** – Design system & theming
- **Vanilla JavaScript (ES6+)** – Modular architecture
- **LocalStorage API** – Persistence

No frameworks. No dependencies. Fully standalone.

---

## Project Structure

```
.
├── index.html     # Main application UI
├── styles.css     # Design system & themes
└── script.js      # Editor core logic
```

---

## Architecture Overview

The editor is modularly structured inside `script.js`:

- ThemeManager
- ZoomManager
- Storage (Autosave & Recent Docs)
- Editor Core
- Find & Replace Engine
- File Operations
- Modal System
- Floating Toolbar
- Insert Utilities

State is centrally managed and persisted where necessary.


---

## License

Open-source —> free to modify and extend.
