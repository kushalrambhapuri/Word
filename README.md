# Create Your Own Word Processor
Now you can ditch MS Word and build your own word processor with just 300 lines of Python code! Packed with essential features, it provides functionality similar to MS Word.
This project was developed using PyCharm, a popular Python code editor. Follow these steps to create your own custom word processor:

Prerequisites
Python Installed: Make sure Python is installed on your system.
Install Required Libraries: Run the following command in the terminal:
pip install PyQt5 docx2txt  
Key Features
File Management: Open, save, rename, and export files as PDFs.
Editing Tools: Undo, redo, copy, cut, paste, and clear text.
Text Formatting: Bold, italicize, underline, and align text (left, right, center/justify).
Font Options: Choose fonts, adjust font size, and apply formatting.
Screen Views: Full-screen, normal view, and minimize options.
Zoom Options: Zoom in and out for better readability.
How It Works
Step 1: Setup
Import Modules:

Import the necessary modules, including sys, PyQt5 components (QtGui, QtWidgets, QtCore, QtPrintSupport), and docx2txt.
Create the Main Application Class:

Use the QMainWindow class from PyQt5 to design the window.
Example code for initialization:

class MainApp(QMainWindow):  
    def __init__(self):  
        super().__init__()  
        self.title = "Your Word Processor"  
        self.setWindowTitle(self.title)  
        self.editor = QTextEdit(self)  
        self.setCentralWidget(self.editor)  
        self.create_menu_bar()  
        self.create_toolbar()  
        self.path = ''  
Step 2: Menu Bar and Toolbar
Menu Bar:

Add menus for file, edit, and view operations.
Example for the File menu:

def create_menu_bar(self):  
    menuBar = QMenuBar(self)  
    file_menu = QMenu("File", self)  
    menuBar.addMenu(file_menu)  

    # Add Save, Open, Rename, and PDF export actions  
    save_action = QAction('Save', self)  
    save_action.triggered.connect(self.file_save)  
    file_menu.addAction(save_action)  

    open_action = QAction('Open', self)  
    open_action.triggered.connect(self.file_open)  
    file_menu.addAction(open_action)  

    pdf_action = QAction("Save as PDF", self)  
    pdf_action.triggered.connect(self.save_pdf)  
    file_menu.addAction(pdf_action)  

    self.setMenuBar(menuBar)  
Toolbar:

Add undo, redo, copy, paste, and text formatting actions.
Use icons to make the toolbar user-friendly.
Example for Undo and Redo actions:

def create_toolbar(self):  
    ToolBar = QToolBar("Tools", self)  

    undo_action = QAction(QIcon("undo.png"), 'Undo', self)  
    undo_action.triggered.connect(self.editor.undo)  
    ToolBar.addAction(undo_action)  

    redo_action = QAction(QIcon("redo.png"), 'Redo', self)  
    redo_action.triggered.connect(self.editor.redo)  
    ToolBar.addAction(redo_action)  

    self.addToolBar(ToolBar)  
Font Settings:

Add a combo box for font selection and a spin box for font size.
Example:

self.font_combo = QComboBox(self)  
self.font_combo.addItems(["Courier Std", "Helvetica", "Arial", "Times", "Monospace"])  
self.font_combo.activated.connect(self.set_font)  
ToolBar.addWidget(self.font_combo)  

self.font_size = QSpinBox(self)  
self.font_size.setValue(12)  
self.font_size.valueChanged.connect(self.set_font_size)  
ToolBar.addWidget(self.font_size)  
Step 3: Define Actions and Functions
Text Formatting Functions:
Example for Bold text:

def bold_text(self):  
    if self.editor.fontWeight() != QFont.Bold:  
        self.editor.setFontWeight(QFont.Bold)  
    else:  
        self.editor.setFontWeight(QFont.Normal)  
File Operations:
Example for Open File:

def file_open(self):  
    self.path, _ = QFileDialog.getOpenFileName(self, "Open file", "", "Text documents (*.txt);;All files (*.*)")  
    if self.path:  
        with open(self.path, 'r') as file:  
            text = file.read()  
            self.editor.setText(text)  
PDF Export:
Example:

def save_pdf(self):  
    f_name, _ = QFileDialog.getSaveFileName(self, "Export PDF", None, "PDF files (*.pdf);;All files ()")  
    if f_name:  
        printer = QPrinter(QPrinter.HighResolution)  
        printer.setOutputFormat(QPrinter.PdfFormat)  
        printer.setOutputFileName(f_name)  
        self.editor.document().print_(printer)  
Step 4: Run the Application
At the end of the script, initialize and execute the application:

if __name__ == "__main__":  
   app = QApplication(sys.argv)  
   window = MainApp()  
   window.show()  
   sys.exit(app.exec_())  
   
Ready to Use!
Now you have your own fully functional word processor. Customize it further, or enjoy using it as is. Say goodbye to MS Word and hello to your personal Word alternative!
