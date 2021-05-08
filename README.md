Now you can ditch MS Word and make your own Word by 300 lines of python code it has very much functions like the MS Word!!
I have done it by using Pycharm which is an python code editor.
In the terminal first I have installed a pip(python installation package) called **pip install PyQt5**.
In the first I have imported sys which is deafult there in Pycharm.
Then from PyQt5 I have imported QtGui, QtWidgets, QtCore and QtPrintSupport I have imported them fully so for that I have written from PyQt5.QtGui import * from PyQt5.QtWidgets import * from PyQt5.QtCore import * from PyQt5.QtPrintSupport import *.
Then I have imported docx2txt this is also default in Pycharm.
Then I have written class MainApp(QMainWindow): amd inside this I have written def __init__(self): and inside this I have written super().__init__() then for the title action, editor action, to create the menu bar, to create the tool bar and font action I have written the code as self.title = "THE NAME YOU WANT TO DISPLAY ON YOU WORD ON THE LEFT TOP CORNER" self.setWindowTitle(self.title) self.editor = QTextEdit(self) self.setCentralWidget(self.editor) self.create_menu_bar() self.create_toolbar() font = QFont('Times', 12) self.editor.setFont(font) self.editor.setFontPointSize(12) than after this I have written self.path = ''.
Then I have created a function called def create_menu_bar(self): and inside this function I have created a variable called menuBar and entered the value as QMenuBar(self) and after this I have created another variable called file_menu and entered the value as QMenu("File", self) and then to add it in the menuBar I have written menuBar.addMenu(file_menu) then again for the save action, open action, rename action, pdf action, edit action, paste action, clear action, select action, view Menu action full screen action, normal screen action, minimize screen action for all of these I have written the code as file_menu = QMenu("File", self) menuBar.addMenu(file_menu) save_action = QAction('Save', self) save_action.triggered.connect(self.file_save) file_menu.addAction(save_action) open_action = QAction('Open', self) open_action.triggered.connect(self.file_open) file_menu.addAction(open_action) rename_action = QAction('Rename', self) rename_action.triggered.connect(self.file_saveas) file_menu.addAction(rename_action) pdf_action = QAction("Save as PDF", self) pdf_action.triggered.connect(self.save_pdf) file_menu.addAction(pdf_action) edit_menu = QMenu("Edit", self) menuBar.addMenu(edit_menu) paste_action = QAction('Paste', self) paste_action.triggered.connect(self.editor.paste) edit_menu.addAction(paste_action) clear_action = QAction('Clear', self) clear_action.triggered.connect(self.editor.clear) edit_menu.addAction(clear_action) select_action = QAction('Select All', self) select_action.triggered.connect(self.editor.selectAll) edit_menu.addAction(select_action) view_menu = QMenu("View", self) menuBar.addMenu(view_menu) fullscr_action = QAction('Full Screen View', self) fullscr_action.triggered.connect(lambda: self.showFullScreen()) view_menu.addAction(fullscr_action) normscr_action = QAction('Normal View', self) normscr_action.triggered.connect(lambda: self.showNormal()) view_menu.addAction(normscr_action) minscr_action = QAction('Minimize', self) minscr_action.triggered.connect(lambda: self.showMinimized() view_menu.addAction(minscr_action). then I have written self.setMenuBar(menuBar).
Then I have created a function called def create_toolbar(self): and inside this function I have created a variable callled ToolBar and entered the value as QToolBar("Tools", self)Then after that undo action, redo action, copy action, cut action, paste action, set font action, set fon size action, bold action, underline action, italic action, right alignment action, left alignment action, justification action, zoom in action, zoom out action for all these actions I have given them a space and how to give them a space is just write ToolBar.addSeparator() after each action I have given four of thm each and these all actions I have written for triggered example if the bold action is triggered then what to do is written in the actions and the code for all these triggered actions with the seperators is undo_action = QAction(QIcon("undo.png"), 'Undo', self) undo_action.triggered.connect(self.editor.undo) ToolBar.addAction(undo_action) ToolBar.addSeparator() ToolBar.addSeparator() ToolBar.addSeparator() ToolBar.addSeparator() redo_action = QAction(QIcon("redo.png"), 'Redo', self) redo_action.triggered.connect(self.editor.redo) ToolBar.addAction(redo_action) ToolBar.addSeparator() ToolBar.addSeparator() ToolBar.addSeparator() ToolBar.addSeparator() copy_action = QAction(QIcon("copy.png"), 'Copy', self) copy_action.triggered.connect(self.editor.copy) ToolBar.addAction(copy_action) ToolBar.addSeparator() ToolBar.addSeparator() ToolBar.addSeparator() ToolBar.addSeparator() cut_action = QAction(QIcon("cut.png"), 'Cut', self) cut_action.triggered.connect(self.editor.cut) ToolBar.addAction(cut_action) ToolBar.addSeparator() ToolBar.addSeparator() ToolBar.addSeparator() ToolBar.addSeparator() paste_action = QAction(QIcon("paste.png"), 'Paste', self) paste_action.triggered.connect(self.editor.paste) ToolBar.addAction(paste_action) ToolBar.addSeparator() ToolBar.addSeparator() ToolBar.addSeparator() ToolBar.addSeparator() self.font_combo = QComboBox(self) self.font_combo.addItems() ["Courier Std", "Hellentic Typewriter Regular", "Helvetica", "Arial", "SansSerif", "Helvetica", "Times", "Monospace"]) self.font_combo.activated.connect(self.set_font)  # connect with function toolBar.addWidget(self.font_combo)toolBar.addSeparator() ToolBar.addSeparator()  ToolBar.addSeparator() ToolBar.addSeparator() self.font_size = QSpinBox(self) self.font_size.setValue(12) self.font_size.valueChanged.connect(self.set_font_size)  # connect with funcion ToolBar.addWidget(self.font_size) toolBar.addSeparator() ToolBar.addSeparator() ToolBar.addSeparator() ToolBar.addSeparator() bold_action = QAction(QIcon("bold.png"), 'Bold', self)
        bold_action.triggered.connect(self.bold_text)
        ToolBar.addAction(bold_action)
        ToolBar.addSeparator()
        ToolBar.addSeparator()
        ToolBar.addSeparator()
        ToolBar.addSeparator()
        underline_action = QAction(QIcon("underline.png"), 'Underline', self)
        underline_action.triggered.connect(self.underline_text)
        ToolBar.addAction(underline_action)
        ToolBar.addSeparator()
        ToolBar.addSeparator()
        ToolBar.addSeparator()
        ToolBar.addSeparator()
        italic_action = QAction(QIcon("italic.png"), 'Italic', self)
        italic_action.triggered.connect(self.italic_text)
        ToolBar.addAction(italic_action)
        ToolBar.addSeparator()
        ToolBar.addSeparator()
        ToolBar.addSeparator()
        ToolBar.addSeparator()
        right_alignment_action = QAction(QIcon("right-align.png"), 'Align Right', self)
        right_alignment_action.triggered.connect(lambda: self.editor.setAlignment(Qt.AlignRight))
        ToolBar.addAction(right_alignment_action)
        ToolBar.addSeparator()
        ToolBar.addSeparator()
        ToolBar.addSeparator()
        ToolBar.addSeparator()
        left_alignment_action = QAction(QIcon("left-align.png"), 'Align Left', self)
        left_alignment_action.triggered.connect(lambda: self.editor.setAlignment(Qt.AlignLeft))
        ToolBar.addAction(left_alignment_action)
        ToolBar.addSeparator()
        ToolBar.addSeparator()
        ToolBar.addSeparator()
        ToolBar.addSeparator()
        justification_action = QAction(QIcon("justification.png"), 'Center/Justify', self)
        justification_action.triggered.connect(lambda: self.editor.setAlignment(Qt.AlignCenter))
        ToolBar.addAction(justification_action)
        ToolBar.addSeparator()
        ToolBar.addSeparator()
        ToolBar.addSeparator()
        ToolBar.addSeparator()
        zoom_in_action = QAction(QIcon("zoom-in.png"), 'Zoom in', self)
        zoom_in_action.triggered.connect(self.editor.zoomIn)
        ToolBar.addAction(zoom_in_action)
        ToolBar.addSeparator()
        ToolBar.addSeparator()
        ToolBar.addSeparator()
        ToolBar.addSeparator()
        zoom_out_action = QAction(QIcon("zoom-out.png"), 'Zoom out', self)
        zoom_out_action.triggered.connect(self.editor.zoomOut)
        ToolBar.addAction(zoom_out_action)
        ToolBar.addSeparator()
        ToolBar.addSeparator()
        ToolBar.addSeparator()
        ToolBar.addSeparator()
Then after all this I have written self.addToolBar(ToolBar).
And then for all these actions which are coded to trigger they all have a function created and the code is     def italic_text(self):
            state = self.editor.fontItalic()
            self.editor.setFontItalic(not (state))
    def underline_text(self):
            state = self.editor.fontUnderline()
            self.editor.setFontUnderline(not (state))
    def bold_text(self):
            if self.editor.fontWeight() != QFont.Bold:
                self.editor.setFontWeight(QFont.Bold)
                return
            self.editor.setFontWeight(QFont.Normal)
    def set_font(self):
        font = self.font_combo.currentText()
        self.editor.setCurrentFont(QFont(font))
    def set_font_size(self):
        value = self.font_size.value()
        self.editor.setFontPointSize(value)
    def file_open(self):
        self.path, _ = QFileDialog.getOpenFileName(self, "Open file", "",
                                                   "Text documents (*.text);Text documents (*.txt);All files (*.*)")
        self.editor.setText(text)
        self.update_title()
    def file_save(self):
        print(self.path)
        if self.path == '':
            self.file_saveas()
        text = self.editor.toPlainText()
    def file_saveas(self):
        self.path, _ = QFileDialog.getSaveFileName(self, "Save file", "",
                                                   "text documents (*.text);Text documents (*.txt);All files (*.*)")
        if self.path == '':
            return
        text = self.editor.toPlainText()
    def update_title(self):
        self.setWindowTitle(self.title + ' ' + self.path)
    def save_pdf(self):
        f_name, _ = QFileDialog.getSaveFileName(self, "Export PDF", None, "PDF files (.pdf);;All files()")
        print(f_name).
        Then I have created a if condition called if f_name != '': and inside this I have created a variable called printer and entered th value as QPrinter(QPrinter.HighResolution) and then I have written printer.setOutputFormat(QPrinter.PdfFormat)
            printer.setOutputFileName(f_name)
            self.editor.document().print_(printer).
            And then to run the app and to show the window and to execute the app I have written app = QApplication(sys.argv)
window = MainApp()
window.show()
sys.exit(app.exec_()).
Now you are ready with your own Word now you can ditch MS Word!! Enjoy!!
