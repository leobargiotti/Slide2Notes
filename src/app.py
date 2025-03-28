import sys
from PyQt6.QtCore import Qt
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout,
                             QPushButton, QFileDialog, QComboBox, QLabel, QProgressBar,
                             QListWidget, QHBoxLayout, QCheckBox, QMessageBox, QProgressDialog)
import os
from dotenv import load_dotenv
from languages import TRANSLATIONS
from utils import (send_request_to_api, save_as_docx_file, extract_text_and_images_from_pptx, extract_text_from_pdf, extract_text_from_pptx,
                    save_as_pdf_file, create_summary_prompt, extract_text_and_images_from_pdf)


class DocumentSummaryApp(QMainWindow):
    """
    A PyQt6 application for summarizing documents and saving the summaries in DOCX or PDF format.

    Attributes:
    - input_files (list): List of selected input files.
    - output_language (str): Language for the output summary.
    - current_language (str): Current language of the UI.
    - save_as_docx (bool): Flag to save the summary as a DOCX file.
    - save_as_pdf (bool): Flag to save the summary as a PDF file.
    """

    def __init__(self):
        """
        Initialize the DocumentSummaryApp.
        """
        super().__init__()
        self.input_files = []
        self.output_language = "Italian"  # Default output language
        self.current_language = "Italiano"  # Track current UI language
        self.save_as_docx = True
        self.save_as_pdf = False
        self.extract_images = False
        self.init_ui()

    def init_ui(self):
        """
        Initialize the user interface.
        """
        self.setWindowTitle('Document to Word Summary')
        self.setGeometry(100, 100, 800, 600)  # Increased window size for file list

        # Main widget and layout
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        layout = QVBoxLayout()

        # Language selection for UI
        self.ui_language_label = QLabel("Interface Language:")
        self.ui_language_combo = QComboBox()
        self.ui_language_combo.addItems(["English", "Français", "Español", "Italiano"])
        self.ui_language_combo.setCurrentText("Italiano")
        self.ui_language_combo.currentTextChanged.connect(self.change_ui_language)

        # Output language selection
        self.output_language_label = QLabel("Output Summary Language:")
        self.output_language_combo = QComboBox()
        self.output_language_combo.addItems(["English", "French", "Spanish", "Italian"])
        self.output_language_combo.setCurrentText("Italian")
        self.output_language_combo.currentTextChanged.connect(self.set_output_language)

        # Output format options
        self.output_format_label = QLabel("Output Format:")
        format_layout = QHBoxLayout()

        self.docx_checkbox = QCheckBox("DOCX")
        self.docx_checkbox.setChecked(True)
        self.docx_checkbox.stateChanged.connect(self.toggle_docx)

        self.pdf_checkbox = QCheckBox("PDF")
        self.pdf_checkbox.setChecked(False)
        self.pdf_checkbox.stateChanged.connect(self.toggle_pdf)

        format_layout.addWidget(self.docx_checkbox)
        format_layout.addWidget(self.pdf_checkbox)
        format_layout.addStretch()

        # Image extraction option
        self.image_extraction_label = QLabel("Content Options:")
        image_layout = QHBoxLayout()
        self.image_checkbox = QCheckBox("Include Images")
        self.image_checkbox.setChecked(False)
        self.image_checkbox.stateChanged.connect(self.toggle_image_extraction)  # type: ignore
        image_layout.addWidget(self.image_checkbox)
        image_layout.addStretch()

        # Files list section
        self.files_label = QLabel("Selected files:")
        self.files_list = QListWidget()

        # File management buttons
        file_buttons_layout = QHBoxLayout()
        self.move_up_btn = QPushButton("Move Up")
        self.move_down_btn = QPushButton("Move Down")
        self.remove_btn = QPushButton("Remove")

        self.move_up_btn.clicked.connect(self.move_item_up)
        self.move_down_btn.clicked.connect(self.move_item_down)
        self.remove_btn.clicked.connect(self.remove_selected_file)

        file_buttons_layout.addWidget(self.move_up_btn)
        file_buttons_layout.addWidget(self.move_down_btn)
        file_buttons_layout.addWidget(self.remove_btn)

        # Main buttons
        self.select_files_btn = QPushButton('Select Files')
        self.select_files_btn.clicked.connect(self.select_files)

        self.process_btn = QPushButton('Generate Summary')
        self.process_btn.clicked.connect(self.process_files)
        self.process_btn.setEnabled(False)

        # Progress bar and status
        self.status_label = QLabel('')

        # Add widgets to layout
        layout.addWidget(self.ui_language_label)
        layout.addWidget(self.ui_language_combo)
        layout.addWidget(self.output_language_label)
        layout.addWidget(self.output_language_combo)
        layout.addWidget(self.output_format_label)
        layout.addLayout(format_layout)
        layout.addWidget(self.image_extraction_label)  # Moved here
        layout.addLayout(image_layout)
        layout.addWidget(self.select_files_btn)
        layout.addWidget(self.files_label)
        layout.addWidget(self.files_list)
        layout.addLayout(file_buttons_layout)
        layout.addWidget(self.process_btn)
        layout.addWidget(self.status_label)

        main_widget.setLayout(layout)

        # Initialize UI in Italian
        self.change_ui_language("Italiano")

    def toggle_docx(self, state):
        """
        Toggle the DOCX output format.

        Parameters:
        - state (int): The state of the checkbox (0 or 2).
        """
        self.save_as_docx = state > 0
        # Ensure at least one format is selected
        if not self.save_as_docx and not self.save_as_pdf:
            self.pdf_checkbox.setChecked(True)
            self.save_as_pdf = True

    def toggle_pdf(self, state):
        """
        Toggle the PDF output format.

        Parameters:
        - state (int): The state of the checkbox (0 or 2).
        """
        self.save_as_pdf = state > 0
        # Ensure at least one format is selected
        if not self.save_as_docx and not self.save_as_pdf:
            self.docx_checkbox.setChecked(True)
            self.save_as_docx = True

    def move_item_up(self):
        """
        Move the selected file up in the list.
        """
        current_row = self.files_list.currentRow()
        if current_row > 0:
            current_item = self.files_list.takeItem(current_row)
            current_file = self.input_files.pop(current_row)
            self.files_list.insertItem(current_row - 1, current_item)
            self.input_files.insert(current_row - 1, current_file)
            self.files_list.setCurrentRow(current_row - 1)

    def move_item_down(self):
        """
        Move the selected file down in the list.
        """
        current_row = self.files_list.currentRow()
        if current_row < self.files_list.count() - 1:
            current_item = self.files_list.takeItem(current_row)
            current_file = self.input_files.pop(current_row)
            self.files_list.insertItem(current_row + 1, current_item)
            self.input_files.insert(current_row + 1, current_file)
            self.files_list.setCurrentRow(current_row + 1)

    def remove_selected_file(self):
        """
        Remove the selected file from the list.
        """
        current_row = self.files_list.currentRow()
        if current_row >= 0:
            self.files_list.takeItem(current_row)
            self.input_files.pop(current_row)
            if len(self.input_files) == 0:
                self.process_btn.setEnabled(False)

    def toggle_image_extraction(self, state):
        """
        Toggle image extraction option.

        Parameters:
        - state (int): The state of the checkbox (0 or 2).
        """
        self.extract_images = state > 0

    def change_ui_language(self, language):
        """
        Change the UI language.

        Parameters:
        - language (str): The selected language.
        """
        self.current_language = language
        selected_lang = TRANSLATIONS.get(language, TRANSLATIONS["Italiano"])

        # Update all UI elements with translations
        self.setWindowTitle(selected_lang["window_title"])
        self.select_files_btn.setText(selected_lang["select_files"])
        self.process_btn.setText(selected_lang["generate"])
        self.ui_language_label.setText(selected_lang["interface_lang"])
        self.output_language_label.setText(selected_lang["output_lang"])
        self.output_format_label.setText(selected_lang.get("output_format"))
        self.files_label.setText(selected_lang["selected_files"])
        self.move_up_btn.setText(selected_lang["move_up"])
        self.move_down_btn.setText(selected_lang["move_down"])
        self.remove_btn.setText(selected_lang["remove"])
        self.docx_checkbox.setText("DOCX")
        self.pdf_checkbox.setText("PDF")
        self.image_extraction_label.setText(selected_lang.get("content_options"))
        self.image_checkbox.setText(selected_lang.get("include_images"))

    def set_output_language(self, language):
        """
        Set the output language for the summary.

        Parameters:
        - language (str): The selected output language.
        """
        self.output_language = language

    def select_files(self):
        """
        Open a file dialog to select files and add them to the list.
        """
        files, _ = QFileDialog.getOpenFileNames(
            self,
            "Select Files",
            "",
            "Document Files (*.pdf *.pptx)"
        )
        if files:
            # Add new files to the list
            self.input_files.extend(files)
            # Add filenames to the list widget
            for file_path in files:
                self.files_list.addItem(os.path.basename(file_path))
            self.process_btn.setEnabled(True)

            # Update status with current language
            selected_lang = TRANSLATIONS.get(self.current_language, TRANSLATIONS["Italiano"])
            self.status_label.setText(f"{selected_lang['selected_files']} {len(self.input_files)}")

    def process_files(self):
        """
        Process the selected files, generate summaries, and save them in the selected format.
        """
        global docx_path, pdf_path, file_filter
        if not self.input_files:
            return

        total_pages = 0
        file_page_counts = {}

        for file_path in self.input_files:
            file_extension = os.path.splitext(file_path)[1].lower()
            try:
                if file_extension == '.pdf':
                    import PyPDF2
                    with open(file_path, 'rb') as file:
                        pdf_reader = PyPDF2.PdfReader(file)
                        page_count = len(pdf_reader.pages)
                elif file_extension == '.pptx':
                    from pptx import Presentation
                    ppt = Presentation(file_path)
                    page_count = len(ppt.slides)
                else:
                    page_count = 0

                file_page_counts[file_path] = page_count
                total_pages += page_count
            except Exception as e:
                print(f"Error counting pages in {file_path}: {e}")
                file_page_counts[file_path] = 0

        progress = QProgressDialog("Processing Files...", None, 0, total_pages+len(self.input_files), self)
        progress.setWindowModality(Qt.WindowModality.WindowModal)
        progress.show()

        summaries = []
        current_page_progress = 0

        for i, file_path in enumerate(self.input_files):

            # Extract text based on file type
            file_extension = os.path.splitext(file_path)[1].lower()

            # Inside the process_files method, replace the file type checking block with:
            try:
                if file_extension == '.pdf':
                    if self.extract_images:
                        text, current_page_progress = extract_text_and_images_from_pdf(file_path, progress, current_page_progress)
                    else:
                        text, current_page_progress = extract_text_from_pdf(file_path, progress, current_page_progress)
                elif file_extension == '.pptx':
                    if self.extract_images:
                        text, current_page_progress = extract_text_and_images_from_pptx(file_path, progress, current_page_progress)
                    else:
                        text, current_page_progress = extract_text_from_pptx(file_path, progress, current_page_progress)
                else:
                    continue

                # Create prompt and get summary
                prompt = create_summary_prompt(text, self.output_language)
                section_content = send_request_to_api(prompt)

                # Update progress bar
                current_page_progress += 1
                progress.setValue(current_page_progress)
                QApplication.processEvents()  # Ensure UI updates

                summaries.append({
                    'title' : f"{i + 1}. {os.path.splitext(os.path.basename(file_path))[0]}",
                    'content': section_content
                })

            except Exception as e:
                error_msg = f"Error processing {os.path.basename(file_path)}: {str(e)}"
                summaries.append({
                    'title': f"{i + 1}. {os.path.splitext(os.path.basename(file_path))[0]}",
                    'content': error_msg
                })

        selected_lang = TRANSLATIONS.get(self.current_language, TRANSLATIONS["Italiano"])

        if self.save_as_docx and self.save_as_pdf:
            file_filter = "Word o PDF Files (*.docx *.pdf)"
        elif self.save_as_docx:
            file_filter = "Word Files (*.docx)"
        elif self.save_as_pdf:
            file_filter = "PDF Files (*.pdf)"

        output_file, _ = QFileDialog.getSaveFileName(
            self,
            selected_lang.get("save_dialog", "Save Summary"),
            "",
            file_filter
        )

        if output_file:
            base_path, ext = os.path.splitext(output_file)

            if not ext:
                if self.save_as_docx:
                    docx_path = f"{base_path}.docx"
                if self.save_as_pdf:
                    pdf_path = f"{base_path}.pdf"
            else:
                if ext.lower() == '.docx':
                    docx_path = output_file
                    pdf_path = f"{base_path}.pdf"
                elif ext.lower() == '.pdf':
                    pdf_path = output_file
                    docx_path = f"{base_path}.docx"
                else:
                    docx_path = f"{base_path}.docx"
                    pdf_path = f"{base_path}.pdf"

            try:
                if self.save_as_docx:
                    save_as_docx_file(docx_path, summaries)

                if self.save_as_pdf:
                    save_as_pdf_file(pdf_path, summaries)

                self.status_label.setText(selected_lang.get("success_message", "Summary created successfully!"))

            except Exception as e:
                QMessageBox.critical(self, "Error", f"Error saving file: {str(e)}")


def main():
    app = QApplication(sys.argv)
    window = DocumentSummaryApp()
    window.show()
    sys.exit(app.exec())


if __name__ == '__main__':
    os.chdir(os.path.dirname(os.path.abspath(__file__)))
    load_dotenv(dotenv_path="../.env")
    main()