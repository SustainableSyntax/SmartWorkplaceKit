import sys
import os
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget,
    QPushButton, QVBoxLayout, QHBoxLayout,
    QLabel, QFileDialog, QStatusBar,
    QMessageBox, QComboBox, QSlider,
    QInputDialog, QDialog, QProgressBar,
    QGridLayout, QListWidget
)
from PyQt5.QtGui import QPixmap, QImage
from PyQt5.QtCore import Qt


class ImageProcessor:
    @staticmethod
    def load_image(file_path):
        return QImage(file_path)

    @staticmethod
    def save_image(image, file_path):
        image.save(file_path)

    @staticmethod
    def resize_image(image, factor, algorithm=Qt.SmoothTransformation):
        new_width = int(image.width() * factor)
        new_height = int(image.height() * factor)
        return image.scaled(
            new_width,
            new_height,
            Qt.KeepAspectRatio,
            algorithm
        )


class ImageData:
    def __init__(self):
        self.image_history = []
        self.current_index = -1  # No image at the start

    def add_to_history(self, image):
        # When a new action is done, truncate the list to the current index
        self.image_history = self.image_history[:self.current_index + 1]
        self.image_history.append(image.copy())
        self.current_index += 1

    def undo(self):
        if self.current_index > 0:
            self.current_index -= 1
            return self.image_history[self.current_index]
        return None

    def redo(self):
        if self.current_index < len(self.image_history) - 1:
            self.current_index += 1
            return self.image_history[self.current_index]
        return None


class UIComponents:
    @staticmethod
    def create_button(text, callback):
        button = QPushButton(text)
        button.clicked.connect(callback)
        return button

    @staticmethod
    def create_combo_box(options):
        combo_box = QComboBox()
        # Include the new resize options
        combo_box.addItems(['25%', '50%', '75%', '100%', '200%',
                           '300%', '400%', '500%', '700%', 'Anpassen...'])
        return combo_box

    @staticmethod
    def create_label(text):
        label = QLabel(text)
        label.setAlignment(Qt.AlignCenter)
        return label


class UIElements:
    def __init__(self):
        self.open_button = None
        self.size_combo = None
        self.file_size_slider = None
        self.file_size_label = None
        self.apply_button = None
        self.save_button = None
        self.image_label = None
        self.width_input = None
        self.height_input = None
        self.aspect_ratio_lock = None

    def __repr__(self):
        return "UIElements()"

    def __str__(self):
        return "This class holds the necessary UI elements"


class DraggableListWidget(QListWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAcceptDrops(True)
        self.setDragDropMode(QListWidget.DragDrop)  # Enable drag & drop
        # Enable multiple selection
        self.setSelectionMode(QListWidget.ExtendedSelection)

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dragMoveEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dropEvent(self, event):
        for url in event.mimeData().urls():
            if url.isLocalFile():
                self.addItem(url.toLocalFile())


class BatchProcessingDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle('Batch Image Resizing')
        self.layout = QGridLayout(self)

        # Use the DraggableListWidget instead of QListWidget
        self.image_list_widget = DraggableListWidget(self)
        self.layout.addWidget(self.image_list_widget, 0, 0, 1, 2)

        self.add_images_button = QPushButton('Add Images', self)
        self.add_images_button.clicked.connect(self.add_images)
        self.layout.addWidget(self.add_images_button, 1, 0)

        self.start_button = QPushButton('Start Resizing', self)
        self.start_button.clicked.connect(self.start_resizing)
        self.layout.addWidget(self.start_button, 1, 1)

        self.progress_bar = QProgressBar(self)
        self.layout.addWidget(self.progress_bar, 2, 0, 1, 2)

        self.resize_factors = {'25%': 0.25, '50%': 0.50, '75%': 0.75,
                               '100%': 1.00, '200%': 2.00, '300%': 3.00,
                               '400%': 4.00, '500%': 5.00, '700%': 7.00
                               }
        self.selected_factor = None

    def add_images(self):
        files, _ = QFileDialog.getOpenFileNames(
            self, "Select Images", "", "Image Files (*.png *.jpg *.jpeg)")
        for file in files:
            self.image_list_widget.addItem(file)

    def start_resizing(self):
        self.selected_factor = self.parent().get_resize_factor()
        for i in range(self.image_list_widget.count()):
            item = self.image_list_widget.item(i)
            image_path = item.text()
            image = ImageProcessor.load_image(image_path)

            resized_image = ImageProcessor.resize_image(
                image, self.selected_factor)

            # Erstellt einen neuen Dateinamen,
            # indem "_resized" vor der Dateiendung eingefügt wird.
            base, extension = os.path.splitext(image_path)
            new_image_path = f"{base}_resized{extension}"

            ImageProcessor.save_image(resized_image, new_image_path)

            self.progress_bar.setValue(
                int((i + 1) / self.image_list_widget.count() * 100))
        QMessageBox.information(self, "Batch Resizing", "Resizing Completed")


class ImageResizingApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.title = 'Komprimierungsfreier Image Resizer'
        self.image_data = ImageData()  # Use ImageData instance
        self.ui_elements = UIElements()  # Use UIElements instance
        self.init_ui()

        self.setAcceptDrops(True)

    def init_ui(self):
        self.setWindowTitle(self.title)
        # Adjusted window size for a more spacious layout
        self.setGeometry(100, 100, 900, 650)
        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)

        # Main widget and layout configuration
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        main_layout = QVBoxLayout()
        controls_layout = QHBoxLayout()
        resize_layout = QVBoxLayout()  # New layout for resize options

        # Initialize UI elements
        self.ui_elements.open_button = UIComponents.create_button(
            'Bild öffnen', self.open_image)
        controls_layout.addWidget(self.ui_elements.open_button)

        self.ui_elements.size_combo = UIComponents.create_combo_box(
            ['25%', '50%', '75%', 'Anpassen...'])
        resize_layout.addWidget(UIComponents.create_label('Größe anpassen:'))
        resize_layout.addWidget(self.ui_elements.size_combo)

        # File Size Selection (New Feature)
        self.ui_elements.file_size_slider = QSlider(Qt.Horizontal)
        self.ui_elements.file_size_slider.setMinimum(
            100)  # Minimum file size (KB)
        self.ui_elements.file_size_slider.setMaximum(
            5000)  # Maximum file size (KB)
        self.ui_elements.file_size_slider.setValue(1000)  # Default value (KB)
        self.ui_elements.file_size_slider.setTickPosition(QSlider.TicksBelow)
        self.ui_elements.file_size_slider.setTickInterval(500)

        # Label to display the current file size
        self.ui_elements.file_size_label = QLabel('1000 KB')
        self.ui_elements.file_size_slider.valueChanged.connect(
            lambda: self.ui_elements.file_size_label.setText(
                f"{self.ui_elements.file_size_slider.value()} KB")
        )
        resize_layout.addWidget(
            UIComponents.create_label('Maximale Dateigröße (KB):'))
        resize_layout.addWidget(self.ui_elements.file_size_slider)
        resize_layout.addWidget(self.ui_elements.file_size_label)

        self.ui_elements.apply_button = UIComponents.create_button(
            'Anwenden', self.apply_resize)
        controls_layout.addWidget(self.ui_elements.apply_button)

        self.ui_elements.batch_button = UIComponents.create_button(
            'Batch Processing', self.open_batch_dialog)
        controls_layout.addWidget(self.ui_elements.batch_button)

        self.ui_elements.save_button = UIComponents.create_button(
            'Bild speichern', self.save_image)
        self.ui_elements.save_button.setEnabled(False)
        controls_layout.addWidget(self.ui_elements.save_button)

        self.ui_elements.image_label = UIComponents.create_label(
            'Bild wird hier angezeigt')
        # Ensure there's enough space to display images
        self.ui_elements.image_label.setMinimumSize(400, 300)

        # Assembling the layout
        main_layout.addLayout(controls_layout)
        main_layout.addLayout(resize_layout)  # Add the new resize layout
        main_layout.addWidget(self.ui_elements.image_label)
        main_widget.setLayout(main_layout)

        # Styling (example)
        self.apply_styling()

    def open_batch_dialog(self):
        dialog = BatchProcessingDialog(self)
        dialog.exec_()

    def apply_styling(self):
        self.setStyleSheet("""
        QMainWindow {
            background-color: #303030; /* Anthracite for main window background */
        }
        QWidget, QComboBox QAbstractItemView {
            font-size: 14px;
            color: #D8BFD8; /* Lilac for better readability */
            background-color: #303030; /* Anthracite for dropdown background */
            selection-background-color: #483D8B; /* Dark blue for selection */
            selection-color: #E1E1E1; /* Light grey for selected item text */
        }
        QPushButton {
            background-color: #483D8B; /* Dark blue */
            color: #E1E1E1; /* Light grey for text to ensure readability */
            border-radius: 5px;
            padding: 10px;
            border: 1px solid #9370DB; /* A lighter shade of dark blue for border */
        }
        QPushButton:disabled {
            background-color: #4F4F4F; /* Dark grey */
            color: #A9A9A9; /* Light grey */
            border: 1px solid #696969; /* Dim grey for border */
        }
        QLabel {
            qproperty-alignment: 'AlignCenter';
            color: #E1E1E1; /* Changing label text to light grey for readability */
        }
        QComboBox {
            margin: 5px;
            background-color: #2F4F4F; /* Dark Slate Grey */
            border: 1px solid #9370DB; /* A lighter shade of dark blue for border */
        }
        QComboBox::drop-down {
            background-color: #2F4F4F;
            border: 1px solid #9370DB;
        }
        QComboBox::down-arrow {
            image: url(:/icons/down-arrow.png); /* Ensure you have an appropriate icon */
            width: 14px;
            height: 14px;
        }
        QSlider::groove:horizontal {
            border: 1px solid #999999;
            height: 8px; /* The Groove's height */
            background: #483D8B; /* Dark blue */
            margin: 2px 0;
        }
        QSlider::handle:horizontal {
            background: #D8BFD8; /* Lilac */
            border: 1px solid #5c5c5c;
            width: 18px;
            margin: -2px 0; /* Handle is larger than groove */
            border-radius: 3px;
        }
        QStatusBar {
            background-color: #303030; /* Anthracite */
            color: #E1E1E1; /* Light grey for readability */
        }
        """)

    def open_image(self):
        file_name, _ = QFileDialog.getOpenFileName(
            self, "Bild öffnen", "", "Bild Dateien (*.png *.jpg *.jpeg)")
        if file_name:
            self.image_data.original_image_path = file_name
            self.image_data.original_image = ImageProcessor.load_image(
                file_name)
            self.ui_elements.image_label.setPixmap(
                QPixmap.fromImage(self.image_data.original_image).scaled(
                    self.ui_elements.image_label.size(), Qt.KeepAspectRatio))
            self.status_bar.showMessage("Bild erfolgreich geladen")
            self.ui_elements.save_button.setEnabled(False)

    def apply_resize(self):
        if not self.image_data.original_image:
            QMessageBox.warning(self, "Fehler", "Bitte zuerst ein Bild öffnen")
            return

        factor = self.get_resize_factor()
        if factor is not None:
            self.image_data.resized_image = ImageProcessor.resize_image(
                self.image_data.original_image, factor)
            self.ui_elements.image_label.setPixmap(
                QPixmap.fromImage(self.image_data.resized_image))
            self.status_bar.showMessage("Bildgröße erfolgreich geändert")
            self.ui_elements.save_button.setEnabled(True)

    def get_resize_factor(self):
        selected_text = self.ui_elements.size_combo.currentText()
        if selected_text == 'Anpassen...':
            # Open a dialog to get the exact resize factor
            factor, ok = QInputDialog.getDouble(
                self, "Größe anpassen", "Neuer Größenfaktor:",
                1.00, 0.01, 10.00, 2)
            if ok:
                return factor
        else:
            # For predefined percentages, strip the '%' character and convert to a float
            return float(selected_text.rstrip('%')) / 100

    def save_image(self):
        if not self.image_data.resized_image:
            QMessageBox.warning(
                self, "Fehler", "Es wurde kein Bild zum Speichern gefunden")
            return

        base, extension = os.path.splitext(self.image_data.original_image_path)
        suggested_file_name = base + "_cropped" + extension
        file_name, _ = QFileDialog.getSaveFileName(
            self, "Bild speichern", suggested_file_name, "Bild Dateien (*.png *.jpg *.jpeg)")
        if file_name:
            ImageProcessor.save_image(self.image_data.resized_image, file_name)
            self.status_bar.showMessage(
                "Bild erfolgreich gespeichert als: " + file_name)

    # Override the dragEnterEvent method
    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    # Override the dropEvent method
    def dropEvent(self, event):
        urls = event.mimeData().urls()
        if urls:
            # Attempt to load the first image from the dropped files
            image_path = urls[0].toLocalFile()
            self.load_image(image_path)

    def load_image(self, file_path):
        """Load and display an image from a file path."""
        if os.path.isfile(file_path):
            self.image_data.original_image_path = file_path
            self.image_data.original_image = ImageProcessor.load_image(
                file_path)
            self.ui_elements.image_label.setPixmap(
                QPixmap.fromImage(self.image_data.original_image).scaled(
                    self.ui_elements.image_label.size(), Qt.KeepAspectRatio, Qt.SmoothTransformation))
            self.status_bar.showMessage("Bild erfolgreich geladen")
            self.ui_elements.save_button.setEnabled(False)
            # Update the image history
            self.image_data.add_to_history(self.image_data.original_image)


if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = ImageResizingApp()
    ex.show()
    sys.exit(app.exec_())
