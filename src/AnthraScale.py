import sys
import os
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget,
    QPushButton, QVBoxLayout, QLabel,
    QFileDialog, QStatusBar, QMessageBox,
    QComboBox, QSlider, QInputDialog,
    QDialog, QProgressBar, QGridLayout,
    QListWidget, QToolBar, QAction
)
from PyQt5.QtGui import QPixmap, QImage
from PyQt5.QtCore import Qt


class ImageProcessor:
    """
    Simple utility class to process images
    and manage image operations such as load,
    resize, and save.

    Attributes:
    -----------
            None

    Methods:
    ----------
            - load_image(file_path) -> QImage
            - save_image(image, file_path) -> None
            - resize_image(
                image,
                factor,
                algorithm=Qt.SmoothTransformation) -> QImage
    """
    @staticmethod
    def load_image(file_path):
        """
        Load an image from a file path.

        Parameters:
        ------------
            file_path : str
                The image file path to load an image from.

        Returns:
        ------------
            QImage object
                The loaded image.
        """
        return QImage(file_path)

    @staticmethod
    def save_image(image, file_path, quality=85, format=None):
        """
        Save an image to a file.

        Parameters:
        ------------
            image : QImage
                The image to save.

            file_path : str
                The file path to save the image to.

            quality : int
                The quality of the image (for JPEGs).

            format : str
                The format of the image to save.
                Supported options are: 'PNG', 'JPG', 'BMP', and 'GIF'.

        Returns:
        ------------
            None
        """
        if format is None:
            format = 'PNG'  # Default format
        image.save(file_path, format, quality)

    @staticmethod
    def resize_image(image, factor, algorithm=Qt.SmoothTransformation):
        """
        Resize an image by a given factor.

        Parameters:
        ------------
            image : QImage
                The image to resize.

            factor : float
                The factor to resize the image by.

            algorithm : Qt.TransformationMode
                The transformation algorithm to use.
                Default is Qt.SmoothTransformation.

        Returns:
        ------------
            QImage object
                The resized image.
        """
        new_width = int(image.width() * factor)
        new_height = int(image.height() * factor)
        return image.scaled(
            new_width,
            new_height,
            Qt.KeepAspectRatio,
            algorithm
        )


class ImageData:
    """
    Holds the image history and the current image
    being worked on.

    Attributes:
    ------------
            image_history : list
                A list to keep track of all image modifications.

            current_index : int
                An index to keep track of the current image
                shown in the UI.

            original_image_path : str
                The file path of the original image.

            original_image : QImage
                The original image loaded from file.

            resized_image : QImage
                The image that has been modified.

    Methods:
    -----------
            - add_to_history(image: QImage) -> None
            - undo() -> QImage
            - redo() -> QImage
    """

    def __init__(self):
        """
        Initialize the image history and the current index.

        Parameters:
        -----------
            None

        Returns:
        ----------
            None
        """
        self.image_history = []
        self.current_index = -1  # No image at the start

    def add_to_history(self, image):
        """
        Add the new image to the image history.

        Parameters:
        ------------
            image : QImage
                The image to add to the history.

        Returns:
        ---------
            None
        """
        # When a new action is done, truncate the list to the current index
        self.image_history = self.image_history[:self.current_index + 1]
        self.image_history.append(image.copy())
        self.current_index += 1

    def undo(self):
        """
        Return the previous image from the history.

        Parameters:
        -----------
            None

        Returns:
        ---------
            QImage object
                The previous image.
        """
        if self.current_index > 0:
            self.current_index -= 1
            return self.image_history[self.current_index]
        return None

    def redo(self):
        """
        Return the next image from the history.

        Parameters:
        -----------
            None

        Returns:
        ---------
            QImage object
                The next image.
        """
        if self.current_index < len(self.image_history) - 1:
            self.current_index += 1
            return self.image_history[self.current_index]
        return None


class UIComponents:
    """
    A collection of common UI components for reuse.

    Attributes:
    -----------
            None

    Methods:
    -----------
            - create_button(text, callback) -> QPushButton
            - create_combo_box(options) -> QComboBox
            - create_label(text) -> QLabel
    """
    @staticmethod
    def create_button(text, callback):
        """
        Create a QPushButton.

        Parameters:
        ------------
            text : str
                The text to display on the button.

            callback : function
                The function to call when the button is clicked.

        Returns:
        ---------
            QPushButton
                The newly created button.
        """
        button = QPushButton(text)
        button.clicked.connect(callback)
        return button

    @staticmethod
    def create_combo_box(options):
        """
        Create a QComboBox.

        Parameters:
        -----------
            options : list
                The options to include in the combo box.

        Returns:
        ---------
            QComboBox
                The newly created combo box.
        """
        combo_box = QComboBox()
        # Include the new resize options
        combo_box.addItems(options)
        return combo_box

    @staticmethod
    def create_label(text):
        """
        Create a QLabel.

        Parameters:
        -----------
            text : str
                The text to display in the label.

        Returns:
        ---------
            QLabel
                The newly created label.
        """
        label = QLabel(text)
        label.setAlignment(Qt.AlignCenter)
        return label


class UIElements:
    """
    A collection of the necessary UI elements
    for the main application.

    Attributes:
    ------------
            open_button : QPushButton
                A button to open an image file.

            size_combo : QComboBox
                A combo box to select the resize size.

            file_size_slider : QSlider
                A slider to select the maximum file size.

            file_size_label : QLabel
                A label to display the current file size.

            apply_button : QPushButton
                A button to apply the resize operation.

            save_button : QPushButton
                A button to save the resized image.

            image_label : QLabel
                A label to display the image.

            width_input : QLineEdit
                An input to enter the new width.

            height_input : QLineEdit
                An input to enter the new height.

            aspect_ratio_lock : QCheckBox
                A checkbox to lock the aspect ratio.

            quality_slider : QSlider
                A slider to adjust the image quality.

            format_combo_box : QComboBox
                A combo box to select the image format.

    Methods:
    ---------
            None
    """

    def __init__(self):
        """
        Initialize all the UI elements as None.

        Parameters:
        -----------
            None

        Returns:
        ----------
            None
        """
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
        self.quality_slider = None
        self.format_combo_box = None
        self.undo_button = None
        self.redo_button = None

    def __repr__(self):
        """
        Return the UIElements instance.

        Parameters:
        -----------
            None

        Returns:
        ----------
            str
                A string representation of the UIElements instance.
        """
        return [str(self.__class__), self.__dict__]

    def __str__(self):
        """
        Return a description of the UIElements instance.

        Parameters:
        -----------
            None

        Returns:
        ---------
            str
                A string describing the UIElements instance.
        """
        return "This class holds the necessary UI elements"


class DraggableListWidget(QListWidget):
    """
    A QListWidget that accepts dragged files.
    Inherits from QListWidget to allow the handling of
    drag and drop events.

    Methods:
    -----------
            dragEnterEvent(event) -> None
            dragMoveEvent(event) -> None
            dropEvent(event) -> None
    """

    def __init__(self, parent=None):
        """
        Initialize the QListWidget.

        Parameters:
        ------------
            parent : QWidget
                The parent widget, if any, for the QListWidget.

        Returns:
        ---------
            None
        """
        super().__init__(parent)
        self.setAcceptDrops(True)
        self.setDragDropMode(QListWidget.DragDrop)  # Enable drag & drop
        # Enable multiple selection
        self.setSelectionMode(QListWidget.ExtendedSelection)

    def dragEnterEvent(self, event):
        """
        Accept the file drag event.

        Parameters:
        ------------
            event : QDragEnterEvent
                The drag enter event.

        Returns:
        ---------
            None
        """
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dragMoveEvent(self, event):
        """
        Accept the file move event.

        Parameters:
        ------------
            event : QDragMoveEvent
                The drag move event.

        Returns:
        ---------
            None
        """
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dropEvent(self, event):
        """
        Accept the file drop event.

        Parameters:
        ------------
            event : QDropEvent
                The drop event.

        Returns:
        ---------
            None
        """
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
        self.setGeometry(100, 100, 900, 650)
        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)

        # Menüleiste
        menubar = self.menuBar()
        file_menu = menubar.addMenu('&Datei')
        edit_menu = menubar.addMenu('&Bearbeiten')

        # Haupt-Widget und Layout-Konfiguration
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        main_layout = QVBoxLayout(main_widget)

        # Toolbar für die Aktionen
        action_toolbar = QToolBar("Actions")
        self.addToolBar(Qt.TopToolBarArea, action_toolbar)

        self.actionExit = QAction('&Quit', self)
        self.actionExit.setShortcut('Ctrl+Q')
        self.actionExit.triggered.connect(self.close)

        self.actionAbout = QAction('&About', self)
        self.actionAbout.setShortcut('Ctrl+A')
        self.actionAbout.triggered.connect(self.showAbout)

        self.actionTutorial = QAction('&Tutorial', self)
        self.actionTutorial.setShortcut('Ctrl+T')
        self.actionTutorial.triggered.connect(self.showTutorial)

        self.actionContact = QAction('&Contact', self)
        self.actionContact.setShortcut('Ctrl+C')
        self.actionContact.triggered.connect(self.showContact)

        # Öffnen-Aktion
        open_action = QAction('Bild öffnen', self)
        open_action.triggered.connect(self.open_image)
        file_menu.addAction(open_action)
        action_toolbar.addAction(open_action)

        # Anwenden-Aktion
        apply_action = QAction('Anwenden', self)
        apply_action.triggered.connect(self.apply_resize)
        edit_menu.addAction(apply_action)
        action_toolbar.addAction(apply_action)

        # Undo-Aktion
        undo_action = QAction('Undo', self)
        undo_action.triggered.connect(self.undo_action)
        edit_menu.addAction(undo_action)
        action_toolbar.addAction(undo_action)

        # Redo-Aktion
        redo_action = QAction('Redo', self)
        redo_action.triggered.connect(self.redo_action)
        edit_menu.addAction(redo_action)
        action_toolbar.addAction(redo_action)

        # Batch-Verarbeitung-Aktion
        batch_action = QAction('Batch Processing', self)
        batch_action.triggered.connect(self.open_batch_dialog)
        edit_menu.addAction(batch_action)
        action_toolbar.addAction(batch_action)

        # Speichern-Aktion
        save_action = QAction('Bild speichern', self)
        save_action.triggered.connect(self.save_image)
        file_menu.addAction(save_action)
        action_toolbar.addAction(save_action)

        file_menu.addAction(self.actionAbout)
        file_menu.addAction(self.actionTutorial)
        file_menu.addAction(self.actionContact)
        file_menu.addSeparator()
        file_menu.addAction(self.actionExit)

        # Rest der UI-Elemente wie Slider, ComboBoxen etc.
        # Größenanpassungs-ComboBox
        main_layout.addWidget(UIComponents.create_label('Größe anpassen:'))
        self.ui_elements.size_combo = UIComponents.create_combo_box(
            ['25%', '50%', '75%', '100%', '200%', '300%', '400%', '500%', '700%', 'Anpassen...'])
        main_layout.addWidget(self.ui_elements.size_combo)

        # Max File Size Slider
        main_layout.addWidget(UIComponents.create_label('Max file size (KB):'))
        self.ui_elements.file_size_slider = QSlider(Qt.Horizontal)
        self.ui_elements.file_size_slider.setMinimum(100)
        self.ui_elements.file_size_slider.setMaximum(5000)
        self.ui_elements.file_size_slider.setValue(1000)
        self.ui_elements.file_size_slider.setTickPosition(QSlider.TicksBelow)
        self.ui_elements.file_size_slider.setTickInterval(500)
        main_layout.addWidget(self.ui_elements.file_size_slider)
        self.ui_elements.file_size_label = QLabel('1000 KB')
        main_layout.addWidget(self.ui_elements.file_size_label)

        # Quality Slider
        main_layout.addWidget(UIComponents.create_label('Quality:'))
        self.ui_elements.quality_slider = QSlider(Qt.Horizontal)
        self.ui_elements.quality_slider.setMinimum(1)
        self.ui_elements.quality_slider.setMaximum(100)
        self.ui_elements.quality_slider.setValue(85)
        self.ui_elements.quality_slider.setTickPosition(QSlider.TicksBelow)
        self.ui_elements.quality_slider.setTickInterval(10)
        main_layout.addWidget(self.ui_elements.quality_slider)

        # Format Combo Box
        main_layout.addWidget(UIComponents.create_label('Format:'))
        self.ui_elements.format_combo_box = UIComponents.create_combo_box(
            ['JPEG', 'PNG', 'BMP', 'GIF'])
        main_layout.addWidget(self.ui_elements.format_combo_box)

        # Image Label
        self.ui_elements.image_label = UIComponents.create_label(
            'Bild wird hier angezeigt')
        self.ui_elements.image_label.setMinimumSize(400, 300)
        main_layout.addWidget(self.ui_elements.image_label)

        # Fügen Sie das layout zum main_widget hinzu
        main_widget.setLayout(main_layout)

        self.apply_styling()

    def open_batch_dialog(self):
        dialog = BatchProcessingDialog(self)
        dialog.exec_()

    def apply_styling(self):
        self.setStyleSheet("""
        QMainWindow {
            background-color: #303030;
        }
        QWidget, QComboBox QAbstractItemView {
            font-size: 14px;
            color: #D8BFD8;
            background-color: #303030;
            selection-background-color: #483D8B;
            selection-color: #E1E1E1;
        }
        QPushButton {
            background-color: #483D8B;
            color: #E1E1E1;
            border-radius: 5px;
            padding: 10px;
            border: 1px solid #9370DB;
        }
        QPushButton:disabled {
            background-color: #4F4F4F;
            color: #A9A9A9;
            border: 1px solid #696969;
        }
        QLabel {
            qproperty-alignment: 'AlignCenter';
            color: #E1E1E1;
        }
        QComboBox {
            margin: 5px;
            background-color: #2F4F4F;
            border: 1px solid #9370DB;
        }
        QComboBox::drop-down {
            background-color: #2F4F4F;
            border: 1px solid #9370DB;
        }
        QComboBox::down-arrow {
            image: url(:/icons/down-arrow.png);
            width: 14px;
            height: 14px;
        }
        QSlider::groove:horizontal {
            border: 1px solid #999999;
            height: 8px;
            background: #483D8B;
            margin: 2px 0;
        }
        QSlider::handle:horizontal {
            background: #D8BFD8;
            border: 1px solid #5c5c5c;
            width: 18px;
            margin: -2px 0;
            border-radius: 3px;
        }
        QStatusBar {
            background-color: #303030;
            color: #E1E1E1;
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
            # For predefined percentages,
            # strip the '%' character and convert to a float
            return float(selected_text.rstrip('%')) / 100

    def save_image(self):
        if not self.image_data.resized_image:
            QMessageBox.warning(
                self, "Fehler", "Es wurde kein Bild zum Speichern gefunden")
            return

        base, extension = os.path.splitext(self.image_data.original_image_path)
        suggested_file_name = base + "_cropped" + extension
        file_name, _ = QFileDialog.getSaveFileName(
            self,
            "Bild speichern",
            suggested_file_name,
            "Bild Dateien (*.png *.jpg *.jpeg)"
        )
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
                    self.ui_elements.image_label.size(),
                    Qt.KeepAspectRatio,
                    Qt.SmoothTransformation
                )
            )
            self.status_bar.showMessage("Bild erfolgreich geladen")
            self.ui_elements.save_button.setEnabled(False)
            # Update the image history
            self.image_data.add_to_history(self.image_data.original_image)

    def undo_action(self):
        previous_image = self.image_data.undo()
        if previous_image:
            self.ui_elements.image_label.setPixmap(
                QPixmap.fromImage(previous_image))
            self.status_bar.showMessage("Undo: Bild wurde zurückgesetzt")

    def redo_action(self):
        next_image = self.image_data.redo()
        if next_image:
            self.ui_elements.image_label.setPixmap(
                QPixmap.fromImage(next_image))
            self.status_bar.showMessage("Redo: Bild wurde wiederhergestellt")

    def showAbout(self):
        QMessageBox.about(
            self,
            "About ImageResizer",
            "ImageResizer is a small image resizing utility built with PyQt5 and OpenCV"
        )

    def showTutorial(self):
        QMessageBox.information(
            self,
            "Tutorial",
            "This is a tutorial message"
        )

    def showContact(self):
        QMessageBox.information(
            self,
            "Contact",
            "Email: siemenshendrik1@gmail.com"
        )

    def closeEvent(self, event):
        reply = QMessageBox.question(
            self,
            'Message',
            "Are you sure you want to quit?",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )
        if reply == QMessageBox.Yes:
            event.accept()
        elif reply == QMessageBox.No:
            event.ignore()


if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = ImageResizingApp()
    ex.show()
    sys.exit(app.exec_())
