import os
import sys
import pandas
import random
import csv
from PyQt5.QtCore import *
from PyQt5.QtGui import *
from PyQt5.QtWidgets import *


class MyMainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Tube Sorter Worklist Generator")
        self.setMinimumSize(750, 275)

        self.initUI()
        self.central_widget = Application()
        self.layout = QVBoxLayout()
        self.layout.addWidget(self.central_widget)
        self.setCentralWidget(self.central_widget)

    def initUI(self):
        main_menu = self.menuBar()

        file_menu = main_menu.addMenu("File")
        help_menu = main_menu.addMenu("Help")

        global set_destination_button
        set_destination_button = QAction("Set Output Destination", self)
        set_destination_button.setStatusTip("Set destination path for output worklist")
        file_menu.addAction(set_destination_button)

        global open_button
        open_button = QAction("Open SSMINI File", self)
        open_button.setStatusTip("Open SSMINI data file for upload")
        file_menu.addAction(open_button)

        exit_button = QAction("Exit", self)
        exit_button.setStatusTip("Exit application")
        exit_button.triggered.connect(self.close)
        file_menu.addAction(exit_button)

        view_help_button = QAction("View Help", self)
        view_help_button.setStatusTip("View instructions for use")
        help_menu.addAction(view_help_button)
        view_help_button.triggered.connect(self.HelpWindow)

        about_button = QAction("About", self)
        about_button.setStatusTip("About Tube Sorter Worklist Generator")
        help_menu.addAction(about_button)
        about_button.triggered.connect(self.AboutApp)

    def AboutApp(self):
        about_window = QMessageBox.information(self, "About Application", "Version: 1.0\nAuthor: Dennis Khine",
                                           QMessageBox.Ok)

    def HelpWindow(self):
        self.help_window = QMainWindow()
        self.help_window.setWindowTitle("Application Help")
        self.help_window.setMinimumSize(1050, 600)
        self.help_window.setStyleSheet("background-color: #7FDAFF")
        self.help_window_central_scroll_area = QScrollArea()
        self.help_window_central_widget = QWidget()
        self.help_window_central_scroll_area.setWidget(self.help_window_central_widget)
        self.help_window_central_scroll_area.setWidgetResizable(True)
        self.help_window.setCentralWidget(self.help_window_central_scroll_area)
        self.help_window_central_widget_layout = QVBoxLayout()

        self.headers = [QLabel("Overview"), QLabel("\nMenu Bar"), QLabel("\nMain Interface"), QLabel("\nTroubleshooting")]
        for i in self.headers:
            i.setStyleSheet("color: black; text-decoration: underline; font: bold 24pt Times MS")
            i.setTextInteractionFlags(Qt.TextSelectableByMouse)

        self.subheaders = [QLabel("\nFile"), QLabel("Help"), QLabel("\nSet Destination Path for Output Worklist"),
                           QLabel("Upload SSMINI Data File (.csv)"), QLabel("Enter Target Rack Barcode:"),
                           QLabel("Select Orientation of Transfer from Source to Target"),
                           QLabel("Include Randomized Initials and DOBs for XL LabelPro QC"),
                           QLabel("Submit")]
        for i in self.subheaders:
            i.setStyleSheet("color: black; font: bold 18pt Times MS")
            i.setTextInteractionFlags(Qt.TextSelectableByMouse)

        self.subsubheaders = [QLabel("Set Output Destination"), QLabel("Open SSMINI File"), QLabel("Exit"),
                              QLabel("View Help"), QLabel("About")]
        for i in self.subsubheaders:
            i.setStyleSheet("color: black; font: bold italic 12pt Times MS")
            i.setTextInteractionFlags(Qt.TextSelectableByMouse)

        self.bodies = [QLabel("Tube Sorter Worklist Generator can take an input CSV (comma-separated values) file for a single 96-tube matrix rack from a SampleScan Mini (SSMINI), and output a CSV worklist for the BioMicroLab XL20. The requirements are as follows:\n\n"
                              "1. The input file must be an unaltered CSV file outputted from the SSMINI directly.\n\n"
                              "\t-The file must contain only the 4 default fields (Tube Rack ID, Tube Position, Tube 2D Barcode, Status Code) for each row.\n"
                              "\t-The number of rows cannot exceed 96 (one full matrix rack). However, the number can be less than 96, and worklists for partial matrix racks can be generated. Although the SSMINI will always\n"
							  "\t output 96 rows containing data for every single position on the matrix rack, the positions that have no tube present will be excluded when the worklist is generated.\n"
                              "\t-There cannot be multiple Tube Rack IDs.\n"
                              "\t-There cannot be Tube Positions outside of the A01-H12 range, and each Tube Position must be unique.\n"
                              "\t-Each row should appear as follows:\n\n"
                              "\t(Tube Rack ID),(Tube Position),(Tube 2D Barcode),(Status Code)\n\n"
                              "2. The worklist settings for the BioMicroLab XL20 must be configured correctly for the output worklist to be usable.\n\n"
                              "\t-Worklist File Extension must be set to .CSV.\n"
                              "\t-Skip first N rows of input file must be set to 1 under Ignore Rows.\n"
                              "\t-Field Delimiter must be set to COMMA under Delimiters.\n"
                              "\t-Items must be assigned by default as follows: Source Rack ID -> Column 1; Source Position -> Column 2; Expected Barcode -> Column 3; Target Rack ID -> Column 4; Target Position -> Column 5.\n\n"
                              "3. The Target Rack ID must be exactly as it appears on the matrix rack (alphanumeric without any special characters). Tubes can be returned back to the source rack by assigning the Target Rack ID to be the same as the Source Rack ID.\n\n"
                              "The output worklist will be a CSV file, and can include randomized initials and DOBs for use with the XL LabelPro if the option is selected. Each row should appear as follows:\n\n"
                              "Without randomized initials and DOBs:\n\n"
                              "\t(SourceRackID),(SourcePosition),(SampleBarcode),(DestinationRackID),(DestinationPosition)\n\n"
                              "With randomized initials and DOBs:\n\n"
                              "\t(SourceRackID),(SourcePosition),(SampleBarcode),(DestinationRackID),(DestinationPosition),(SampleBarcode),(DateOfBirth),(Initials)\n\n"
                              "Consult the respective SSMINI, BioMicroLab XL20, and XL LabelPro manuals for instructions on operating the instruments."),
                        QLabel("The Menu bar contains the File and Help options."),
                        QLabel("The File menu contains the Set Output Destination, Open SSMINI File, and Exit options."),
                        QLabel("See the Set Destination Path for Output Worklist section."),
                        QLabel("See the Upload SSMINI Data File section."),
                        QLabel("Exits the application."),
                        QLabel("The Help menu contains the View Help and About options."),
                        QLabel("Opens the Application Help (current window)."),
                        QLabel("Displays information about the application."),
                        QLabel("The main interface allows for setting the output destination, uploading the SSMINI file, entering the desired target rack barcode, selecting the orientation of tube transfer, and generating the output worklist."),
                        QLabel("Define the destination on the user's system where the ouput worklist will be generated. The destination path can be entered into the field manually. The Browse button opens a dialog for selecting the destination folder."),
                        QLabel("Select the SSMINI CSV file on the user's system to be processed. The path for the file can be entered into the field manually. The Browse button opens a dialog for selecting the file."),
                        QLabel("Enter into the field manually or scan in the alphanumeric target rack barcode that will be included in the worklist. To return tubes back to the source rack, enter the Source Rack ID. If Target Rack ID = Source Rack ID, the orientation of transfer will default to No Flip (flipping will lead to collisions on the XL20 due to pre-existing tubes at Destination Positions)."),
                        QLabel("Select how the tubes will be transferred from the source rack to the target rack. There are 4 options as follows:\n\n"
                               "-No Flip (default): Compared to the source rack's tube positions, the positions on the target rack will be exactly the same. For example, tube at source A1 will be transferred to target A1, tube at source H12 will be transferred to H12, etc.\n\n"
                               "-Flip Vertically: Compared to the source rack's tube positions, the positions on the target rack will be reflected over an imaginary horizontal plane between rows D and E that bisect the matrix rack. For example, tube at source A1 will be transferred to target H1, tube at source H12 will be transferred to target H1, etc.\n\n"
                               "-Flip Horizontally: Compared to the source rack's tube positions, the positions on the target rack will be reflected over an imaginary vertical plane between columns 6 and 7 that bisect the matrix rack. For example, tube at source A1 will be transferred to target A12, tube at source H12 will be transferred to target H1, etc.\n\n"
                               "-Flip Diagonally: Compared to the source rack's tube positions, the positions on the target rack will be reflected twice over both imaginary planes described in the previous 2 options. For example, tube at source A1 will be transferred to target H12, tube at source H12 will be transferred to target A1, etc.\n\n"
                               "Note: If the entered Target Rack ID is the same as the Source Rack ID, only the No Flip option will be usable to avoid potential collisions on the XL20."),
                        QLabel("The checkbox is unchecked by default. Checking the checkbox will add 3 additional columns (Sample Barcode, Date of Birth, and Initials) to the worklist which are required for creating label projects for the XL LabelPro. The Sample Barcode will be copied from the Tube 2D Barcode field of the SSMINI file. The Date of Birth and Initials will be randomized each time."),
                        QLabel("Submit the SSMINI file and entered information to generate the worklist. If the attempt is successful, the message 'Worklist successfully created at (Destination Path)!' will be displayed in blue. If the attempt is unsuccessful, an error message will be displayed in red. Refer to the Troubleshooting section for the list of errors and possible solutions."),
                        QLabel("If a red error message is displayed upon clicking the Submit button, consult the table below:")]
        for i in self.bodies:
            i.setWordWrap(True)
            i.setTextInteractionFlags(Qt.TextSelectableByMouse)

        bold_font = QFont("Times", 12, QFont.Bold)

        self.troubleshooting_table = QTableWidget()
        self.troubleshooting_table.setRowCount(13)
        self.troubleshooting_table.setColumnCount(2)
        self.troubleshooting_table.setColumnWidth(0, 300)
        self.troubleshooting_table.setColumnWidth(1, 714)
        self.troubleshooting_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.troubleshooting_table.horizontalHeader().hide()
        self.troubleshooting_table.verticalHeader().hide()
        self.troubleshooting_table.setMinimumSize(1000, 490)
        self.troubleshooting_table.setStyleSheet("background-color: #7FDAFF")
        self.troubleshooting_table.setItem(0, 0, QTableWidgetItem("Error Message"))
        self.troubleshooting_table.item(0, 0).setFont(bold_font)
        self.troubleshooting_table.setItem(0, 1, QTableWidgetItem("Description and Solution"))
        self.troubleshooting_table.item(0, 1).setFont(bold_font)
        self.troubleshooting_table.setItem(1, 0, QTableWidgetItem("No destination path specified!"))
        self.troubleshooting_table.setItem(1, 1, QTableWidgetItem("The Set Destination Path for Output Worklist field is empty. Input manually or click on the Browse button to specify a destination path, and retry."))
        self.troubleshooting_table.setItem(2, 0, QTableWidgetItem("Specified destination path is invalid!"))
        self.troubleshooting_table.setItem(2, 1, QTableWidgetItem("The destination path specified in the Set Destination Path for Output Worklist field does not exist on the user's system. Specify a valid destination path on the user's system, and retry."))
        self.troubleshooting_table.setItem(3, 0, QTableWidgetItem("No data file specified!"))
        self.troubleshooting_table.setItem(3, 1, QTableWidgetItem("The Select SSMINI Data File field is empty. Input manually or click on the Browse button to specify a SSMINI data file, and retry."))
        self.troubleshooting_table.setItem(4, 0, QTableWidgetItem("Specified data file is invalid! File does not exist at directory."))
        self.troubleshooting_table.setItem(4, 1, QTableWidgetItem("The file specified in the Select SSMINI Data File field does not exist on the user's system. Specify a valid file on the user's system, and retry."))
        self.troubleshooting_table.setItem(5, 0, QTableWidgetItem("Specified data file is invalid! File is not a CSV file."))
        self.troubleshooting_table.setItem(5, 1, QTableWidgetItem("The file specified in the Select SSMINI Data File field is not a CSV file. Only CSV files obtained directly from the SSMINI can be processed. Either change the file extension or re-run the SSMINI to output a CSV file, and retry."))
        self.troubleshooting_table.setItem(6, 0, QTableWidgetItem("No target rack barcode specified!"))
        self.troubleshooting_table.setItem(6, 1, QTableWidgetItem("The Enter Target Rack Barcode field is empty. Input manually or scan in the desired alphanumeric target rack barcode, and retry."))
        self.troubleshooting_table.setItem(7, 0, QTableWidgetItem("Specified target rack barcode is invalid! Only alphanumeric characters are allowed."))
        self.troubleshooting_table.setItem(7, 1, QTableWidgetItem("The target rack barcode specified in the Enter Target Rack Barcode field contains special characters. Matrix racks can only have alphanumeric barcodes. Specify an alphanumeric target rack barcode, and retry."))
        self.troubleshooting_table.setItem(8, 0, QTableWidgetItem("Specified data file is invalid! Number of rows and/or columns exceeds the limits of 96 rows and 4 columns."))
        self.troubleshooting_table.setItem(8, 1, QTableWidgetItem("The file specified in the Select SSMINI Data File field contains more than 96 rows and/or 4 columns which exceeds the limits for a SSMINI data file generated from a single 96-tube matrix rack with default fields. Verify that the file is an unedited file outputted directly from the SSMINI with the default fields as described in the Overview section. If necessary, re-run the SSMINI to obtain a new file, and retry."))
        self.troubleshooting_table.setItem(9, 0, QTableWidgetItem("Specified data file is invalid! Multiple source rack barcodes detected."))
        self.troubleshooting_table.setItem(9, 1, QTableWidgetItem("The file specified in the Select SSMINI Data File field contains multiple source rack barcodes. An unedited SSMINI data file should only contain a single source rack barcode since the SSMINI is not designed to scan multiple matrix racks at the same time. Verify that the file is an unedited file outputted directly from the SSMINI since multiple source barcodes should not be possible if a single matrix rack was scanned. If necessary, re-run the SSMINI to obtain a new file, and retry."))
        self.troubleshooting_table.setItem(10, 0, QTableWidgetItem("Specified data file is invalid! Source positions can only range from A(01-12)-H(01-12)."))
        self.troubleshooting_table.setItem(10, 1, QTableWidgetItem("The file specified in the Select SSMINI Data File field contains source positions outside of the A(01-12)-H(01-12) range. A 96-tube matrix rack cannot have source positions outside of the aforementioned range. Verify that the file is an unedited file outputted directly from the SSMINI since invalid source positions should not be generated if a 96-tube matrix rack was scanned. If necessary, re-run the SSMINI to obtain a new file, and retry."))
        self.troubleshooting_table.setItem(11, 0, QTableWidgetItem("Specified data file is invalid! Duplicate source positions detected."))
        self.troubleshooting_table.setItem(11, 1, QTableWidgetItem("The file specified in the Select SSMINI Data File field contains duplicate source positions. A matrix rack cannot have more than a single tube at any source position. Verify that the file is an unedited file outputted directly from the SSMINI since source positions should not be duplicated if a single matrix rack was scanned. If necessary, re-run the SSMINI to obtain a new file, and retry."))
        self.troubleshooting_table.setItem(12, 0, QTableWidgetItem("Source Rack ID is the same as the Target Rack ID! Flipping is not possible."))
        self.troubleshooting_table.setItem(12, 1, QTableWidgetItem("The Source Rack ID in the SSMINI file and the specified Destination Rack ID from the Enter Target Rack Barcode field are the same which means that tubes will be returned back to the source rack. Flipping the orientation of transfer is prohibited to prevent potential collisions on the BioMicroLab XL20. Either choose the No Flip option or specify a different target rack barcode, and retry."))
        self.troubleshooting_table.resizeRowsToContents()


        self.help_window_central_widget_layout.addWidget(self.headers[0])
        self.help_window_central_widget_layout.addWidget(self.bodies[0])
        self.help_window_central_widget_layout.addWidget(self.headers[1])
        self.help_window_central_widget_layout.addWidget(self.bodies[1])
        self.help_window_central_widget_layout.addWidget(self.subheaders[0])
        self.help_window_central_widget_layout.addWidget(self.bodies[2])
        self.help_window_central_widget_layout.addWidget(self.subsubheaders[0])
        self.help_window_central_widget_layout.addWidget(self.bodies[3])
        self.help_window_central_widget_layout.addWidget(self.subsubheaders[1])
        self.help_window_central_widget_layout.addWidget(self.bodies[4])
        self.help_window_central_widget_layout.addWidget(self.subsubheaders[2])
        self.help_window_central_widget_layout.addWidget(self.bodies[5])
        self.help_window_central_widget_layout.addWidget(self.subheaders[1])
        self.help_window_central_widget_layout.addWidget(self.bodies[6])
        self.help_window_central_widget_layout.addWidget(self.subsubheaders[3])
        self.help_window_central_widget_layout.addWidget(self.bodies[7])
        self.help_window_central_widget_layout.addWidget(self.subsubheaders[4])
        self.help_window_central_widget_layout.addWidget(self.bodies[8])
        self.help_window_central_widget_layout.addWidget(self.headers[2])
        self.help_window_central_widget_layout.addWidget(self.bodies[9])
        self.help_window_central_widget_layout.addWidget(self.subheaders[2])
        self.help_window_central_widget_layout.addWidget(self.bodies[10])
        self.help_window_central_widget_layout.addWidget(self.subheaders[3])
        self.help_window_central_widget_layout.addWidget(self.bodies[11])
        self.help_window_central_widget_layout.addWidget(self.subheaders[4])
        self.help_window_central_widget_layout.addWidget(self.bodies[12])
        self.help_window_central_widget_layout.addWidget(self.subheaders[5])
        self.help_window_central_widget_layout.addWidget(self.bodies[13])
        self.help_window_central_widget_layout.addWidget(self.subheaders[6])
        self.help_window_central_widget_layout.addWidget(self.bodies[14])
        self.help_window_central_widget_layout.addWidget(self.subheaders[7])
        self.help_window_central_widget_layout.addWidget(self.bodies[15])
        self.help_window_central_widget_layout.addWidget(self.headers[3])
        self.help_window_central_widget_layout.addWidget(self.bodies[16])
        self.help_window_central_widget_layout.addWidget(self.troubleshooting_table)

        self.help_window_central_widget.setLayout(self.help_window_central_widget_layout)
        self.help_window.show()


class Application(QWidget):
    def __init__(self):
        super().__init__()
        self.CreateApp()

    def CreateApp(self):
        main_layout = QVBoxLayout()
        main_layout.setAlignment(Qt.AlignCenter)

        set_destination = QWidget()
        set_destination_layout = QHBoxLayout()
        set_destination.setLayout(set_destination_layout)

        file_upload = QWidget()
        file_upload_layout = QHBoxLayout()
        file_upload.setLayout(file_upload_layout)

        enter_target_rack = QWidget()
        enter_target_rack_layout = QHBoxLayout()
        enter_target_rack.setLayout(enter_target_rack_layout)

        choose_transfer_orientation = QWidget()
        choose_transfer_orientation_layout = QHBoxLayout()
        choose_transfer_orientation.setLayout(choose_transfer_orientation_layout)

        set_destination_label = QLabel("Set Destination Path for Output Worklist:")
        set_destination_label.setTextInteractionFlags(Qt.TextSelectableByMouse)
        self.set_destination_line_edit = QLineEdit()
        set_destination_browse = QPushButton("Browse")
        set_destination_browse.clicked.connect(self.SetDestinationPath)
        set_destination_button.triggered.connect(self.SetDestinationPath)

        set_destination_layout.addWidget(set_destination_label)
        set_destination_layout.addWidget(self.set_destination_line_edit)
        set_destination_layout.addWidget(set_destination_browse)

        upload_label = QLabel("Select SSMINI Data File (.csv):")
        upload_label.setTextInteractionFlags(Qt.TextSelectableByMouse)
        self.upload_line_edit = QLineEdit()
        upload_browse = QPushButton("Browse")
        upload_browse.clicked.connect(self.UploadSSMINIDataFile)
        open_button.triggered.connect(self.UploadSSMINIDataFile)

        enter_target_rack_label = QLabel("Enter Target Rack Barcode:")
        enter_target_rack_label.setTextInteractionFlags(Qt.TextSelectableByMouse)
        self.enter_target_rack_edit = QLineEdit()

        choose_transfer_orientation_label = QLabel("Select Orientation of Transfer from Source to Target:")
        choose_transfer_orientation_label.setTextInteractionFlags(Qt.TextSelectableByMouse)
        global choose_transfer_orientation_combo_box
        choose_transfer_orientation_combo_box = QComboBox()
        choose_transfer_orientation_combo_box.setMinimumSize(450, 25)
        choose_transfer_orientation_combo_box.addItem("No Flip (A1 -> A1; H12 -> H12)")
        choose_transfer_orientation_combo_box.addItem("Flip Vertically (A1 -> H1; H12 -> A12)")
        choose_transfer_orientation_combo_box.addItem("Flip Horizontally (A1 -> A12; H12 -> H1)")
        choose_transfer_orientation_combo_box.addItem("Flip Diagonally (A1 -> H12; H12 -> A1)")

        self.add_dob_initials_checkbox = QCheckBox("Include Randomized Initials and DOBs for XL LabelPro QC")
        global checkbox_state
        checkbox_state = False
        self.add_dob_initials_checkbox.stateChanged.connect(self.CheckBoxforIncludingDOBsandInitials)

        upload_submit = QPushButton("Submit", self)
        upload_submit.clicked.connect(self.CreateWorklist)

        self.submit_result = QLabel("")
        self.submit_result.setTextInteractionFlags(Qt.TextSelectableByMouse)

        file_upload_layout.addWidget(upload_label)
        file_upload_layout.addWidget(self.upload_line_edit)
        file_upload_layout.addWidget(upload_browse)

        enter_target_rack_layout.addWidget(enter_target_rack_label)
        enter_target_rack_layout.addWidget(self.enter_target_rack_edit)

        choose_transfer_orientation_layout.addWidget(choose_transfer_orientation_label)
        choose_transfer_orientation_layout.addWidget(choose_transfer_orientation_combo_box, 0, Qt.AlignLeft)

        main_layout.addWidget(set_destination)
        main_layout.addWidget(file_upload)
        main_layout.addWidget(enter_target_rack)
        main_layout.addWidget(choose_transfer_orientation)
        main_layout.addWidget(self.add_dob_initials_checkbox, 0, Qt.AlignCenter)
        main_layout.addWidget(upload_submit, 0, Qt.AlignCenter)
        upload_submit.setMinimumSize(150, 40)
        main_layout.addWidget(self.submit_result, 1, Qt.AlignCenter)
        self.submit_result.setStyleSheet("color: red; font: bold 12pt Times MS")

        self.setLayout(main_layout)
        self.show()

    def SetDestinationPath(self):
        destinationName = QFileDialog.getExistingDirectory(self, "Select Destination Path for Worklist File")
        if destinationName:
            self.set_destination_line_edit.setText(destinationName)

    def UploadSSMINIDataFile(self):
        fileName, _ = QFileDialog.getOpenFileName(self, "Choose SSMINI Output CSV File for Source Rack", "",
                                                  "CSV Files (*.csv);;All Files (*)")
        if fileName:
            self.upload_line_edit.setText(fileName)

    def CheckBoxforIncludingDOBsandInitials(self, state):
        if state == Qt.Checked:
            global checkbox_state
            checkbox_state = True
        else:
            checkbox_state = False

    def GenerateRandomDate(self):
        month = random.randrange(1, 12)
        day = random.randrange(1, 31)
        year = str(random.randrange(0, 9)) + str(random.randrange(0, 9))
        leap_years = ["00", "04", "08", "12", "16", "20", "24", "28", "32", "36", "40", "44", "48", "52", "56", "60",
                      "64", "68", "72", "76", "80", "84", "88", "92", "96"]
        thirty_day_months = [4, 6, 9, 11]

        if month == 2 and day > 29:
            self.GenerateRandomDate()
        elif month == 2 and day == 29 and year not in leap_years:
            self.GenerateRandomDate()
        elif month in thirty_day_months and day == 31:
            self.GenerateRandomDate()
        else:
            global randomized_date
            randomized_date = str(month) + "/" + str(day) + "/" + year

    def CreateWorklist(self):
        destinationName = self.set_destination_line_edit.text()
        fileName = self.upload_line_edit.text()
        targetrackBarcode = self.enter_target_rack_edit.text()

        if destinationName == "":
            self.submit_result.setStyleSheet("color: red; font: bold 12pt Times MS")
            self.submit_result.setText("No destination path specified!")
        elif os.path.isdir(destinationName) == False:
            self.submit_result.setStyleSheet("color: red; font: bold 12pt Times MS")
            self.submit_result.setText("Specified destination path is invalid!")
        else:
            if fileName == "":
                self.submit_result.setStyleSheet("color: red; font: bold 12pt Times MS")
                self.submit_result.setText("No data file specified!")
            elif os.path.isfile(fileName) == False:
                self.submit_result.setStyleSheet("color: red; font: bold 12pt Times MS")
                self.submit_result.setText("Specified data file is invalid! File does not exist at directory.")
            elif os.path.splitext(fileName)[1] != ".csv":
                self.submit_result.setStyleSheet("color: red; font: bold 12pt Times MS")
                self.submit_result.setText("Specified data file is invalid! File is not a CSV file.")
            elif self.enter_target_rack_edit.text() == "":
                self.submit_result.setStyleSheet("color: red; font: bold 12pt Times MS")
                self.submit_result.setText("No target rack barcode specified!")
            elif (self.enter_target_rack_edit.text()).isalnum() == False:
                self.submit_result.setStyleSheet("color: red; font: bold 12pt Times MS")
                self.submit_result.setText("Specified target rack barcode is invalid! Only alphanumeric characters are allowed.")
            else:
                reader = csv.reader(open(fileName, "r"))
                num_rows = 0
                num_cols = []
                for row in reader:
                    num_rows += 1
                    columns = len(row)
                    num_cols.append(columns)

                if num_rows > 96 or max(num_cols) > 4:
                    self.submit_result.setStyleSheet("color: red; font: bold 12pt Times MS")
                    self.submit_result.setMinimumWidth(750)
                    self.submit_result.setText("Specified data file is invalid! Number of rows and/or columns exceeds the limits of 96 rows and 4 columns.")
                    self.submit_result.setWordWrap(True)
                else:
                    df = pandas.read_csv(fileName, header=None, dtype=object, engine="python")
                    df.columns = ["Tube Rack ID", "Tube Position", "Tube 2D Barcode", "Status Code"]
                    df = df[df["Status Code"] == "0"]
                    df = df[df["Tube 2D Barcode"] == df["Tube 2D Barcode"]]

                    source_rack_id = (df.iloc[0:96, 0]).values
                    if len(set(source_rack_id)) > 1:
                        self.submit_result.setStyleSheet("color: red; font: bold 12pt Times MS")
                        self.submit_result.setText("Specified data file is invalid! Multiple source rack barcodes detected.")
                    else:
                        source_position = (df.iloc[0:96, 1]).values

                        allowed_source_positions = ["A01", "A02", "A03", "A04", "A05", "A06", "A07", "A08", "A09", "A10", "A11", "A12",
                                                    "B01", "B02", "B03", "B04", "B05", "B06", "B07", "B08", "B09", "B10", "B11", "B12",
                                                    "C01", "C02", "C03", "C04", "C05", "C06", "C07", "C08", "C09", "C10", "C11", "C12",
                                                    "D01", "D02", "D03", "D04", "D05", "D06", "D07", "D08", "D09", "D10", "D11", "D12",
                                                    "E01", "E02", "E03", "E04", "E05", "E06", "E07", "E08", "E09", "E10", "E11", "E12",
                                                    "F01", "F02", "F03", "F04", "F05", "F06", "F07", "F08", "F09", "F10", "F11", "F12",
                                                    "G01", "G02", "G03", "G04", "G05", "G06", "G07", "G08", "G09", "G10", "G11", "G12",
                                                    "H01", "H02", "H03", "H04", "H05", "H06", "H07", "H08", "H09", "H10", "H11", "H12"]

                        for i in source_position:
                            if i in allowed_source_positions:
                                global source_position_check1
                                source_position_check1 = True
                            else:
                                source_position_check1 = False
                                break

                        if len(set(source_position)) == len(source_position):
                            source_position_check2 = True
                        else:
                            source_position_check2 = False

                        if source_position_check1 == False:
                            self.submit_result.setStyleSheet("color: red; font: bold 12pt Times MS")
                            self.submit_result.setText("Specified data file is invalid! Source positions can only range from A(01-12)-H(01-12).")
                        elif source_position_check2 == False:
                            self.submit_result.setStyleSheet("color: red; font: bold 12pt Times MS")
                            self.submit_result.setText("Specified data file is invalid! Duplicate source positions detected.")
                        else:
                            sample_barcode = (df.iloc[0:96, 2]).values
                            destination_rack_id = []
                            if choose_transfer_orientation_combo_box.currentText() == "No Flip (A1 -> A1; H12 -> H12)":
                                destination_position = source_position
                            elif choose_transfer_orientation_combo_box.currentText() == "Flip Vertically (A1 -> H1; H12 -> A12)":
                                if targetrackBarcode == source_rack_id[0]:
                                    self.submit_result.setStyleSheet("color: red; font: bold 12pt Times MS")
                                    self.submit_result.setMinimumWidth(750)
                                    self.submit_result.setText("Source Rack ID is the same as the Target Rack ID! Flipping is not possible.")
                                    self.submit_result.setWordWrap(True)
                                    return None
                                else:
                                    vertical_flip_positions = ["H01", "H02", "H03", "H04", "H05", "H06", "H07", "H08", "H09", "H10", "H11", "H12",
                                                            "G01", "G02", "G03", "G04", "G05", "G06", "G07", "G08", "G09", "G10", "G11", "G12",
                                                            "F01", "F02", "F03", "F04", "F05", "F06", "F07", "F08", "F09", "F10", "F11", "F12",
                                                            "E01", "E02", "E03", "E04", "E05", "E06", "E07", "E08", "E09", "E10", "E11", "E12",
                                                            "D01", "D02", "D03", "D04", "D05", "D06", "D07", "D08", "D09", "D10", "D11", "D12",
                                                            "C01", "C02", "C03", "C04", "C05", "C06", "C07", "C08", "C09", "C10", "C11", "C12",
                                                            "B01", "B02", "B03", "B04", "B05", "B06", "B07", "B08", "B09", "B10", "B11", "B12",
                                                            "A01", "A02", "A03", "A04", "A05", "A06", "A07", "A08", "A09", "A10", "A11", "A12",]

                                    destination_position = []
                                    for i in source_position:
                                        vertical_index = allowed_source_positions.index(i)
                                        destination_position.extend([vertical_flip_positions[vertical_index]])
                            elif choose_transfer_orientation_combo_box.currentText() == "Flip Horizontally (A1 -> A12; H12 -> H1)":
                                if targetrackBarcode == source_rack_id[0]:
                                    self.submit_result.setStyleSheet("color: red; font: bold 12pt Times MS")
                                    self.submit_result.setMinimumWidth(750)
                                    self.submit_result.setText("Source Rack ID is the same as the Target Rack ID! Flipping is not possible.")
                                    self.submit_result.setWordWrap(True)
                                    return None
                                else:
                                    horizontal_flip_positions = ["A12", "A11", "A10", "A09", "A08", "A07", "A06", "A05", "A04", "A03", "A02", "A01",
                                                                 "B12", "B11", "B10", "B09", "B08", "B07", "B06", "B05", "B04", "B03", "B02", "B01",
                                                                 "C12", "C11", "C10", "C09", "C08", "C07", "C06", "C05", "C04", "C03", "C02", "C01",
                                                                 "D12", "D11", "D10", "D09", "D08", "D07", "D06", "D05", "D04", "D03", "D02", "D01",
                                                                 "E12", "E11", "E10", "E09", "E08", "E07", "E06", "E05", "E04", "E03", "E02", "E01",
                                                                 "F12", "F11", "F10", "F09", "F08", "F07", "F06", "F05", "F04", "F03", "F02", "F01",
                                                                 "G12", "G11", "G10", "G09", "G08", "G07", "G06", "G05", "G04", "G03", "G02", "G01",
                                                                 "H12", "H11", "H10", "H09", "H08", "H07", "H06", "H05", "H04", "H03", "H02", "H01"]

                                    destination_position = []
                                    for i in source_position:
                                        horizontal_index = allowed_source_positions.index(i)
                                        destination_position.extend([horizontal_flip_positions[horizontal_index]])
                            else:
                                if targetrackBarcode == source_rack_id[0]:
                                    self.submit_result.setStyleSheet("color: red; font: bold 12pt Times MS")
                                    self.submit_result.setMinimumWidth(750)
                                    self.submit_result.setText("Source Rack ID is the same as the Target Rack ID! Flipping is not possible.")
                                    self.submit_result.setWordWrap(True)
                                    return None
                                else:
                                    diagonal_flip_positions = ["H12", "H11", "H10", "H09", "H08", "H07", "H06", "H05", "H04", "H03", "H02", "H01",
                                                               "G12", "G11", "G10", "G09", "G08", "G07", "G06", "G05", "G04", "G03", "G02", "G01",
                                                               "F12", "F11", "F10", "F09", "F08", "F07", "F06", "F05", "F04", "F03", "F02", "F01",
                                                               "E12", "E11", "E10", "E09", "E08", "E07", "E06", "E05", "E04", "E03", "E02", "E01",
                                                               "D12", "D11", "D10", "D09", "D08", "D07", "D06", "D05", "D04", "D03", "D02", "D01",
                                                               "C12", "C11", "C10", "C09", "C08", "C07", "C06", "C05", "C04", "C03", "C02", "C01",
                                                               "B12", "B11", "B10", "B09", "B08", "B07", "B06", "B05", "B04", "B03", "B02", "B01",
                                                               "A12", "A11", "A10", "A09", "A08", "A07", "A06", "A05", "A04", "A03", "A02", "A01"]

                                    destination_position = []
                                    for i in source_position:
                                        diagonal_index = allowed_source_positions.index(i)
                                        destination_position.extend([diagonal_flip_positions[diagonal_index]])

                            date_of_birth = []
                            initials = []


                            for i in source_rack_id:
                                destination_rack_id.extend([targetrackBarcode])
                                initials.extend([random.choice("ABCDEFGHIJKLMNOPQRSTUVWXYZ") + "." + random.choice("ABCDEFGHIJKLMNOPQRSTUVWXYZ") + "."])
                                self.GenerateRandomDate()
                                date_of_birth.extend([randomized_date])


                            if checkbox_state == True:
                                worklist = {"SourceRackID": source_rack_id, "SourcePosition": source_position, "SampleBarcode": sample_barcode,
                                            "DestinationRackID": destination_rack_id, "DestinationPosition": destination_position,
                                            "SampleBarcodes": sample_barcode, "DateOfBirth": date_of_birth, "Initials": initials}
                                new_df = pandas.DataFrame(worklist, index=None)
                                new_df.columns = ["SourceRackID", "SourcePosition", "SampleBarcode", "DestinationRackID", "DestinationPosition", "SampleBarcode", "DateofBirth", "Initials"]
                            else:
                                worklist = {"SourceRackID": source_rack_id, "SourcePosition": source_position, "SampleBarcode": sample_barcode,
                                            "DestinationRackID": destination_rack_id, "DestinationPosition": destination_position}
                                new_df = pandas.DataFrame(worklist, index=None)

                            export_csv = new_df.to_csv(destinationName + "/" + source_rack_id[0] + " to " + targetrackBarcode + ".csv",
                                                           index=None, header=True)
                            self.submit_result.setStyleSheet("color: blue; font: bold 12pt Times MS")
                            self.submit_result.setText("Worklist successfully created at " + destinationName + "!")


if __name__ == '__main__':
    app = QApplication(sys.argv)
    app.setStyle('Fusion')
    window = MyMainWindow()
    window.show()
    sys.exit(app.exec())


