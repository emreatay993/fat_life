# region Import libraries
import csv
import context_menu
import clr
clr.AddReference('mscorlib')  # Ensure the core .NET assembly is referenced
from System.IO import StreamWriter, FileStream, FileMode, FileAccess
from System.Text import UTF8Encoding
from System.Diagnostics import Process, ProcessWindowStyle
import os
# endregion

# region Define the solution directory based on the selected solution object
solution_directory_path = sol_selected_environment.WorkingDir[:-1]
solution_directory_path = solution_directory_path.Replace("\\", "\\\\") + r'\\'
# endregion

# region Set the default units of Mechanical as StandardMM and Celsius
ExtAPI.Application.ActiveUnitSystem = MechanicalUnitSystem.StandardNMM
ExtAPI.Application.ActiveMetricTemperatureUnit = MetricTemperatureUnitType.Celsius
# endregion

# region Get the names of all named selections in the tree to select from
list_of_obj_of_NS=DataModel.GetObjectsByType(DataModelObjectCategory.NamedSelection)
list_of_names_of_NS=[list_of_obj_of_NS[i].Name for i in range(len(list_of_obj_of_NS))]

# region Export both mean and alternating VPS results for fatigue postprocessing
DataModel.GetObjectsByName("VPS_Fatigue_Mean")[0].Activate()
file_path_of_VPS_fatigue_mean_txt = solution_directory_path + "VPS_Fatigue_Mean.txt"
DataModel.GetObjectsByName("VPS_Fatigue_Mean")[0].ExportToTextFile(file_path_of_VPS_fatigue_mean_txt)

DataModel.GetObjectsByName("VPS_Fatigue_Alternating_Stress")[0].Activate()
file_path_of_VPS_fatigue_alternating_stress_txt = solution_directory_path + "VPS_Fatigue_Alternating_Stress.txt"
DataModel.GetObjectsByName("VPS_Fatigue_Alternating_Stress")[0].ExportToTextFile(file_path_of_VPS_fatigue_alternating_stress_txt)
# endregion

# region Export temperature distribution for fatigue postprocessing
DataModel.GetObjectsByName("Fatigue_Temperature")[0].Activate()
file_path_of_VPS_fatigue_temperature_txt = solution_directory_path + "Fatigue_Temperature.txt"
DataModel.GetObjectsByName("Fatigue_Temperature")[0].ExportToTextFile(file_path_of_VPS_fatigue_temperature_txt)
# endregion

cpython_script_name = "harmonic_fatigue_loading_history_generator_v0.py"
cpython_script_path = sol_selected_environment.WorkingDir + cpython_script_name

cpython_code = """
import pandas as pd
import os
import sys
from PyQt5.QtWidgets import QApplication, QMainWindow, QVBoxLayout, QWidget, QLabel, QLineEdit, QComboBox, QPushButton, QHBoxLayout
from PyQt5.QtGui import QFont
from PyQt5.QtCore import Qt

# region Define material list
material_dropdown_list_items = [
    '13-8PH AMS5629 Bar H1150',
    '15-5PH AMS5659 Bar H1150',
    '15-5PH AMS5659 Forging H1150',
    '17-4 PH EN5643 Bar H1150',
    '17-4PH TMS01-030669 Forging H1150',
    '17-7PH AMS5644 Bar TH1050',
    '2024 AMS QQ A 225-6 Bar T851',
    '2024 AMS QQ A 250-4 Plate T851',
    '2124 AMS4101 Sheet T851',
    '2219 AMS QQ A 250-30 Plate T851',
    '33CrMoV12-9 AMS6481 Bar 41-45 HRC',
    '40CrMoV13-9 EN3971 Bar 41-45 HRC',
    '40CrMoV13-9 TMS01013400 Bar Quenched and Tempered 40-45 HRC',
    '6061 AMS4027 Sheet T6',
    '6061 AMS4127 Forging T6',
    '6061 AMS4150 Bar T6',
    '7075 AMS4123 Bar Solution and Precipitation Heat Treated',
    '718+ TMS01-030772 Forging Solution and Precipitation Heat Treated',
    'A286 AMS5732 Bar Solution and Precipitation Heat Treated',
    'A286 AMS5737 Sheet Solution and Precipitation Heat Treated',
    'A286 AMS5737 Forging Solution and Precipitation Heat Treated',
    'A357 TMS06-042429 Casting Solution and Precipitation Heat Treated T6',
    'AISI316 AMS5648 Bar Solution Heat Treated',
    'AISI321 AMS5645 Bar Solution Heat Treated',
    'AISI347 AMS5646 Bar Solution Heat Treated',
    'AISI420 AMS5621 Bar 45-46 HRC',
    'AISI420 AMS5621 Bar 47-48 HRC',
    'AISI420 ASTM A276 Bar 45-46 HRC',
    'AISI4340 AMS6414 Bar 33-38 HRC',
    'AISI9310 AMS6265 Bar Annealed',
    'AISI9310 AMS6265 Bar As Quenched',
    'Alloy 188 AMS5608 Sheet Solution Heat Treated',
    'Alloy 188 TMS01-031422 Forging Solution Heat Treated',
    'Alloy 230 AMS5878 Sheet Solution Heat Treated',
    'Alloy 230 AMS5891 Forging Solution Heat Treated',
    'Alloy 247LC TMS-06-003 DS Casting Solution and Precipitation Heat Treated',
    'Alloy 247LC TMS06-031379 Casting Primary and Secondary Precipitation Heat Treated',
    'Alloy 25 AMS5537 Sheet Solution Heat Treated',
    'Alloy 263 AMS5887 Sheet Solution and Precipitation Heat Treated',
    'Alloy 263 AMS5886 Forging Solution and Precipitation Heat Treated',
    'Alloy 600 AMS5665 Bar Annealed',
    'Alloy 600 AMS5665 Forging Annealed',
    'Alloy 625 AMS5599 Sheet Annealed',
    'Alloy 625 AMS5666 Bar Annealed',
    'Alloy 625 TMS01-031452 Forging Annealed',
    'Alloy 625 TMS08-048100 AdditiveXY As Built',
    'Alloy 625 TMS08-048100 AdditiveXY Stress Relieved and Solution Heat Treated',
    'Alloy 625 TMS08-048100 AdditiveZ As Built',
    'Alloy 625 TMS08-048100 AdditiveZ Stress Relieved and Solution Heat Treated',
    'Alloy 713LC TMS06-047729 Casting As Cast',
    'Alloy 718 AMS5596 Sheet Solution and Precipitation Heat Treated',
    'Alloy 718 AMS5663 Bar Solution and Precipitation Heat Treated',
    'Alloy 718 ASTM F3055 AdditiveXY Solution and Precipitation Heat Treated',
    'Alloy 718 ASTM F3055 AdditiveZ Solution and Precipitation Heat Treated',
    'Alloy 718 TMS01-030875 Forging Solution and Precipitation Heat Treated',
    'Alloy 718 TMS06-030653 Casting Homegization Solution and Precipitation Heat Treated',
    'Alloy 718 TMS08-048090 AdditiveXY Solution and Precipitation Heat Treated',
    'Alloy 718 TMS08-048090 AdditiveZ Solution and Precipitation Heat Treated',
    'Alloy 720 TMS01-030759 Forging Solution and Heat Treated',
    'Alloy 730 TMS01-030773 Forging Solution and Heat Treated',
    'Alloy 738LC TMS06-031399 Casting Solution and Precipitation Heat Treated',
    'Alloy 909 TMS01-030670 Forging Re-Solution and Precipitation Heat Treated',
    'Alloy 939 TMS06-032675 Casting Solution Primary and Secondary Precipitation Heat Treated',
    'Alloy 939 TMS08-048092 AdditiveXY Solution and Precipitation Heat Treated',
    'Alloy 939 TMS08-048092 AdditiveZ Solution and Precipitation Heat Treated',
    'Alloy X AMS5754 Bar Solution Heat Treated',
    'Alloy X TMS-08-002 AdditiveXY As Built',
    'Alloy X TMS-08-002 AdditiveZ As Built',
    'Alloy X TMS01-031440 Forging Solution Heat Treated',
    'Alloy X AMS5536 Sheet Solution Heat Treated',
    'AlSi10Mg TMS08-066332 AdditiveXY Stress Relieved and T6 Heat Treated',
    'AlSi10Mg TMS08-066332 AdditiveZ Stress Relieved and T6 Heat Treated',
    'AM355 AMS5744 Bar SCT1000',
    'AM355 AMS5744 Forging SCT1000',
    'H11 AMS6487 Bar 43-46HRC',
    'HastelloyX TMS08-048101 AdditiveXY Stress Relieved and Solution Heat Treated',
    'Haynes188 AMS5772 Bar Solution Heat Treated',
    'HK30 ASTM A351 Casting As Cast',
    'Jethe M152 AMS5719 Forging Hardened and Double Tempered',
    'M50 AMS6491 Bar Hardened and Tempered AMS6491',
    'MarAging250 AMS6512 Bar Maraged',
    'MarM247LC TMS-06-005 Casting HIP and Solution and Precipitation Heat Treated',
    'Nimonic90 BS HR 502 Bar Solution and Precipitation Heat Treated',
    'SX-4 TMS06-047787 Casting Solution Primary and Secondary Precipitation Heat Treated',
    'Ti6242 AMS4975 Bar Solution and Precipitation Heat Treated',
    'Ti6242 TMS01-030887 Forging Solution and Precipitation Heat Treated',
    'Ti6242 TMS616 Forging Solution and Precipitation Heat Treated',
    'Ti64 AMS4928 Bar Annealed',
    'Ti64 TMS01-030882 Forging Annealed',
    'Ti64 TMS06-033017 Casting HIP and Annealed',
    'Waspaloy TMS01-031441 Forging Solution Stabilization and Precipitation Heat Treated'
]
# endregion 

# region Define the class for main GUI
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        
        self.setWindowTitle('TLIFE - Loading History Generator (Constant Amplitude Load with Mean Stress)')
        self.setGeometry(100, 100, 1500, 300)
        self.setStyleSheet("background-color: #F0F0F0;")

        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)
        
        self.layout = QVBoxLayout(self.central_widget)

        # Info Icon (as text)
        self.info_icon_label = QLabel('Help(?)')
        self.info_icon_label.setFont(QFont('Arial', 14))
        
        self.tooltip1_text = '''
        Ensure that following conditions are checked to make sure that the code works properly:
        
        - There should be VPS_Fatigue_Alternating_Stress, VPS_Fatigue_Mean and Fatigue_Temperature
        result objects in the tree to extract the results from.
        
        - For each results object, there must not be more than one object with the same name in the tree.
        
        - All these objects should be scoped to the same named selection.
        
        - From File/Options/Export, following options should be set to Yes
            - Include Node Numbers
            - Include Node Location
            - Show Tensor Components
        '''
        
        self.info_icon_label.setToolTip(self.tooltip1_text)
        self.layout.addWidget(self.info_icon_label, alignment=Qt.AlignRight)

        # Mission Name
        self.mission_name_label = QLabel('Mission Name:')
        self.mission_name_label.setFont(QFont('Arial', 12))
        self.layout.addWidget(self.mission_name_label)
        self.mission_name_input = QLineEdit(self)
        self.mission_name_input.setFont(QFont('Arial', 12))
        self.mission_name_input.setStyleSheet("background-color: white; padding: 5px;")
        self.layout.addWidget(self.mission_name_input)

        self.layout.addSpacing(10)

        # Named Selection
        self.named_selection_label = QLabel('Named Selection:')
        self.named_selection_label.setFont(QFont('Arial', 12))
        self.layout.addWidget(self.named_selection_label)
        self.named_selection_dropdown = QComboBox(self)
        self.named_selection_dropdown.addItems("""+ str(list_of_names_of_NS) +""")
        self.named_selection_dropdown.setFont(QFont('Arial', 12))
        self.named_selection_dropdown.setStyleSheet("background-color: white; padding: 5px;")
        self.named_selection_dropdown.setEditable(False)
        self.layout.addWidget(self.named_selection_dropdown)

        self.layout.addSpacing(10)

        # Material Name
        self.material_name_label = QLabel('Material Name:')
        self.material_name_label.setFont(QFont('Arial', 12))
        self.layout.addWidget(self.material_name_label)
        self.material_name_dropdown = QComboBox(self)
        self.material_name_dropdown.addItems(material_dropdown_list_items)
        self.material_name_dropdown.setFont(QFont('Arial', 12))
        self.material_name_dropdown.setStyleSheet("background-color: white; padding: 5px;")
        self.material_name_dropdown.setEditable(False)
        self.layout.addWidget(self.material_name_dropdown)

        self.layout.addSpacing(20)

        # Submit Button
        self.submit_button = QPushButton('Submit')
        self.submit_button.setFont(QFont('Arial', 12))
        self.submit_button.setStyleSheet("background-color: #87CEFA; color: black; padding: 10px;")
        self.layout.addWidget(self.submit_button)
        self.submit_button.clicked.connect(self.submit)

        self.center()
        self.show()

    def submit(self):
        mission_name = self.mission_name_input.text()
        named_selection_name = self.named_selection_dropdown.currentText()
        material_name = self.material_name_dropdown.currentText()

        self.close()

        generate_csv_files(mission_name, named_selection_name, material_name)
        
    def center(self):
        qr = self.frameGeometry()
        cp = QApplication.desktop().availableGeometry().center()
        qr.moveCenter(cp)
        self.move(qr.topLeft())
# endregion

# region Define the main file generation routine
def generate_csv_files(mission_name, named_selection_name, material_name):
    # Define file paths
    mean_file_path = os.path.join('""" + solution_directory_path + """', 'VPS_Fatigue_Mean.txt')
    alternating_file_path = os.path.join('""" + solution_directory_path + """', 'VPS_Fatigue_Alternating_Stress.txt')
    fatigue_temperature_file_path = os.path.join('""" + solution_directory_path + """', 'Fatigue_Temperature.txt')

    # Columns to exclude from numerical operations
    exclude_cols = ["Node Number", "X Location (mm)", "Y Location (mm)", "Z Location (mm)"]

    # Load the data from the text files into DataFrames
    mean_df = pd.read_csv(mean_file_path, sep='\\t', encoding='latin1')
    alternating_df = pd.read_csv(alternating_file_path, sep='\\t', encoding='latin1')
    temperature_df = pd.read_csv(fatigue_temperature_file_path, sep='\\t', encoding='latin1')

    # Rename 'BFE ()' to 'T' in temperature_df
    if "BFE ()" in temperature_df.columns:
        temperature_df = temperature_df.rename(columns={"BFE ()": "T"})
    else:
        raise KeyError("'BFE()' column not found in temperature_df")

    # Add 273.15 to all values in the T column
    temperature_df["T"] = temperature_df["T"] + 273.15

    # Drop excluded columns if they exist in numerical DataFrames
    mean_df_num = mean_df.drop(columns=[col for col in exclude_cols if col in mean_df.columns])
    alternating_df_num = alternating_df.drop(columns=[col for col in exclude_cols if col in alternating_df.columns])

    # Ensure the DataFrames have the same structure for numerical columns
    if not mean_df_num.columns.equals(alternating_df_num.columns):
        raise ValueError("The files do not have the same columns for numerical operations")

    # Add and subtract the numerical values
    add_df_num = mean_df_num + alternating_df_num
    subtract_df_num = mean_df_num - alternating_df_num

    # Merge with Node Number and Location columns
    add_df = alternating_df[["Node Number", "X Location (mm)", "Y Location (mm)", "Z Location (mm)"]].copy()
    add_df = add_df.join(add_df_num)

    subtract_df = alternating_df[["Node Number", "X Location (mm)", "Y Location (mm)", "Z Location (mm)"]].copy()
    subtract_df = subtract_df.join(subtract_df_num)

    # Rename columns
    rename_dict = {
        "Node Number": "node_id",
        "X Location (mm)": "x",
        "Y Location (mm)": "y",
        "Z Location (mm)": "z",
        "SX (MPa)": "sxx",
        "SY (MPa)": "syy",
        "SZ (MPa)": "szz",
        "SXY (MPa)": "sxy",
        "SYZ (MPa)": "syz",
        "SXZ (MPa)": "sxz"
    }
    
    mean_df = mean_df.rename(columns=rename_dict)
    add_df = add_df.rename(columns=rename_dict)
    subtract_df = subtract_df.rename(columns=rename_dict)

    # Add mission_name, Load Step End Time, named_selection_name, and material_name columns
    mean_df.insert(0, "mission_name", mission_name)
    mean_df.insert(mean_df.columns.get_loc("node_id") + 1, "Load Step End Time", 1)
    mean_df.insert(mean_df.columns.get_loc("z") + 1, "named_selection_name", named_selection_name)
    mean_df.insert(mean_df.columns.get_loc("named_selection_name") + 1, "material_name", material_name)
    mean_df.insert(mean_df.columns.get_loc("material_name") + 1, "T", temperature_df["T"])

    add_df.insert(0, "mission_name", mission_name)
    add_df.insert(add_df.columns.get_loc("node_id") + 1, "Load Step End Time", 2)
    add_df.insert(add_df.columns.get_loc("z") + 1, "named_selection_name", named_selection_name)
    add_df.insert(add_df.columns.get_loc("named_selection_name") + 1, "material_name", material_name)
    add_df.insert(add_df.columns.get_loc("material_name") + 1, "T", temperature_df["T"])

    subtract_df.insert(0, "mission_name", mission_name)
    subtract_df.insert(subtract_df.columns.get_loc("node_id") + 1, "Load Step End Time", 3)
    subtract_df.insert(subtract_df.columns.get_loc("z") + 1, "named_selection_name", named_selection_name)
    subtract_df.insert(subtract_df.columns.get_loc("named_selection_name") + 1, "material_name", material_name)
    subtract_df.insert(subtract_df.columns.get_loc("material_name") + 1, "T", temperature_df["T"])

    # Write the alternating stress and mean stress CSV files in TLIFE format
    mean_df.to_csv(os.path.join('""" + solution_directory_path + """', 'TLIFE_mean_stress.csv'), index=False)
    add_df.to_csv(os.path.join('""" + solution_directory_path + """', 'TLIFE_mean_plus_alternating_stress.csv'), index=False)
    subtract_df.to_csv(os.path.join('""" + solution_directory_path + """', 'TLIFE_mean_minus_alternating_stress.csv'), index=False)

    # Create loading_history_alt_w_mean.csv with interleaved rows from mean_df, add_df, and subtract_df
    loading_history_df = pd.concat([mean_df, add_df, subtract_df]).sort_index(kind='stable').reset_index(drop=True)

    loading_history_df.to_csv(os.path.join('""" + solution_directory_path + """', 'loading_history_alt_w_mean.csv'), index=False)

    print("CSV files generated successfully.")
# endregion

# region Run the GUI
if __name__ == '__main__':
    QApplication.setAttribute(Qt.AA_EnableHighDpiScaling, True)
    app = QApplication(sys.argv)
    main_win = MainWindow()
    sys.exit(app.exec_())
# endregion
"""

# Use StreamWriter with FileStream to write the file with UTF-8 encoding
with StreamWriter(FileStream(cpython_script_path, FileMode.Create, FileAccess.Write), UTF8Encoding(True)) as writer:
    writer.Write(cpython_code)

print("Python file created successfully with UTF-8 encoding.")
# endregion

# Run the CPython script synchronously
process = Process()
# Configure the process to hide the window and not use the shell execute feature
#process.StartInfo.CreateNoWindow = True

process.StartInfo.UseShellExecute = True
# Set the command to run the Python interpreter with your script as the argument
process.StartInfo.WindowStyle = ProcessWindowStyle.Minimized
process.StartInfo.FileName = "cmd.exe"  # Use cmd.exe to allow window manipulation
process.StartInfo.Arguments = '/c python "' + cpython_script_path + '"'
# Start the process
process.Start()

# Wait for the process to complete
process.WaitForExit()

print("Python script executed successfully.")
# endregion
