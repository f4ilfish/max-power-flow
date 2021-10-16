import data_extract
import file_creating
import calc_criteria
import os

print("Hello. Let`s calculate maximum power flow (MPF)")

# State variables
flowgate_lines = None
trajectory_nodes = None
faults_lines = None
p_fluctuations = 0

# Select path to flowgate .json
try:
    flowgate_lines = data_extract.json_to_dic(input("Select path to "
                                                    "flowgate .json: "))
except:
    print("Unknown path, type of file or folder does not contain a file")
    exit()

# Select path to faults .json
try:
    faults_lines = data_extract.json_to_dic(input("Select path to "
                                                  "faults .json: "))
except:
    print("Unknown path, type of file or folder does not contain a file")
    exit()

# Select path to trajectory .csv
try:
    trajectory_nodes = data_extract.csv_to_list(input("Select path to "
                                                      "trajectory .csv: "))
except:
    print("Unknown path, type of file or folder does not contain a file")
    exit()

# Input power fluctuation
try:
    p_fluctuations = int(input("Input positive power fluctuations: "))
except:
    print("Input positive number")
    exit()

# Creating RastrWin3 files
file_creating.create_file_sch(flowgate_lines, 'new')
file_creating.create_file_ut2(trajectory_nodes)

# Calculation of criteria
print(f"MPF in normal regime (0.8*Pmax): "
      f"{calc_criteria.criteria1(p_fluctuations)}")
print(f"MPF by the acceptable voltage level "
      f"in the pre-emergency regime (1,15*Ucr): "
      f"{calc_criteria.criteria2(p_fluctuations)}")
print(f"MPF in the post-emergency regime after fault (0.92*Pmax): "
      f"{calc_criteria.criteria3(p_fluctuations, faults_lines)}")
print(f"MPF by the acceptable voltage level "
      f"in the post-emergency regime after fault (1.1*Ucr): "
      f"{calc_criteria.criteria4(p_fluctuations, faults_lines)}")
print(f"MPF by acceptable current in normal regime (Iacc): "
      f"{calc_criteria.criteria5(p_fluctuations)}")
print(f"MPF by acceptable current "
      f"in the post-emergency regime after fault (Iem_acc): "
      f"{calc_criteria.criteria6(p_fluctuations, faults_lines)}")

# Delete temporary files
for item in os.listdir('.'):
    if item.endswith(".sch") or item.endswith(".ut2"):
        os.remove(item)

