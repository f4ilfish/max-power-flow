import sys

import data_extract
import file_creating
import calc_criteria

print("Hello. Let`s calculate maximum power flow (MPF)")

# Select path to flowgate .json
print("Select path to flowgate .json:")
try:
    flowgate_lines = data_extract.json_to_dic(input())
except Exception:
    print("Unknown path, type of file or folder does not contain a file")
    exit()

# Select path to faults .json
print("Select path to faults .json:")
try:
    faults_lines = data_extract.json_to_dic(input())
except Exception:
    print("Unknown path, type of file or folder does not contain a file")
    exit()

# Select path to trajectory .csv
print("Select path to trajectory .csv:")
try:
    trajectory_nodes = data_extract.csv_to_list(input())
except Exception:
    print("Unknown path, type of file or folder does not contain a file")
    exit()

# Input power fluctuation
print("Input positive power fluctuations:")
try:
    p_fluctuations = int(input())
except Exception:
    print("Input positive number")
    exit()

# Creating RastrWin3 files
file_creating.do_sch(flowgate_lines, 'new')
file_creating.do_ut2(trajectory_nodes)

# Calcutale criterias
print(f"MPF in normal scheme (0.8*Pmax): "
      f"{calc_criteria.criteria1(p_fluctuations)}")
print(f"MPF by voltage in normal scheme (1,15*Ucr): "
      f"{calc_criteria.criteria2(30)}")
print(f"MPF in after emergency scheme (0.92*Pmax): "
      f"{calc_criteria.criteria3(30, faults_lines)}")
print(f"MPF by voltage in after emergency scheme (1.1*Ucr): "
      f"{calc_criteria.criteria4(faults_lines)}")
print(f"MPF by current in normal scheme (I - allowable): "
      f"{calc_criteria.criteria5(flowgate_lines)}")
print(f"MPF by current in after emergency scheme (I - critical): "
      f"{calc_criteria.criteria6(faults_lines, flowgate_lines)}")
print()

