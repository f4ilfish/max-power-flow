import data_extract
import file_creating
import calc_criteria


# Считываем файлы
flowgate_lines = data_extract.json_to_dic('flowgate.json')
faults_lines = data_extract.json_to_dic('faults.json')
trajectory_nodes = data_extract.csv_to_list('vector.csv')

# Создаем файлы
file_creating.do_sch(flowgate_lines, 'new')
file_creating.do_ut2(trajectory_nodes)

# Считаем режим
print(calc_criteria.criteria1(30))
print(calc_criteria.criteria2(30))
print(calc_criteria.criteria3(30, faults_lines))

