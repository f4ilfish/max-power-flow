# maxPowerFlow ⚡
## Intro
It is Python script for calculating Maximum Power Flow (MPF) in flowgate of power system. MPF determine according with [standart](https://www.so-ups.ru/fileadmin/files/laws/standards/st_max_power_rules_004-2020.pdf) of Russian System Operator of the United Power System (["SO UPS", JSC](https://www.so-ups.ru))

## Technical requirments 
* [Python (32-bit)](https://www.python.org/downloads/windows/)
* [pywin32](https://pypi.org/project/pywin32/)
* [RastrWin3 (x86) v 2.0.0.5709 or better](https://www.rastrwin.ru/rastr/)

## Calculating requirments
* Power system model .rg2 file
* Flowgate .json file
* Trajectory .csv file
* Faults .json file

## Manual
1. Make yours .rg2 file by RastrWin3 and put it in /samples folder
2. Run "main" file
3. Input path to .json and .csv files forming flowgate and trajectory
4. Input value of active power fluctuation

_Example of script running:_
```commandline
>> Select path to flowgate .json: C:\...\samples\flowgate.json
>> Select path to faults .json: C:\...\samples\faults.json
>> Select path to trajectory .csv: C:\...\samples\vector.csv
>> Input positive power fluctuations: 30
```
_Example of script output:_
```commandline
>> MPF in normal regime (0.8*Pmax): 2216.57
>> MPF by the acceptable voltage level in the pre-emergency regime (1,15*Ucr): 2757.89
>> MPF in the post-emergency regime after fault (0.92*Pmax): 2132.61
>> MPF by the acceptable voltage level in the post-emergency regime after fault (1.1*Ucr): 2339.69
>> MPF by acceptable current in normal regime (Iacc): 1787.14
>> MPF by acceptable current in the post-emergency regime after fault (Iem_acc): 1479.26

```
___
## Explanation
_About .rg2 files_
```
To find out how to make a .rg file use RastrWin3 -> Помощь -> Справка -> User Manual EN
```

_About flowgate .json_

Use flowgate .json file such a this structure:
```json
    {
	    "line_1": {
		    "ip": 17, 
		    "iq": 16, 
		    "np": 0
	    }, 
	    "line_2": {
		    "ip": 6, 
		    "iq": 11, 
		    "np": 0
	    }, 
	    "line_3": {
		    "ip": 4, 
		    "iq": 14, 
		    "np": 0
	    }
    }
```
| Parameter | Meaning | Values
:-------- |:-----:| -------:
ip  | Node number according to the start of the line in .rg2 | int
iq  | Node number according to the end of the line in .rg2 | int
np  | Parallel line number in .rg2 | usually 1 or 2

_About faults .json_

Use faults .json file such a this structure:
```json
    {
	    "outage_of_6_11": {
		    "ip": 6, 
		    "iq": 11, 
		    "np": 0,
		    "sta": 1
	    }, 
	    "outage_of_4_14": {
		    "ip": 4, 
		    "iq": 14,
		    "np": 0,
		    "sta": 1
	    }
    }
```
| Parameter | Meaning | Values
:-------- |:-----:| -------:
ip  | Node number according to the start of the line in .rg2 | int
iq  | Node number according to the end of the line in .rg2 | int
np  | Parallel line number in .rg2 | usually 1 or 2
sta | Line status     | 0 - on, 1 - off

_About trajectory .csv_

Use trajectory .csv file such a this structure:
```editorconfig
    variable,node,value,tg
    pn,23,-3.0,1
    pn,24,-3.0,1
    ...
    pg,6,-4.0,0
    pg,7,-4.0,0
    pg,8,-2.0,0
    pg,9,-1.0,0
```
| Parameter | Meaning | Values
:-------- |:-----:| -------:
variable  | Type of node in .rg2 | pn - load, pg - generator
node  | Node number in .rg2 | int
value  | Changing the power in the node | float
tg | Power factor in node | float
