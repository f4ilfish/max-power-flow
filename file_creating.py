import win32com.client

rastr = win32com.client.Dispatch('Astra.Rastr')


def do_sch(flowgate_lines: dict, flowgate_name: str):
    """ Create .sch file
    based on RastrWin3 template from flowgate dictionary"""

    # Create new empty .sch file based on template
    rastr.Save('sech.sch',
               'C:/Users/mishk/Documents/RastrWin3/SHABLON/сечения.sch')
    # Open the created file
    rastr.Load(1, 'sech.sch',
               'C:/Users/mishk/Documents/RastrWin3/SHABLON/сечения.sch')

    # Redefining objects RastrWin3
    FlowGate = rastr.Tables('sechen')
    GroupLine = rastr.Tables('grline')

    # Just in case clear rows in .sch
    FlowGate.DelRows()
    GroupLine.DelRows()

    # Create flowgate
    FlowGate.AddRow()
    FlowGate.Cols('ns').SetZ(0, 1)
    # Give a name for the flowgate
    FlowGate.Cols('name').SetZ(0, flowgate_name)
    FlowGate.Cols('sta').SetZ(0, 1)

    # Fill a list of transmission lines forms the flowgate
    i = 0

    for line in flowgate_lines:
        GroupLine.AddRow()
        GroupLine.Cols('ns').SetZ(i, 1)

        # Start of the transmission line
        start_node = flowgate_lines[line]['ip']
        # End of the transmission line
        end_node = flowgate_lines[line]['iq']

        GroupLine.Cols('ip').SetZ(i, start_node)
        GroupLine.Cols('iq').SetZ(i, end_node)

        i += 1

    # Resave .sch file
    rastr.Save('sech.sch',
               'C:/Users/mishk/Documents/RastrWin3/SHABLON/сечения.sch')


def do_ut2(trajectory_nodes: list):
    """ Create .ut2 file
    based on RastrWin3 template from list of trajectorie`s nodes"""

    # Create new empty .ut2 file based on template
    rastr.Save('traj.ut2',
               'C:/Users/mishk/Documents/RastrWin3/SHABLON/траектория утяжеления.ut2')
    # Open the created file
    rastr.Load(1,
               'traj.ut2',
               'C:/Users/mishk/Documents/RastrWin3/SHABLON/траектория утяжеления.ut2')

    # Redefining objects RastrWin3
    Trajectory = rastr.Tables('ut_node')

    # Just in case clear rows in .ut2
    Trajectory.DelRows()

    # Fill a .ut2 list of nodes forms trajectory
    i = 0
    # To avoid duplicates of nodes create empty dictionary
    # This is intended for nodes
    # that can be generator and load at the same time
    node_data = {}  # create empty dictionary

    for node in trajectory_nodes:
        node_type = node['variable']  # Pg - generator / Pn - load
        node_number = node['node']
        power_change = float(node['value'])
        power_tg = node['tg']  # Load's power factor

        # Check whether the dictionary contains a node
        if node_number not in node_data:
            # Create a pair node number - index
            node_data[node_number] = i
            i += 1
            # Fill row in .ut2
            Trajectory.AddRow()
            Trajectory.Cols('ny').SetZ(node_data[node_number],
                                       node_number)
            Trajectory.Cols(node_type).SetZ(node_data[node_number],
                                            power_change)
        else:
            # Find existing pair and add to existing row in .ut2
            Trajectory.Cols(node_type).SetZ(node_data[node_number],
                                            power_change)

        # Try add load's power factor
        if Trajectory.Cols('tg').Z(node_data[node_number]) == 0:
            Trajectory.Cols('tg').SetZ(node_data[node_number], power_tg)

    # Resave .ut2 file
    rastr.Save('traj.ut2',
               'C:/Users/mishk/Documents/RastrWin3/SHABLON/траектория утяжеления.ut2')

