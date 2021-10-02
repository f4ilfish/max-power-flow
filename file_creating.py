import win32com.client

rastr = win32com.client.Dispatch('Astra.Rastr')


def do_sch(flowgate_lines: dict, flowgate_name: str) -> None:
    """ Create .sch file based on RastrWin3 template from flowgate dictionary"""

    # Create new empty .sch file based on template
    rastr.Save('sech.sch', 'shablon/сечения.sch')
    # Open the created file
    rastr.Load(1, 'sech.sch', 'shablon/сечения.sch')

    # Redefining objects RastrWin3
    flow_gate = rastr.Tables('sechen')
    group_line = rastr.Tables('grline')

    # Just in case clear rows in .sch
    flow_gate.DelRows()
    group_line.DelRows()

    # Create flowgate
    flow_gate.AddRow()
    flow_gate.Cols('ns').SetZ(0, 1)
    # Give a name for the flowgate
    flow_gate.Cols('name').SetZ(0, flowgate_name)
    flow_gate.Cols('sta').SetZ(0, 1)

    # Fill a list of transmission lines forms the flowgate
    i = 0

    for line in flowgate_lines:
        group_line.AddRow()
        group_line.Cols('ns').SetZ(i, 1)

        # Start of the transmission line
        start_node = flowgate_lines[line]['ip']
        # End of the transmission line
        end_node = flowgate_lines[line]['iq']

        group_line.Cols('ip').SetZ(i, start_node)
        group_line.Cols('iq').SetZ(i, end_node)

        i += 1

    # Resave .sch file
    rastr.Save('sech.sch', 'shablon/сечения.sch')


def do_ut2(trajectory_nodes: list) -> None:
    """ Create .ut2 file based on RastrWin3 template from list of trajectories`s nodes"""

    # Create new empty .ut2 file based on template
    rastr.Save('traj.ut2', 'shablon/траектория утяжеления.ut2')
    # Open the created file
    rastr.Load(1, 'traj.ut2', 'shablon/траектория утяжеления.ut2')

    # Redefining objects RastrWin3
    trajectory = rastr.Tables('ut_node')

    # Just in case clear rows in .ut2
    trajectory.DelRows()

    # Fill a .ut2 list of nodes forms trajectory
    i = 0
    # To avoid duplicates of nodes create empty dictionary this is intended for nodes
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
            trajectory.AddRow()
            trajectory.Cols('ny').SetZ(node_data[node_number], node_number)
            trajectory.Cols(node_type).SetZ(node_data[node_number], power_change)
        else:
            # Find existing pair and add to existing row in .ut2
            trajectory.Cols(node_type).SetZ(node_data[node_number], power_change)

        # Try add load's power factor
        if trajectory.Cols('tg').Z(node_data[node_number]) == 0:
            trajectory.Cols('tg').SetZ(node_data[node_number], power_tg)

    # Resave .ut2 file
    rastr.Save('traj.ut2', 'shablon/траектория утяжеления.ut2')
