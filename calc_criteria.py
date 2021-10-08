import win32com.client
import regime_config

rastr = win32com.client.Dispatch('Astra.Rastr')


def criteria1(p_fluctuations: float) -> float:
    """ Calculation of the maximum power flow (MPF) in normal regime """

    # Load a regime files and set weighting parameters
    regime_config.load_clean_regime(rastr)
    regime_config.load_sech(rastr)
    regime_config.load_traj(rastr)
    regime_config.set_regime(rastr, 200, 1, 1, 1)

    # Iterative weighting of regime
    regime_config.do_regime_weight(rastr)

    # Maximum power flow by criteria 1
    mpf_1 = abs(
        rastr.Tables('sechen').Cols('psech').Z(0)) * 0.8 - p_fluctuations
    mpf_1 = round(mpf_1, 2)
    return mpf_1


def criteria2(p_fluctuations: float) -> float:
    """ Calculation of the maximum power flow by the acceptable voltage level
    in the pre-emergency regime """

    # Load a regime files and set weighting parameters
    regime_config.load_clean_regime(rastr)
    regime_config.load_sech(rastr)
    regime_config.load_traj(rastr)
    regime_config.set_regime(rastr, 200, 1, 0, 1)

    # Redefine the COM path to the RastrWin3 node table
    nodes = rastr.Tables('node')

    # Determining the acceptable voltage level of nodes with load
    i = 0
    while i < nodes.Size:
        # Load node search (1 - type of node with load)
        if nodes.Cols('tip').Z(i) == 1:
            u_kr = nodes.Cols('uhom').Z(i) * 0.7  # Critical voltage level
            u_min = u_kr * 1.15  # Acceptable voltage level
            nodes.Cols('umin').SetZ(i, u_min)
            nodes.Cols('contr_v').SetZ(i, 1)
        i += 1

    # Iterative weighting of regime
    regime_config.do_regime_weight(rastr)

    # MPF by criteria 2
    mpf_2 = abs(rastr.Tables('sechen').Cols('psech').Z(0)) - p_fluctuations
    mpf_2 = round(mpf_2, 2)
    return mpf_2


def criteria3(p_fluctuations: float, faults_lines: dict) -> float:
    """ Calculation of the maximum power flow (MPF)
    in the post-emergency regime after fault """

    # Load a regime files and set weighting parameters
    regime_config.load_clean_regime(rastr)
    regime_config.load_sech(rastr)
    regime_config.load_traj(rastr)
    regime_config.set_regime(rastr, 200, 1, 1, 1)

    # Redefine the COM path to the RastrWin3 branch table
    branches = rastr.Tables('vetv')
    # Redefine the COM path to the RastrWin3 flowgate table
    flowgate = rastr.Tables('sechen')

    # List of MPF for each fault
    mpf_3 = []

    # Iterating over each fault
    for line in faults_lines:
        # Node number of the start branch
        node_start_branch = faults_lines[line]['ip']
        # Node number of the start branch
        node_end_branch = faults_lines[line]['iq']
        # Number of parallel branch
        parallel_number = faults_lines[line]['np']
        # Status of branch (0 - on / 1 - off)
        branch_status = faults_lines[line]['sta']

        # Iterating over each branches in RastrWin3
        i = 0
        while i < branches.Size:

            # Search branch with fault
            if (branches.Cols('ip').Z(i) == node_start_branch) and \
                    (branches.Cols('iq').Z(i) == node_end_branch) and \
                    (branches.Cols('np').Z(i) == parallel_number):

                # Remember previous branch status
                pr_branch_status = branches.Cols('sta').Z(i)
                # Do fault
                branches.Cols('sta').SetZ(i, branch_status)

                # Do regime weighing
                regime_config.do_regime_weight(rastr)

                # MPF in the post-emergency regime after fault
                mpf = abs(flowgate.Cols('psech').Z(0))
                # Acceptable level of MPF in such scheme
                mpf_acceptable = abs(flowgate.Cols('psech').Z(0)) * 0.92

                # Iterative return to Acceptable level of MPF
                j = 1
                while mpf > mpf_acceptable:
                    rastr.GetToggle().MoveOnPosition(
                        len(rastr.GetToggle().GetPositions()) - j)
                    mpf = abs(flowgate.Cols('psech').Z(0))
                    j += 1

                # Remove fault
                branches.Cols('sta').SetZ(i, pr_branch_status)
                # Re-calculation of regime
                rastr.rgm('p')

                # MPF by criteria 3
                mpf = abs(
                    rastr.Tables('sechen').Cols('psech').Z(0)) - p_fluctuations
                mpf = round(mpf, 2)
                mpf_3.append(mpf)

                # Reset to clean regime
                regime_config.load_clean_regime(rastr)
            i += 1
    return min(mpf_3)


def criteria4(p_fluctuations: float, faults_lines: dict) -> float:
    """ Calculation of the maximum power flow (MPF)
    by the acceptable voltage level
    in the post-emergency regime after fault """

    # Load a regime files and set weighting parameters
    regime_config.load_clean_regime(rastr)
    regime_config.load_sech(rastr)
    regime_config.load_traj(rastr)
    regime_config.set_regime(rastr, 200, 1, 0, 1)

    # Redefine the COM path to the RastrWin3 node table
    nodes = rastr.Tables('node')
    # Redefine the COM path to the RastrWin3 branch table
    branches = rastr.Tables('vetv')
    # Redefine the COM path to the RastrWin3 flowgate table
    flowgate = rastr.Tables('sechen')

    # Determining the acceptable voltage level of nodes with load
    j = 0
    while j < nodes.Size:
        # Load node search (1 - type of node with load)
        if nodes.Cols('tip').Z(j) == 1:
            # Critical voltage level
            u_kr = nodes.Cols('uhom').Z(j) * 0.7
            # Acceptable voltage level
            u_min = u_kr * 1.1
            nodes.Cols('umin').SetZ(j, u_min)
        j += 1

    # List of MPF for each fault
    mpf_4 = []

    # Iterating over each fault
    for line in faults_lines:
        # Node number of the start transmission line
        node_start_branch = faults_lines[line]['ip']
        # Node number of the start transmission line
        node_end_branch = faults_lines[line]['iq']
        # Number of branch
        parallel_number = faults_lines[line]['np']
        # Status of branch (0 - on / 1 - off)
        branch_status = faults_lines[line]['sta']

        # Iterating over branch in RastrWin3
        i = 0
        while i < branches.Size:

            # Search branch with fault
            if (branches.Cols('ip').Z(i) == node_start_branch) and \
                    (branches.Cols('iq').Z(i) == node_end_branch) and \
                    (branches.Cols('np').Z(i) == parallel_number):

                # Remember previous branch status
                pr_branch_status = branches.Cols('sta').Z(i)
                # Do fault
                branches.Cols('sta').SetZ(i, branch_status)

                # Do regime weighing
                regime_config.do_regime_weight(rastr)
                # Remove fault
                branches.Cols('sta').SetZ(i, pr_branch_status)
                # Re-calculation of regime
                rastr.rgm('p')

                # MPF be criteria 4
                mpf = abs(
                    flowgate.Cols('psech').Z(0)) - p_fluctuations
                mpf = round(mpf, 2)
                mpf_4.append(mpf)

                # Reset to clean regime
                regime_config.load_clean_regime(rastr)
            i += 1
    return min(mpf_4)


def criteria5(p_fluctuations: float, flowgate_lines: dict) -> float:
    """ Calculation of a maximum power flow (MPF) by acceptable current
    in normal regime """

    # Load a regime files and set weighting parameters
    regime_config.load_clean_regime(rastr)
    regime_config.load_sech(rastr)
    regime_config.load_traj(rastr)
    regime_config.set_regime(rastr, 200, 1, 1, 0)

    # Redefine the COM path to the RastrWin3 branch table
    branches = rastr.Tables('vetv')
    # Redefine the COM path to the RastrWin3 flowgate table
    flowgate = rastr.Tables('sechen')
    # Redefine the COM path to collection of regimes RastrWin3

    # Determine of acceptable current level of flowgate`s branch
    for control_lines in flowgate_lines:
        line_param = flowgate_lines[control_lines]

        # Iterating over each branches in RastrWin3
        i = 0
        while i < branches.Size:

            # Search branch that need to control
            if (line_param['ip'] == branches.Cols('ip').Z(i)) and \
                    (line_param['iq'] == branches.Cols('iq').Z(i)) and \
                    (line_param['np'] == branches.Cols('np').Z(i)):
                # Take into control certain branch
                branches.Cols('contr_i').SetZ(i, 1)
                branches.Cols('i_dop').SetZ(i, branches.Cols('i_dop_r').Z(i))
            i += 1

    # Iterative weighting of regime
    regime_config.do_regime_weight(rastr)

    # MPF by criteria 5
    mpf_5 = abs(flowgate.Cols('psech').Z(0)) - p_fluctuations
    mpf_5 = round(mpf_5, 2)
    return mpf_5


def criteria6(p_fluctuations: float,
              faults_lines: dict,
              flowgate_lines: dict):
    """ Calculation of a maximum power flow (MPF)
    by acceptable current  in the post-emergency regime after fault """

    # Load a regime files and set weighting parameters
    regime_config.load_clean_regime(rastr)
    regime_config.load_sech(rastr)
    regime_config.load_traj(rastr)
    regime_config.set_regime(rastr, 200, 1, 1, 0)

    # Redefine the COM path to the RastrWin3 branch table
    branches = rastr.Tables('vetv')
    # Redefine the COM path to the RastrWin3 flowgate table
    flowgate = rastr.Tables('sechen')

    # List of MPF for each fault
    mpf_6 = []

    # Iterating over each fault
    for line in faults_lines:
        # Node number of the start branch
        node_start_branch = faults_lines[line]['ip']
        # Node number of the end branch
        node_end_branch = faults_lines[line]['iq']
        # Number of parallel branch
        parallel_number = faults_lines[line]['np']
        # Status of branch (0 - on / 1 - off)
        branch_status = faults_lines[line]['sta']

        # Determine of acceptable current level of flowgate`s branch
        for control_lines in flowgate_lines:
            line_param = flowgate_lines[control_lines]

            # Iterating over each branch in RastrWin3
            j = 0
            while j < branches.Size:
                # Search branch that need to control
                if (line_param['ip'] == branches.Cols('ip').Z(j)) and \
                        (line_param['iq'] == branches.Cols('iq').Z(j)) and \
                        (line_param['np'] == branches.Cols('np').Z(j)):

                    # Take into control certain branch
                    branches.Cols('contr_i').SetZ(j, 1)
                    branches.Cols('i_dop').SetZ(
                        j, branches.Cols('i_dop_r_av').Z(j))
                j += 1

        # Iterating over each branch in RastrWin3
        i = 0
        while i < branches.Size:

            # # Search branch with fault
            if (branches.Cols('ip').Z(i) == node_start_branch) and \
                    (branches.Cols('iq').Z(i) == node_end_branch) and \
                    (branches.Cols('np').Z(i) == parallel_number):

                # Remember previous branch status
                pr_branch_status = branches.Cols('sta').Z(i)
                # Do fault
                branches.Cols('sta').SetZ(i, branch_status)

                # Iterative weighting of regime
                regime_config.do_regime_weight(rastr)

                # Remove fault
                branches.Cols('sta').SetZ(i, pr_branch_status)
                # Re-calculation of regime
                rastr.rgm('p')

                # MPF by criteria 6
                mpf = abs(flowgate.Cols('psech').Z(0)) - p_fluctuations
                mpf = round(mpf, 2)
                mpf_6.append(mpf)

                # Reset to clean regime
                regime_config.load_clean_regime(rastr)
            i += 1  # To the next row of Branches
    return min(mpf_6)
