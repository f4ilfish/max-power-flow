import win32com.client
import regime_config

rastr = win32com.client.Dispatch('Astra.Rastr')


def criteria1(p_fluctuations: float) -> float:
    """ Calculation of the maximum power flow (mpf) in flowgate in normal scheme"""

    # Load a clean regime and set regime parameters
    regime_config.load_clean(rastr)
    regime_config.set_regime(rastr, 200, 1, 1, 1)

    # Iterative weighting of regime
    regime_config.do_ut(rastr)

    # Maximum power flow
    mpf_1 = round((abs(rastr.Tables('sechen').Cols('psech').Z(0)) * 0.8 - p_fluctuations), 2)
    return mpf_1


def criteria2(p_fluctuations: float) -> float:
    """Calculation of the maximum power flow (mpf) by load's nodes voltage in flowgate in normal scheme"""

    # Load a clean regime and set regime parameters
    regime_config.load_clean(rastr)
    regime_config.set_regime(rastr, 200, 1, 0, 1)

    # Redefining objects RastrWin3
    nodes = rastr.Tables('node')

    # Determination of the minimum voltage in load's nodes
    i = 0

    while i < rastr.Tables('node').Size:
        # Conditions to determine node type (1 - load's nodes)
        if nodes.Cols('tip').Z(i) == 1:
            u_kr = nodes.Cols('uhom').Z(i) * 0, 7  # Critical voltage (Ucr = Unom * 0,7)
            u_min = u_kr * 1, 15  # Minimum voltage (Umin = Ucr * 1,15)
            nodes.Cols('umin').SetZ(i, u_min)
            nodes.Cols('contr_v').SetZ(i, 1)
        i += 1

    # Iterative weighting of regime
    regime_config.do_ut(rastr)

    # Maximum power flow
    mpf_2 = round((abs(rastr.Tables('sechen').Cols('psech').Z(0)) - p_fluctuations), 2)
    return mpf_2


def criteria3(p_fluctuations: float, faults_lines: dict) -> float:
    """ Calculation of the maximum power flow (mpf) in flowgate in after emergency scheme"""

    # Load a clean regime and set regime parameters
    regime_config.load_clean(rastr)
    regime_config.set_regime(rastr, 200, 1, 1, 1)

    # Create a list of mpf for each faults
    mpf_3 = []

    # Redefining objects RastrWin3
    branches = rastr.Tables('vetv')

    # Determine mpf
    i = 0
    j = 1
    # Enumerate each fault's lines
    for line in faults_lines:
        line_polus = faults_lines[line]['ip']  # Node number of the start transmission line
        line_quit = faults_lines[line]['iq']  # Node number of the start transmission line

        # Enumerate each row in Branches
        while i < branches.Size:

            # Condition for finding transmission line
            if (branches.Cols('ip').Z(i) == line_polus) and (branches.Cols('iq').Z(i) == line_quit):
                rastr.Tables('vetv').Cols('sta').SetZ(i, 1)  # Do fault

                # Do regime weighing
                regime_config.do_ut(rastr)

                # Find mpf in after emergency scheme
                mpf = abs(rastr.Tables('sechen').Cols('psech').Z(0))
                mpf_reserve = abs(rastr.Tables('sechen').Cols('psech').Z(0)) * 0.92

                # Step back
                while mpf > mpf_reserve:
                    rastr.GetToggle().MoveOnPosition(len(rastr.GetToggle().GetPositions()) - j)
                    mpf = abs(rastr.Tables('sechen').Cols('psech').Z(0))
                    j += 1
                j = 1  # Resetting step the counter

                rastr.Tables('vetv').Cols('sta').SetZ(i, 0)  # Remove fault
                rastr.rgm('p')  # Calculate regime

                # Maximum power flow
                mpf = round((abs(rastr.Tables('sechen').Cols('psech').Z(0)) - p_fluctuations), 2)
                mpf_3.append(mpf)

                # Reset to clear regime
                rastr.Load(1, 'regime.rg2', 'shablon/режим.rg2')

            i += 1  # To the next row of Branches
        i = 0  # Reset the row counter / To the next fault's line
    return min(mpf_3)


def criteria4(faults_lines: dict) -> float:
    """ Calculation of the maximum power flow (mpf) in flowgate by voltage in after emergency scheme"""

    # Load a clean regime and set parameters
    regime_config.load_clean(rastr)
    regime_config.set_regime(rastr, 200, 1, 0, 1)

    # Create a list of mpf for each faults
    mpf_4 = []

    # Redefining objects RastrWin3
    branches = rastr.Tables('vetv')
    nodes = rastr.Tables('node')

    # Determine mpf
    i = 0
    k = 0

    # Determine critical and minimum load`s nodes voltages
    while k < nodes.Size:
        if nodes.Cols('tip').Z(k) == 1:
            u_kr = nodes.Cols('uhom').Z(k) * 0, 7
            u_min = u_kr * 1, 1
            nodes.Cols('umin').SetZ(k, u_min)
        k = k + 1

    # Enumerate each fault's lines
    for line in faults_lines:
        line_polus = faults_lines[line]['ip']  # Node number of the start transmission line
        line_quit = faults_lines[line]['iq']  # Node number of the start transmission line

        # Enumerate each row in Branches
        while i < branches.Size:

            # Condition for finding transmission line
            if (branches.Cols('ip').Z(i) == line_polus) and (branches.Cols('iq').Z(i) == line_quit):
                rastr.Tables('vetv').Cols('sta').SetZ(i, 1)  # Do fault

                # Do regime weighing
                regime_config.do_ut(rastr)

                rastr.Tables('vetv').Cols('sta').SetZ(i, 1)  # Remove fault
                rastr.rgm('p')

                # Maximum power flow
                mpf = round(abs(rastr.Tables('sechen').Cols('psech').Z(0)), 2)
                mpf_4.append(mpf)

                # Reset to clear regime
                rastr.Load(1, 'regime.rg2', 'shablon/режим.rg2')

            i += 1  # To the next row of Branches
        i = 0  # Reset the row counter / To the next fault's line
    return min(mpf_4)


def criteria5(flowgate_lines: dict) -> float:
    """ Calculation of a maximum power flow (mpf) in flowgate by current in normal scheme"""

    # Load a clean regime and set regime parameters
    regime_config.load_clean(rastr)
    regime_config.set_regime(rastr, 200, 1, 1, 0)

    # Redefining objects RastrWin3
    branches = rastr.Tables('vetv')

    # Take into control flowgate`s lines
    i = 0

    for line in flowgate_lines:
        while i < rastr.Tables('vetv').Size:
            if (flowgate_lines[line]['ip'] == branches.Cols('ip').Z(i)) and (
                    flowgate_lines[line]['iq'] == branches.Cols('iq').Z(i)):
                branches.Cols('contr_i').SetZ(i, 1)
                branches.Cols('i_dop').SetZ(i, rastr.Tables('vetv').Cols('i_dop_r').Z(i))
            i = i + 1
        i = 0

    # Determine mpf
    regime_config.do_ut(rastr)

    mpf_5 = round(abs(rastr.Tables('sechen').Cols('psech').Z(0)), 2)
    return mpf_5


def criteria6(faults_lines: dict, flowgate_lines: dict):
    """ Calculation of the maximum power flow (mpf) in flowgate by current in after emergency scheme"""

    # Load a clean regime and set regime parameters
    regime_config.load_clean(rastr)
    regime_config.set_regime(rastr, 200, 1, 1, 0)

    # Create a list of mpf for each faults
    mpf_6 = []

    # Redefining objects RastrWin3
    branches = rastr.Tables('vetv')

    # Determine mpf
    i = 0
    j = 0

    # Enumerate each fault's lines
    for line in faults_lines:
        line_polus = faults_lines[line]['ip']  # Node number of the start transmission line
        line_quit = faults_lines[line]['iq']  # Node number of the start transmission line

        # Take into control flowgate`s lines
        for abc in flowgate_lines:
            while j < branches.Size:
                if (flowgate_lines[abc]['ip'] == branches.Cols('ip').Z(j)) and (
                        flowgate_lines[abc]['iq'] == branches.Cols('iq').Z(j)):
                    branches.Cols('contr_i').SetZ(j, 1)
                    branches.Cols('i_dop').SetZ(j, branches.Cols('i_dop_r_av').Z(j))
                j += 1
            # Step to the next line flowgate
            j = 0

        # Enumerate each row in Branches
        while i < branches.Size:

            # Condition for finding transmission line
            if (branches.Cols('ip').Z(i) == line_polus) and (branches.Cols('iq').Z(i) == line_quit):
                rastr.Tables('vetv').Cols('sta').SetZ(i, 1)  # Do fault

                # Do regime weighing
                regime_config.do_ut(rastr)

                rastr.Tables('vetv').Cols('sta').SetZ(i, 1)  # Remove fault
                rastr.rgm('p')

                # Maximum power flow
                mpf = round(abs(rastr.Tables('sechen').Cols('psech').Z(0)), 2)
                mpf_6.append(mpf)

                # Reset to clear regime
                rastr.Load(1, 'regime.rg2', 'shablon/режим.rg2')

            i += 1  # To the next row of Branches
        i = 0  # Reset the row counter / To the next fault's line
    return min(mpf_6)
