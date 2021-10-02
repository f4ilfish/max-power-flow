

def do_ut(rastr) -> None:
    if rastr.ut_utr('i') > 0:
        rastr.ut_utr('')


def load_clean(rastr) -> None:
    """ Load clear .rg2 file """
    rastr.Load(1, 'regime.rg2', '')
    rastr.Load(1, 'sech.sch', 'shablon/сечения.sch')
    rastr.Load(1, 'traj.ut2', 'shablon/траектория утяжеления.ut2')


def set_regime(rastr,
               max_steps: int,
               full_control: int,
               disable_v: int,
               disable_i: int) -> None:
    """ Set regime parameters for regime weighting """

    # Redefining objects RastrWin3
    power_flow_control = rastr.Tables('ut_common')

    # Set parameters
    power_flow_control.Cols('iter').SetZ(0, max_steps)
    power_flow_control.Cols('enable_contr').SetZ(0, full_control)
    power_flow_control.Cols('dis_v_contr').SetZ(0, disable_v)
    power_flow_control.Cols('dis_i_contr').SetZ(0, disable_i)
