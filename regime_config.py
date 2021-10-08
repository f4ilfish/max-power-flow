from win32com.client import Dispatch


def do_regime_weight(rastr: Dispatch) -> None:
    """ Iterative regime weighting """
    if rastr.ut_utr('i') > 0:
        rastr.ut_utr('')


def load_clean_regime(rastr: Dispatch) -> None:
    """ Load clear .rg2 file """
    rastr.Load(1, 'regime.rg2', 'shablon/режим.rg2')


def load_sech(rastr: Dispatch) -> None:
    """ Load clear .sch file """
    rastr.Load(1, 'sech.sch', 'shablon/сечения.sch')


def load_traj(rastr: Dispatch) -> None:
    """ Load .ut2 file """
    rastr.Load(1, 'traj.ut2', 'shablon/траектория утяжеления.ut2')


def set_regime(rastr: Dispatch,
               max_steps: int,
               full_control: int,
               disable_v: int,
               disable_i: int) -> None:
    """ Set regime parameters for regime weighting """

    # Set Maximum number of regime weighting steps
    rastr.Tables('ut_common').Cols('iter').SetZ(0, max_steps)
    # Set control all weighting parameters
    rastr.Tables('ut_common').Cols('enable_contr').SetZ(0, full_control)
    # Disable node voltage control
    rastr.Tables('ut_common').Cols('dis_v_contr').SetZ(0, disable_v)
    # Disable branch current control
    rastr.Tables('ut_common').Cols('dis_i_contr').SetZ(0, disable_i)
