

def set_regime(rastr,
               max_steps: int,
               full_control: int,
               disable_v: int,
               disable_i: int):
    """ Set regime parameters for regime weighting """

    # Redefining objects RastrWin3
    PowerFlowControl = rastr.Tables('ut_common')

    # Set parameters
    PowerFlowControl.Cols('iter').SetZ(0, max_steps)
    PowerFlowControl.Cols('enable_contr').SetZ(0, full_control)
    PowerFlowControl.Cols('dis_v_contr').SetZ(0, disable_v)
    PowerFlowControl.Cols('dis_i_contr').SetZ(0, disable_i)