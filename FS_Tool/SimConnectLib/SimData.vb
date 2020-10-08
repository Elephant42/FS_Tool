Option Explicit On
Option Strict On

Public Class SimData
    '
    'Any extra SimVars to be read must be added to the processSimVars() sub.
    '
    'Any extra Sim Events to be sent must be added to the EventEnum enum.
    '

    'These constants are for convenience only and are not required by the base library.
    Public Const SIM_RATE As String = "SIMULATION RATE"
    Public Const LONGITUDE As String = "PLANE LONGITUDE"
    Public Const LATITUDE As String = "PLANE LATITUDE"
    Public Const HEADING As String = "PLANE HEADING DEGREES MAGNETIC"
    Public Const ALTITUDE As String = "PLANE ALTITUDE"
    Public Const AIRSPEED_INDICATED As String = "AIRSPEED INDICATED"
    Public Const AUTOPILOT_MASTER As String = "AUTOPILOT MASTER"
    Public Const PARKING_BRAKE As String = "BRAKE PARKING INDICATOR"
    'Public Const  As String=""

    Public Enum EventEnum
        AP_MASTER
        AUTOPILOT_ON
        AUTOPILOT_OFF
        THROTTLE_SET
        AP_SPD_VAR_SET
        SIM_RATE_INCR
        SIM_RATE_DECR
        PARKING_BRAKES
        ALL_LIGHTS_TOGGLE
        VOR1_OBI_DEC
        VOR1_OBI_INC
        VOR2_OBI_DEC
        VOR2_OBI_INC
        STROBES_TOGGLE
        PANEL_LIGHTS_TOGGLE
        NAV1_RADIO_FRACT_DEC
        NAV1_RADIO_FRACT_DEC_CARRY
        NAV1_RADIO_FRACT_INC
        NAV1_RADIO_FRACT_INC_CARRY
        NAV1_RADIO_SET
        NAV1_RADIO_SWAP
        NAV1_RADIO_WHOLE_DEC
        NAV1_RADIO_WHOLE_INC
        NAV1_STBY_SET
        NAV2_RADIO_FRACT_DEC
        NAV2_RADIO_FRACT_DEC_CARRY
        NAV2_RADIO_FRACT_INC
        NAV2_RADIO_FRACT_INC_CARRY
        NAV2_RADIO_SET
        NAV2_RADIO_SWAP
        KOHLSMAN_DEC
        KOHLSMAN_INC
        KOHLSMAN_SET
        LANDING_LIGHTS_TOGGLE
        HEADING_BUG_DEC
        HEADING_BUG_INC
        GEAR_TOGGLE
        FLAPS_1
        FLAPS_2
        FLAPS_3
        FLAPS_DECR
        FLAPS_DOWN
        FLAPS_INCR
        FLAPS_UP
        TOGGLE_AVIONICS_MASTER
        TOGGLE_BEACON_LIGHTS
        TOGGLE_DME
        TOGGLE_FLIGHT_DIRECTOR
        TOGGLE_MASTER_BATTERY
        TOGGLE_TAXI_LIGHTS
        PITOT_HEAT_TOGGLE

        NONE
    End Enum

    Public ReadOnly Property SimVars As Dictionary(Of String, SimVar)
        Get
            Return _SimVars
        End Get
    End Property

    Public ReadOnly Property SimVarsIndex As Dictionary(Of Integer, SimVar)
        Get
            Return _SimVarsIndex
        End Get
    End Property

    Public ReadOnly Property SimVarNames As String()
        Get
            Return _SimVars.Keys.ToArray
        End Get
    End Property

    Public Structure SimVar
        Public Index As Integer
        Public Name As String
        Public Units As String
    End Structure

    Public Sub New()
        processSimVars()
    End Sub

    Public Function Contains(ByVal varIndex As Integer) As Boolean
        Return _SimVarsIndex.ContainsKey(varIndex)
    End Function
    Public Function Contains(ByVal varName As String) As Boolean
        Return _SimVars.ContainsKey(varName)
    End Function

    Private Sub processSimVars()

        addSimVar("PLANE LONGITUDE", "degree")
        addSimVar("PLANE LATITUDE", "degree")
        addSimVar("PLANE HEADING DEGREES MAGNETIC", "degree")
        addSimVar("PLANE ALTITUDE", "feet")
        addSimVar("AIRSPEED INDICATED", "knots")
        addSimVar("SIMULATION RATE", "times")
        addSimVar("AUTOPILOT MASTER", "Bool")
        addSimVar("AUTOPILOT DISENGAGED", "Bool")
        addSimVar("BRAKE PARKING INDICATOR", "Bool")

    End Sub

    Private Sub addSimVar(ByVal name As String, ByVal units As String)

        Static index As Integer = 1

        Dim t As SimVar
        t.Index = index
        t.Name = name
        t.Units = units
        _SimVars.Add(name, t)
        _SimVarsIndex.Add(index, t)

        index += 1

    End Sub

    Private _SimVars As New Dictionary(Of String, SimVar)
    Private _SimVarsIndex As New Dictionary(Of Integer, SimVar)

End Class
