Public Module FunctionalTestSequence
    ''' <summary>
    ''' Defines all DVT functional tests in the order for sorting.
    ''' </summary>
    Public ReadOnly Property TestNames As List(Of String)
        Get
            Return New List(Of String) From {
                "Input Voltage Turn On",
                "Input Voltage Turn Off",
                "Input Current",
                "No Load Input Power",
                "Inhibited Input Power",
                "Inrush Current",
                "Power Factor",
                "Efficiency",
                "Output Voltage Setting Accuracy",
                "Line Regulation",
                "Load Regulation",
                "Current Share Voltage",
                "I-Share Voltage",
                "Ripple and Noise",
                "Start-Up Time",
                "Start Up Time",
                "Startup Time",
                "Start Up Rise Time",
                "Start-Up Rise Time",
                "Startup Rise Time",
                "Rise Time",
                "Hold-Up Time",
                "Hold Up Time",
                "Holdup Time",
                "Turn Off Fall Time",
                "Fall Time",
                "Turn On Overshoot",
                "Cross Regulation",
                "Over Current Protection",
                "Over Voltage Protection",
                "Short Circuit Protection",
                "Transient Response",
                "Transient Response (CC)",
                "Transient Response (CR)",
                "Initial Set Accuracy",
                "Output Voltage Stability",
                "Current Limit Set Acc.",
                "Current Read Acc.",
                "Voltage Set-Read Acc.",
                "Voltage Set Acc.",
                "Voltage Read Acc.",
                "IPROG Accuracy",
                "IPROG Emulation",
                "VPROG Accuracy",
                "VPROG Emulation",
                "Slew Rate Rise",
                "Slew Rate Fall",
                "Early Warning Time",
                "AC OK",
                "DC OK",
                "Output Voltage Setting Emulation",
                "IShare",
                "Dip Test"
            }
        End Get
    End Property
End Module