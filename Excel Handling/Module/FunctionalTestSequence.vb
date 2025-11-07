Public Module FunctionalTestSequence
    ''' <summary>
    ''' Defines all DVT functional tests in the order for sorting.
    ''' </summary>
    Public ReadOnly Property TestNames As List(Of String)
        Get
            Return New List(Of String) From {
                "Title", 'Context of << DVT Full Report >> 
                "Summary", 'Context of << DVT Full Report >> 
                "Input Voltage Turn On",
                "Input Voltage Turn Off",
                "Input Current",
                "No Load Input Power",
                "Inhibited Input Power",
                "Inrush Current",
                "Power Factor",
                "Ripple and Noise",
                "Start Up Time",
                "Start Up Rise Time",
                "Hold-Up Time",
                "Turn Off Fall Time",
                "Turn On Overshoot",
                "Efficiency",
                "Output Voltage Setting Accuracy",
                "Line Regulation",
                "Load Regulation",
                "Cross Regulation",
                "Over Current Protection",
                "Over Voltage Protection",
                "Short Circuit Protection",
                "Transient Response",
                "Transient Response (CC)",
                "Transient Response (CR)",
                "Initial Set Accuracy",
                "Output Voltage Stability",
                "Current Limit Setting",
                "Current Output Reading",
                "Voltage Output Setting",
                "Voltage Output Reading",
                "Voltage Output Setting and Reading",
                "IPROG Accuracy",
                "VPROG Accuracy",
                "VPROG Emulation",
                "IPROG Emulation",
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