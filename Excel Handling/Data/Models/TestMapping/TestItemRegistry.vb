Option Strict On
Option Explicit On

Imports System.Collections.ObjectModel

Public NotInheritable Class TestItemRegistry

    Private Sub New()
        ' Prevent instantiation
    End Sub

    ' =========================================
    ' Master Dictionary (Key = ItemName)
    ' =========================================
    Private Shared ReadOnly _items As IReadOnlyDictionary(Of String, TestItem)

    Shared Sub New()

        Dim dict As New Dictionary(Of String, TestItem)(StringComparer.OrdinalIgnoreCase)

        ' ==============================
        ' Core Test Items
        ' ==============================

        Add(dict, "Input Voltage Turn On", "XP-12-02", "OMS0165")
        Add(dict, "Input Voltage Turn Off", "XP-12-03", "OMS0168")
        Add(dict, "Input Current", "XP-12-04", "OMS0166")
        Add(dict, "No Load Input Power", "XP-12-05", "OMS0167")
        Add(dict, "Inrush Current", "XP-12-08", "OMS0171")
        Add(dict, "Power Factor", "XP-12-09", "OMS0170")
        Add(dict, "Ripple and Noise", "XP-12-12", "OMS0174")
        Add(dict, "Start Up Time", "XP-12-13", "OMS0177")
        Add(dict, "Start Up Rise Time", "XP-12-14", "OMS0175")
        Add(dict, "Hold-Up Time", "XP-12-15", "OMS0179")
        Add(dict, "Turn On Overshoot", "XP-12-16", "OMS0176")
        Add(dict, "Efficiency", "XP-12-17", "OMS0181")
        Add(dict, "Output Voltage Setting Accuracy", "XP-12-18", "OMS0180")
        Add(dict, "Line Regulation", "XP-12-19", "OMS0183")
        Add(dict, "Load Regulation", "XP-12-20", "OMS0182")
        Add(dict, "Over Current Protection", "XP-12-22", "OMS0185")
        Add(dict, "Over Voltage Protection", "XP-12-23", "OMS0186")
        Add(dict, "Transient Response", "XP-12-27", "OMS0190")
        Add(dict, "Inhibit Input Power", "XP-12-06", "OMS0169")
        Add(dict, "Harmonics Current", "XP-12-10", "OMS0173")
        Add(dict, "Temperature Coefficient", "XP-12-36", "OMS0199")
        Add(dict, "Thermal Protection", "XP-12-37", "OMS0201")
        Add(dict, "Initial Set Accuracy", "XP-12-45", "OMS0207")
        Add(dict, "Output Voltage Stability", "XP-12-51", "OMS0214")

        ' ==============================
        ' Items With OMS Only
        ' ==============================

        Add(dict, "AC OK", "", "OMS0191")
        Add(dict, "DC OK", "", "OMS0192")
        Add(dict, "Short Circuit Protection", "", "OMS1608")

        ' ==============================
        ' Items With No Codes (Placeholders)
        ' ==============================

        Add(dict, "Current Limit Setting", "", "")
        Add(dict, "Current Output Reading", "", "")
        Add(dict, "Early Warning Time", "", "")
        Add(dict, "IPROG Accuracy", "", "")
        Add(dict, "IPROG Emulation", "", "")
        Add(dict, "Turn Off Fall Time", "", "")
        Add(dict, "Voltage Output Reading", "", "")
        Add(dict, "Voltage Output Setting", "", "")
        Add(dict, "VPROG Accuracy", "", "")
        Add(dict, "VPROG Emulation", "", "")

        _items = New ReadOnlyDictionary(Of String, TestItem)(dict)

    End Sub

    ' =========================================
    ' Private Add Helper (Prevents Duplicates)
    ' =========================================
    Private Shared Sub Add(dict As Dictionary(Of String, TestItem),
                           name As String,
                           dvt As String,
                           oms As String)

        If dict.ContainsKey(name) Then
            Throw New InvalidOperationException($"Duplicate TestItem name detected: {name}")
        End If

        dict(name) = New TestItem(name, dvt, oms)

    End Sub

    ' =========================================
    ' Public Accessors
    ' =========================================

    Public Shared ReadOnly Property Items As IReadOnlyDictionary(Of String, TestItem)
        Get
            Return _items
        End Get
    End Property

    Public Shared Function GetAll() As IEnumerable(Of TestItem)
        Return _items.Values.OrderBy(Function(x) x.ItemName)
    End Function

    Public Shared Function TryGet(itemName As String,
                                  ByRef item As TestItem) As Boolean

        Return _items.TryGetValue(itemName, item)

    End Function

End Class