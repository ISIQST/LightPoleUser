Option Strict Off
Option Explicit On

Public Class Application
    Implements Quasi97.iAddin

    Public WithEvents QuasiEvents As Quasi97.QuasiCls
    Public WithEvents TestSeq As Quasi97.TestSequencer

    Private WithEvents tmrLockOutMon As New System.Windows.Forms.Timer With {.Interval = 1000, .Enabled = False}

    Private Sub tmrLockOutMon_Tick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles tmrLockOutMon.Tick
        If TestStarted Then
            If DateDiff(Microsoft.VisualBasic.DateInterval.Second, LastTestReport, Now) > 30 Then
                If LightPolePTR.GetLight(1) <> 1 Then Call LightPolePTR.LEDOn(1, 1) 'machine idle
            Else
                If LightPolePTR.GetLight(1) <> 0 Then Call LightPolePTR.LEDOn(1, 0) 'machine running
            End If
        End If
    End Sub

    Private Sub QuasiEvents_CurrentHeadTerminate(ByRef hd As Short)
        LastTestReport = Now
    End Sub

    Private Sub QuasiEvents_StartStopDeviceTerminate(ByRef State As Boolean)
        If State Then
            Call LightPolePTR.SetInterval(True, 2, 1)
            Call LightPolePTR.SetInterval(False, 2, 1)
            Call LightPolePTR.LEDOn(2, 1)
            Call LightPolePTR.LEDOn(1, 0)
            LastTestReport = Now
            TestStarted = True
            tmrLockOutMon.Enabled = True
        Else
            LastTestReport = Now
            Call LightPolePTR.LEDOn(2, 0)
            Call LightPolePTR.LEDOn(1, 1)
            TestStarted = False
            tmrLockOutMon.Enabled = False
        End If
    End Sub

    Private Sub TestSeq_ProductionTestComplete() Handles TestSeq.ProductionTestComplete
        Call LightPolePTR.LEDOn(1, 1)
    End Sub

    Public ReadOnly Property CustomStressSupport As Boolean Implements Quasi97.iAddin.CustomStressSupport
        Get
            Return False
        End Get
    End Property

    Public Sub Initialize2(ByRef q As Quasi97.Application) Implements Quasi97.iAddin.Initialize2
        QST = q

        Dim l As Quasi97.clsHardwareOption = CType(QST.HOptionManager.GetPointerByNameNET("LightPole.clsInterface", ""), Quasi97.clsHardwareOption)
        If l Is Nothing Then
            MsgBox("Please add LighPole.clsinterface module to the hardware config list.")
            Return
        Else
            LightPolePTR = CType(l.GetHandle, LightPole.clsInterface)
        End If
        Call LightPolePTR.Initialize(False)
        tmrLockOutMon = New System.Windows.Forms.Timer With {.Interval = 1000, .Enabled = False}

        QuasiEvents = QST.QuasiParameters
        TestSeq = QST.TestSequencer

        If Not QST.QSTHardware.isPowerOn Then Call LightPolePTR.LEDOn(0, 1)
        Call LightPolePTR.LEDOn(3, 1)
        Call LightPolePTR.LEDOn(1, 1)
    End Sub

    Public Property ModuleID As String = "LightPoleUser" Implements Quasi97.iAddin.ModuleID
        
    Public ReadOnly Property QuasiAddIn As Boolean Implements Quasi97.iAddin.QuasiAddIn
        Get
            Return True
        End Get
    End Property

    Public Sub RunCustomStress(sM As Quasi97.clsStress.StressMode, UniqueEventID As Short, ByRef EventParams As String) Implements Quasi97.iAddin.RunCustomStress
        'do nothing
    End Sub

    Public Sub ValidateStressParam(eID As Short, ByRef ParamValue As String, ByRef RetValue As Boolean) Implements Quasi97.iAddin.ValidateStressParam
        'do nothing
    End Sub

#Region "IDisposable Support"
    Private disposedValue As Boolean ' To detect redundant calls

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                tmrLockOutMon.Enabled = False
                TestSeq = Nothing
                QuasiEvents = Nothing
                QST = Nothing
                LightPolePTR = Nothing
            End If

            ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
            ' TODO: set large fields to null.
        End If
        Me.disposedValue = True
    End Sub

    ' TODO: override Finalize() only if Dispose(ByVal disposing As Boolean) above has code to free unmanaged resources.
    Protected Overrides Sub Finalize()
        ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
        Dispose(False)
        MyBase.Finalize()
    End Sub

    ' This code added by Visual Basic to correctly implement the disposable pattern.
    Public Sub Dispose() Implements IDisposable.Dispose
        ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub
#End Region

End Class