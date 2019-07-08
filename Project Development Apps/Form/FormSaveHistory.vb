Imports System.Threading

Public Class FormSaveHistory
    Private Enum TxHistoryEnum
        SavePeriod = 0
        DeletePeriod
    End Enum

    Dim myTxHistoryEnum As TxHistoryEnum

    Private Shared myInstance As FormSaveHistory
    Dim myThread As New System.Threading.Thread(AddressOf DoSavePeriod)
    Public Shared Function getInstance()
        If myInstance Is Nothing Then
            myInstance = New FormSaveHistory
        ElseIf myInstance.IsDisposed Then
            myInstance = New FormSaveHistory
        End If
        Return myInstance
    End Function

    Private Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If MessageBox.Show(String.Format("Saving Period {0}. Continue?", String.Format("{0:yyyy-MM}", DateTimePicker1.Value.Date)), "Saving Period", MessageBoxButtons.OKCancel) = Windows.Forms.DialogResult.OK Then
            myTxHistoryEnum = TxHistoryEnum.SavePeriod
            If Not myThread.IsAlive Then
                ToolStripStatusLabel1.Text = ""

                myThread = New Thread(AddressOf DoSavePeriod)
                myThread.SetApartmentState(ApartmentState.STA)
                myThread.Start()
            Else
                MessageBox.Show("Please wait until the current process is finished.")
            End If
        End If

    End Sub

    Sub DoSavePeriod()
        ProgressReport(6, "Marquee")
        ProgressReport(1, "Processing...")

        Dim TxHistory = New TxHistoryAdapter(DateTimePicker1.Value.Date)
        Select Case myTxHistoryEnum
            Case TxHistoryEnum.DeletePeriod
                TxHistory.Delete()
            Case TxHistoryEnum.SavePeriod
                TxHistory.Save()
        End Select


        ProgressReport(1, "Processing...Done!")
        ProgressReport(5, "Continuous")
    End Sub

    Private Sub ProgressReport(ByVal id As Integer, ByVal message As String)
        If Me.InvokeRequired Then
            Dim d As New ProgressReportDelegate(AddressOf ProgressReport)
            Me.Invoke(d, New Object() {id, message})
        Else
            Select Case id
                Case 1
                    ToolStripStatusLabel1.Text = message
                Case 4

            End Select
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If MessageBox.Show(String.Format("Delete Period {0}. Continue?", String.Format("{0:yyyy-MM}", DateTimePicker1.Value.Date)), "Delete Period", MessageBoxButtons.OKCancel) = Windows.Forms.DialogResult.OK Then
            myTxHistoryEnum = TxHistoryEnum.DeletePeriod
            If Not myThread.IsAlive Then
                ToolStripStatusLabel1.Text = ""

                myThread = New Thread(AddressOf DoSavePeriod)
                myThread.SetApartmentState(ApartmentState.STA)
                myThread.Start()
            Else
                MessageBox.Show("Please wait until the current process is finished.")
            End If
        End If
    End Sub
End Class