Imports System.Threading
Imports Microsoft.Office.Interop
Imports System.Text
Public Class ClassGenReport
    Implements IDisposable
    Dim Myform As Object
    Dim ReportName As String
    Dim myThread As New System.Threading.Thread(AddressOf DoReport)
    Dim mylist As List(Of QueryWorksheet)
    Private myuserinfo As UserInfo = UserInfo.getInstance
    Public Sub New(ByVal myForm As Object, ByVal reportname As String)
        Me.Myform = myForm
        Me.ReportName = reportname
    End Sub

    Public Sub LoadReport(ByVal mylist As List(Of QueryWorksheet))
        If Not myThread.IsAlive Then
            Myform.ToolStripStatusLabel1.Text = ""
            myThread = New Thread(AddressOf DoReport)
            myThread.SetApartmentState(ApartmentState.STA)
            Me.mylist = mylist
            myThread.Start()
        Else
            MessageBox.Show("Please wait until the current process is finished.")
        End If
    End Sub



#Region "IDisposable Support"
    Private disposedValue As Boolean ' To detect redundant calls

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                ' TODO: dispose managed state (managed objects).
            End If

            ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
            ' TODO: set large fields to null.
        End If
        Me.disposedValue = True
    End Sub

    ' TODO: override Finalize() only if Dispose(ByVal disposing As Boolean) above has code to free unmanaged resources.
    'Protected Overrides Sub Finalize()
    '    ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
    '    Dispose(False)
    '    MyBase.Finalize()
    'End Sub

    ' This code added by Visual Basic to correctly implement the disposable pattern.
    Public Sub Dispose() Implements IDisposable.Dispose
        ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub
#End Region

    Sub DoReport()
        ProgressReport(6, "Marquee")
        ProgressReport(1, "Loading Data.")
        Dim sqlstr As String = String.Empty
        Dim sb As New StringBuilder
        'DS = New DataSet

        Dim mymessage As String = String.Empty
        sb.Clear()
        ProgressReport(11, "InitFilter")

        Dim mysaveform As New SaveFileDialog
        mysaveform.FileName = ReportName

        If (mysaveform.ShowDialog() = Windows.Forms.DialogResult.OK) Then
            Dim filename = IO.Path.GetDirectoryName(mysaveform.FileName)
            Dim reportname = IO.Path.GetFileName(mysaveform.FileName)
            Dim datasheet As Integer = 6

            Dim mycallback As FormatReportDelegate = AddressOf Myform.FormattingReport
            Dim PivotCallback As FormatReportDelegate = AddressOf Myform.PivotTable

            'Dim myreport As New ExportToExcel(Myform, mylist, filename, reportname, mycallback, PivotCallback, False, "\Templates\ExcelTemplate.xltx", "A1")
            Dim myreport As New ExportToExcel(Myform, mylist, filename, reportname, mycallback, PivotCallback, False, "\Templates\ExcelTemplate.xltx", "A1")
            myreport.Run(Myform, New EventArgs)
        End If

        ProgressReport(1, "Loading Data.Done!")
        ProgressReport(5, "Continuous")
    End Sub
    Private Sub ProgressReport(ByVal id As Integer, ByVal message As String)
        If Myform.InvokeRequired Then
            Dim d As New ProgressReportDelegate(AddressOf ProgressReport)
            Myform.Invoke(d, New Object() {id, message})
        Else
            Select Case id
                Case 1
                    Myform.ToolStripStatusLabel1.Text = message
                Case 4
                Case 5
                    Myform.ToolStripProgressBar1.Style = ProgressBarStyle.Continuous
                Case 6
                    Myform.ToolStripProgressBar1.Style = ProgressBarStyle.Marquee

            End Select
        End If
    End Sub
End Class
