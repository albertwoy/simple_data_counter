Imports System
Imports System.IO
Imports System.ComponentModel
Imports System.Threading
Imports System.Data.SqlClient
Imports System.Text.RegularExpressions
Imports Excel = Microsoft.Office.Interop.Excel
Public Class frmdatacounter
    Dim a As New OpenFileDialog
    Dim bytImage() As Byte
    Dim bfile As String
    Dim strFileName As String
    Dim FileSize As UInt32
    Dim rpath As String = Nothing
    Dim t1 As Thread = Nothing

    Private Sub frmdatacounter_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        If Not t1 Is Nothing Then
            t1.Abort()
        End If
    End Sub
    Private Sub frmdatacounter_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Control.CheckForIllegalCrossThreadCalls = False
        t1 = Nothing
        setcolumns()
        lbltotres.Text = Nothing
        lblfile.Text = Nothing
    End Sub
    Sub setcolumns()
        Dim ar As String() = {"Respondent", "age", "gender", "year", "se1", "se2", "se3", "se4", "se5", "se6", "se7", "se8", "se9", "se10", _
            "sf1", "sf2", "sf3", "sf4", "sf5", "sf6", "p1", "p2", "p3", "p4", "p5", "ptime1", "ptime2", "ptime3", "ptime4", "ptime5", "freq", _
            "purpose1", "purpose2", "purpose3", "purpose4", "purpose5", "purpose6", "as1", "as2", "as3", "as4", "as5", "ps1", "ps2", _
            "ps3", "ps4", "ps5", "ps6", "ps7", "ps8", "ps9", "ps10", "ss1", "ss2", "ss3", "ss4", "ss5", "ss6", "ss7", "ss8", "ss9", "ss10", _
            "es1", "es2", "es3", "es4", "es5", "ps1", "ps2", "ps3", "ps4", "ps5", "ps6", "ps7", "ps8", "ps9", "ps10"}

        With dgdc
            .Columns.Clear()
            Dim cname As String = Nothing

            For i = 0 To UBound(ar)
                cname = "cnem" & i

                .Columns.Add(cname, ar(i))

                .Columns(i).AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
                '.Columns(i).AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
            Next
        End With


    End Sub
    Public Function ggend(ByVal ans As String)
        Dim index As Integer = 0
        Dim ar As String() = {"Male", "Female"}
        If Array.IndexOf(ar, ans) >= 0 Then
            index = Array.IndexOf(ar, ans) + 1
        End If
        Return index
    End Function
    Public Function gyr(ByVal ans As String)
        Dim index As Integer = 0
        Dim ar As String() = {"First Year", "Second Year", "Third Year", "Fourth Year", "Fifth Year"}
        If Array.IndexOf(ar, ans) >= 0 Then
            index = Array.IndexOf(ar, ans) + 1
        End If
        Return index
    End Function
    Public Function se1(ByVal ans As String)
        Dim index As Integer = 0
        Dim ar As String() = {"Strongly Disagree", "Disagree", "Agree", "Strongly Agree"}
        If Array.IndexOf(ar, ans) >= 0 Then
            index = Array.IndexOf(ar, ans) '+ 1
        End If
        Return index
    End Function
    Public Function se2(ByVal ans As String)
        Dim index As Integer = 0
        Dim ar As String() = {"Strongly Agree", "Agree", "Disagree", "Strongly Disagree"}
        If Array.IndexOf(ar, ans) >= 0 Then
            index = Array.IndexOf(ar, ans) '+ 1
        End If
        Return index
    End Function
    Public Function sf1(ByVal ans As String)
        Dim index As Integer = 0
        Dim ar As String() = {"Not at all true", "Hardly true", "Moderately true", "Exactly true"}
        If Array.IndexOf(ar, ans) >= 0 Then
            index = Array.IndexOf(ar, ans) + 1
        End If
        Return index
    End Function
    Public Function p1(ByVal ans As String)
        Dim index As Integer = 0
        Dim ar As String() = {"No", "Yes"}
        If Array.IndexOf(ar, ans) >= 0 Then
            index = Array.IndexOf(ar, ans)
        End If
        Return index
    End Function
    Public Function socfreq(ByVal ans As String)
        Dim index As Integer = 0
        Dim ar As String() = {"Rarely", "Sometimes", "Occasionally", "Mostly"}
        If ans = "Occasi" Then
            ans = "Occasionally"
        ElseIf ans = "Someti" Then
            ans = "Sometimes"
        End If

        If Array.IndexOf(ar, ans) >= 0 Then
            index = Array.IndexOf(ar, ans) + 1
        End If
        Return index
    End Function
  

    Private Sub btnbrws_Click(sender As Object, e As EventArgs) Handles btnbrws.Click
        Dim xlApp As Excel.Application
        Dim xlWorkBook As Excel.Workbook
        Dim xlWorkSheet As Excel.Worksheet
        Dim range As Excel.Range
        Dim rCnt As Integer
        Dim cCnt As Integer = 0

        Dim objdets As Object = Nothing
        lbltotres.Text = Nothing
        strFileName = Nothing

        Dim fd As OpenFileDialog = New OpenFileDialog()
        fd.Title = "Open File Dialog"
        fd.InitialDirectory = "C:\"
        fd.Filter = "excel or csv files ( *.xlsx, *.csv, *.xls)| *.xlsx; *.csv; *.xls"
        fd.FilterIndex = 2
        fd.RestoreDirectory = True

        If fd.ShowDialog() = DialogResult.OK Then
            dgdc.Rows.Clear()
            strFileName = fd.FileName
            xlApp = New Excel.Application
            xlWorkBook = xlApp.Workbooks.Open(fd.FileName)
            lblfile.Text = fd.FileName
            Dim wkname As String = Nothing
            wkname = xlWorkBook.Sheets(1).name

            xlWorkSheet = xlWorkBook.Worksheets(wkname)

            range = xlWorkSheet.UsedRange

            pbar1.Value = 0
            pbar1.Maximum = range.Rows.Count - 1
            Dim totres As Integer = 0


            If range.Rows.Count > 1 Then
                Dim pono As String = Nothing
                Dim accno As String = Nothing
                Dim ship2acc As String = Nothing
                Dim rmk As String = Nothing
                Dim rdd As Date

                Dim rw As Integer = 0
                Dim xc As Integer = 0


                Dim mygurd As String = Nothing
                Dim myawardno As String = Nothing
                Dim encoder As String = Nothing
                Dim dtres As String = Nothing
                Dim rwx As Integer = 0
                setcolumns()
                
                For rCnt = 2 To range.Rows.Count

                    ' For cCnt = 1 To range.Columns.Count

                    xc += 1
                    'myawardno = Nothing
                    'mygurd = Nothing
                    'encoder = Nothing
                    'dtres = Nothing
                    If CType(range.Cells(rCnt, 2), Excel.Range).Value <> Nothing Then
                        If Trim(CType(range.Cells(rCnt, 2), Excel.Range).Value) <> " " And CType(range.Cells(rCnt, 2), Excel.Range).Value <> Nothing Then
                            'myawardno = CType(range.Cells(rCnt, 2), Excel.Range).Value
                            'mygurd = CType(range.Cells(rCnt, 10), Excel.Range).Value
                            'dtres = CType(range.Cells(rCnt, 14), Excel.Range).Value
                            With dgdc
                                .RowCount += 1
                                rwx = .RowCount - 1
                                'Dim ar As String() = {"Respondent", "age", "gender", "year", "se1", "se2", "se3", "se4", "se5", "se6", "se7", "se8", "se9", "se10", _
                                '    "sf1", "sf2", "sf3", "sf4", "sf5", "sf6", "p1", "p2", "p3", "p4", "p5", "ptime1", "ptime2", "ptime3", "ptime4", "ptime5", "freq", _
                                '    "purpose1", "purpose2", "purpose3", "purpose4", "purpose5", "purpose6", "as1", "as2", "as3", "as4", "as5", "ps1", "ps2", _
                                '    "ps3", "ps4", "ps5", "ps6", "ps7", "ps8", "ps9", "ps10", "ss1", "ss2", "ss3", "ss4", "ss5", "ss6", "ss7", "ss8", "ss9", "ss10", _
                                '    "es1", "es2", "es3", "es4", "es5", "ps1", "ps2", "ps3", "ps4", "ps5", "ps6", "ps7", "ps8", "ps9", "ps10"}
                                .Item(0, rwx).Value = xc 'CType(range.Cells(rCnt, 2), Excel.Range).Value
                                .Item(1, rwx).Value = CType(range.Cells(rCnt, 2), Excel.Range).Value 'AGE
                                .Item(2, rwx).Value = ggend(CType(range.Cells(rCnt, 3), Excel.Range).Value)
                                .Item(3, rwx).Value = gyr(CType(range.Cells(rCnt, 4), Excel.Range).Value)
                                'se1
                                .Item(4, rwx).Value = se1(CType(range.Cells(rCnt, 5), Excel.Range).Value)
                                .Item(5, rwx).Value = se1(CType(range.Cells(rCnt, 6), Excel.Range).Value)
                                .Item(6, rwx).Value = se2(CType(range.Cells(rCnt, 7), Excel.Range).Value) 'reverse
                                .Item(7, rwx).Value = se1(CType(range.Cells(rCnt, 8), Excel.Range).Value)
                                .Item(8, rwx).Value = se2(CType(range.Cells(rCnt, 9), Excel.Range).Value) 'reverse
                                .Item(9, rwx).Value = se1(CType(range.Cells(rCnt, 10), Excel.Range).Value)
                                .Item(10, rwx).Value = se1(CType(range.Cells(rCnt, 11), Excel.Range).Value)
                                .Item(11, rwx).Value = se2(CType(range.Cells(rCnt, 12), Excel.Range).Value) 'reverse
                                .Item(12, rwx).Value = se2(CType(range.Cells(rCnt, 13), Excel.Range).Value) 'reverse
                                .Item(13, rwx).Value = se2(CType(range.Cells(rCnt, 14), Excel.Range).Value) 'reverse
                                'se10
                                'sf1
                                .Item(14, rwx).Value = sf1(CType(range.Cells(rCnt, 15), Excel.Range).Value)
                                .Item(15, rwx).Value = sf1(CType(range.Cells(rCnt, 16), Excel.Range).Value)
                                .Item(16, rwx).Value = sf1(CType(range.Cells(rCnt, 17), Excel.Range).Value)
                                .Item(17, rwx).Value = sf1(CType(range.Cells(rCnt, 18), Excel.Range).Value)
                                .Item(18, rwx).Value = sf1(CType(range.Cells(rCnt, 19), Excel.Range).Value)
                                .Item(19, rwx).Value = sf1(CType(range.Cells(rCnt, 20), Excel.Range).Value)
                                'sf6
                                'p1
                                .Item(20, rwx).Value = p1(CType(range.Cells(rCnt, 21), Excel.Range).Value)
                                .Item(21, rwx).Value = p1(CType(range.Cells(rCnt, 23), Excel.Range).Value)
                                .Item(22, rwx).Value = p1(CType(range.Cells(rCnt, 25), Excel.Range).Value)
                                .Item(23, rwx).Value = p1(CType(range.Cells(rCnt, 27), Excel.Range).Value)
                                .Item(24, rwx).Value = p1(CType(range.Cells(rCnt, 29), Excel.Range).Value)
                                'p5
                                'ptime1
                                .Item(25, rwx).Value = CType(range.Cells(rCnt, 22), Excel.Range).Value
                                .Item(26, rwx).Value = CType(range.Cells(rCnt, 24), Excel.Range).Value
                                .Item(27, rwx).Value = CType(range.Cells(rCnt, 26), Excel.Range).Value
                                .Item(28, rwx).Value = CType(range.Cells(rCnt, 28), Excel.Range).Value
                                .Item(29, rwx).Value = CType(range.Cells(rCnt, 30), Excel.Range).Value
                                'ptime5
                                'freq
                                .Item(30, rwx).Value = socfreq(Strings.Left(CType(range.Cells(rCnt, 31), Excel.Range).Value, 6))
                                'purpose1
                                .Item(31, rwx).Value = p1(CType(range.Cells(rCnt, 32), Excel.Range).Value)
                                .Item(32, rwx).Value = p1(CType(range.Cells(rCnt, 33), Excel.Range).Value)
                                .Item(33, rwx).Value = p1(CType(range.Cells(rCnt, 34), Excel.Range).Value)
                                .Item(34, rwx).Value = p1(CType(range.Cells(rCnt, 35), Excel.Range).Value)
                                .Item(35, rwx).Value = p1(CType(range.Cells(rCnt, 36), Excel.Range).Value)
                                .Item(36, rwx).Value = p1(CType(range.Cells(rCnt, 37), Excel.Range).Value)
                                'purpose5
                                'as1 
                                .Item(37, rwx).Value = Strings.Left(CType(range.Cells(rCnt, 38), Excel.Range).Value, 1)
                                .Item(38, rwx).Value = Strings.Left(CType(range.Cells(rCnt, 39), Excel.Range).Value, 1)
                                .Item(39, rwx).Value = Strings.Left(CType(range.Cells(rCnt, 40), Excel.Range).Value, 1)
                                .Item(40, rwx).Value = Strings.Left(CType(range.Cells(rCnt, 41), Excel.Range).Value, 1)
                                .Item(41, rwx).Value = Strings.Left(CType(range.Cells(rCnt, 42), Excel.Range).Value, 1)
                                .Item(42, rwx).Value = Strings.Left(CType(range.Cells(rCnt, 43), Excel.Range).Value, 1)
                                .Item(43, rwx).Value = Strings.Left(CType(range.Cells(rCnt, 44), Excel.Range).Value, 1)
                                .Item(44, rwx).Value = Strings.Left(CType(range.Cells(rCnt, 45), Excel.Range).Value, 1)
                                .Item(45, rwx).Value = Strings.Left(CType(range.Cells(rCnt, 46), Excel.Range).Value, 1)
                                .Item(46, rwx).Value = Strings.Left(CType(range.Cells(rCnt, 47), Excel.Range).Value, 1)
                                .Item(47, rwx).Value = Strings.Left(CType(range.Cells(rCnt, 48), Excel.Range).Value, 1)
                                .Item(48, rwx).Value = Strings.Left(CType(range.Cells(rCnt, 49), Excel.Range).Value, 1)
                                .Item(49, rwx).Value = Strings.Left(CType(range.Cells(rCnt, 50), Excel.Range).Value, 1)
                                .Item(50, rwx).Value = Strings.Left(CType(range.Cells(rCnt, 51), Excel.Range).Value, 1)
                                .Item(51, rwx).Value = Strings.Left(CType(range.Cells(rCnt, 52), Excel.Range).Value, 1)
                                .Item(52, rwx).Value = Strings.Left(CType(range.Cells(rCnt, 53), Excel.Range).Value, 1)
                                .Item(53, rwx).Value = Strings.Left(CType(range.Cells(rCnt, 54), Excel.Range).Value, 1)
                                .Item(54, rwx).Value = Strings.Left(CType(range.Cells(rCnt, 55), Excel.Range).Value, 1)
                                .Item(55, rwx).Value = Strings.Left(CType(range.Cells(rCnt, 56), Excel.Range).Value, 1)
                                .Item(56, rwx).Value = Strings.Left(CType(range.Cells(rCnt, 57), Excel.Range).Value, 1)
                                .Item(57, rwx).Value = Strings.Left(CType(range.Cells(rCnt, 58), Excel.Range).Value, 1)
                                .Item(58, rwx).Value = Strings.Left(CType(range.Cells(rCnt, 59), Excel.Range).Value, 1)
                                .Item(59, rwx).Value = Strings.Left(CType(range.Cells(rCnt, 60), Excel.Range).Value, 1)
                                .Item(60, rwx).Value = Strings.Left(CType(range.Cells(rCnt, 61), Excel.Range).Value, 1)
                                .Item(61, rwx).Value = Strings.Left(CType(range.Cells(rCnt, 62), Excel.Range).Value, 1)
                                .Item(62, rwx).Value = Strings.Left(CType(range.Cells(rCnt, 63), Excel.Range).Value, 1)
                                .Item(63, rwx).Value = Strings.Left(CType(range.Cells(rCnt, 64), Excel.Range).Value, 1)
                                .Item(64, rwx).Value = Strings.Left(CType(range.Cells(rCnt, 65), Excel.Range).Value, 1)
                                .Item(65, rwx).Value = Strings.Left(CType(range.Cells(rCnt, 66), Excel.Range).Value, 1)
                                .Item(66, rwx).Value = Strings.Left(CType(range.Cells(rCnt, 67), Excel.Range).Value, 1)
                                .Item(67, rwx).Value = Strings.Left(CType(range.Cells(rCnt, 68), Excel.Range).Value, 1)
                                .Item(68, rwx).Value = Strings.Left(CType(range.Cells(rCnt, 69), Excel.Range).Value, 1)
                                .Item(69, rwx).Value = Strings.Left(CType(range.Cells(rCnt, 70), Excel.Range).Value, 1)
                                .Item(70, rwx).Value = Strings.Left(CType(range.Cells(rCnt, 71), Excel.Range).Value, 1)
                                .Item(71, rwx).Value = Strings.Left(CType(range.Cells(rCnt, 72), Excel.Range).Value, 1)
                                .Item(72, rwx).Value = Strings.Left(CType(range.Cells(rCnt, 73), Excel.Range).Value, 1)
                                .Item(73, rwx).Value = Strings.Left(CType(range.Cells(rCnt, 74), Excel.Range).Value, 1)
                                .Item(74, rwx).Value = Strings.Left(CType(range.Cells(rCnt, 75), Excel.Range).Value, 1)
                                .Item(75, rwx).Value = Strings.Left(CType(range.Cells(rCnt, 76), Excel.Range).Value, 1)
                                .Item(76, rwx).Value = Strings.Left(CType(range.Cells(rCnt, 77), Excel.Range).Value, 1)
                                'ps10

                                ''Dim ar As String() = {"Respondent", "age", "gender", "year", "se1", "se2", "se3", "se4", "se5", "se6", "se7", "se8", "se9", "se10", _
                                ''    "sf1", "sf2", "sf3", "sf4", "sf5", "sf6", "p1", "p2", "p3", "p4", "p5", "ptime1", "ptime2", "ptime3", "ptime4", "ptime5", "freq", _
                                ''    "purpose1", "purpose2", "purpose3", "purpose4", "purpose5", "purpose6", "as1", "as2", "as3", "as4", "as5", "ps1", "ps2", _
                                ''    "ps3", "ps4", "ps5", "ps6", "ps7", "ps8", "ps9", "ps10", "ss1", "ss2", "ss3", "ss4", "ss5", "ss6", "ss7", "ss8", "ss9", "ss10", _
                                ''    "es1", "es2", "es3", "es4", "es5", "ps1", "ps2", "ps3", "ps4", "ps5", "ps6", "ps7", "ps8", "ps9", "ps10"}
                                '.Item(0, rwx).Value = xc 'CType(range.Cells(rCnt, 2), Excel.Range).Value
                                '.Item(1, rwx).Value = CType(range.Cells(rCnt, 3), Excel.Range).Value 'AGE
                                '.Item(2, rwx).Value = ggend(CType(range.Cells(rCnt, 4), Excel.Range).Value)
                                '.Item(3, rwx).Value = gyr(CType(range.Cells(rCnt, 5), Excel.Range).Value)
                                ''se1
                                '.Item(4, rwx).Value = se1(CType(range.Cells(rCnt, 6), Excel.Range).Value)
                                '.Item(5, rwx).Value = se1(CType(range.Cells(rCnt, 7), Excel.Range).Value)
                                '.Item(6, rwx).Value = se1(CType(range.Cells(rCnt, 8), Excel.Range).Value)
                                '.Item(7, rwx).Value = se1(CType(range.Cells(rCnt, 9), Excel.Range).Value)
                                '.Item(8, rwx).Value = se1(CType(range.Cells(rCnt, 10), Excel.Range).Value)
                                '.Item(9, rwx).Value = se1(CType(range.Cells(rCnt, 11), Excel.Range).Value)
                                '.Item(10, rwx).Value = se1(CType(range.Cells(rCnt, 12), Excel.Range).Value)
                                '.Item(11, rwx).Value = se1(CType(range.Cells(rCnt, 13), Excel.Range).Value)
                                '.Item(12, rwx).Value = se1(CType(range.Cells(rCnt, 14), Excel.Range).Value)
                                '.Item(13, rwx).Value = se1(CType(range.Cells(rCnt, 15), Excel.Range).Value)
                                ''se10
                                ''sf1
                                '.Item(14, rwx).Value = sf1(CType(range.Cells(rCnt, 16), Excel.Range).Value)
                                '.Item(15, rwx).Value = sf1(CType(range.Cells(rCnt, 17), Excel.Range).Value)
                                '.Item(16, rwx).Value = sf1(CType(range.Cells(rCnt, 18), Excel.Range).Value)
                                '.Item(17, rwx).Value = sf1(CType(range.Cells(rCnt, 19), Excel.Range).Value)
                                '.Item(18, rwx).Value = sf1(CType(range.Cells(rCnt, 20), Excel.Range).Value)
                                '.Item(19, rwx).Value = sf1(CType(range.Cells(rCnt, 21), Excel.Range).Value)
                                ''sf6
                                ''p1
                                '.Item(20, rwx).Value = p1(CType(range.Cells(rCnt, 22), Excel.Range).Value)
                                '.Item(21, rwx).Value = p1(CType(range.Cells(rCnt, 24), Excel.Range).Value)
                                '.Item(22, rwx).Value = p1(CType(range.Cells(rCnt, 26), Excel.Range).Value)
                                '.Item(23, rwx).Value = p1(CType(range.Cells(rCnt, 28), Excel.Range).Value)
                                '.Item(24, rwx).Value = p1(CType(range.Cells(rCnt, 30), Excel.Range).Value)
                                ''p5
                                ''ptime1
                                '.Item(25, rwx).Value = CType(range.Cells(rCnt, 23), Excel.Range).Value
                                '.Item(26, rwx).Value = CType(range.Cells(rCnt, 25), Excel.Range).Value
                                '.Item(27, rwx).Value = CType(range.Cells(rCnt, 27), Excel.Range).Value
                                '.Item(28, rwx).Value = CType(range.Cells(rCnt, 29), Excel.Range).Value
                                '.Item(29, rwx).Value = CType(range.Cells(rCnt, 31), Excel.Range).Value
                                ''ptime5
                                ''freq
                                '.Item(30, rwx).Value = socfreq(Strings.Left(CType(range.Cells(rCnt, 32), Excel.Range).Value, 6))
                                ''purpose1
                                '.Item(31, rwx).Value = p1(CType(range.Cells(rCnt, 33), Excel.Range).Value)
                                '.Item(32, rwx).Value = p1(CType(range.Cells(rCnt, 34), Excel.Range).Value)
                                '.Item(33, rwx).Value = p1(CType(range.Cells(rCnt, 35), Excel.Range).Value)
                                '.Item(34, rwx).Value = p1(CType(range.Cells(rCnt, 36), Excel.Range).Value)
                                '.Item(35, rwx).Value = p1(CType(range.Cells(rCnt, 37), Excel.Range).Value)
                                '.Item(36, rwx).Value = p1(CType(range.Cells(rCnt, 38), Excel.Range).Value)
                                ''purpose5
                                ''as1 
                                '.Item(37, rwx).Value = Strings.Left(CType(range.Cells(rCnt, 39), Excel.Range).Value, 1)
                                '.Item(38, rwx).Value = Strings.Left(CType(range.Cells(rCnt, 40), Excel.Range).Value, 1)
                                '.Item(39, rwx).Value = Strings.Left(CType(range.Cells(rCnt, 41), Excel.Range).Value, 1)
                                '.Item(40, rwx).Value = Strings.Left(CType(range.Cells(rCnt, 42), Excel.Range).Value, 1)
                                '.Item(41, rwx).Value = Strings.Left(CType(range.Cells(rCnt, 43), Excel.Range).Value, 1)
                                '.Item(42, rwx).Value = Strings.Left(CType(range.Cells(rCnt, 44), Excel.Range).Value, 1)
                                '.Item(43, rwx).Value = Strings.Left(CType(range.Cells(rCnt, 45), Excel.Range).Value, 1)
                                '.Item(44, rwx).Value = Strings.Left(CType(range.Cells(rCnt, 46), Excel.Range).Value, 1)
                                '.Item(45, rwx).Value = Strings.Left(CType(range.Cells(rCnt, 47), Excel.Range).Value, 1)
                                '.Item(46, rwx).Value = Strings.Left(CType(range.Cells(rCnt, 48), Excel.Range).Value, 1)
                                '.Item(47, rwx).Value = Strings.Left(CType(range.Cells(rCnt, 49), Excel.Range).Value, 1)
                                '.Item(48, rwx).Value = Strings.Left(CType(range.Cells(rCnt, 50), Excel.Range).Value, 1)
                                '.Item(49, rwx).Value = Strings.Left(CType(range.Cells(rCnt, 51), Excel.Range).Value, 1)
                                '.Item(50, rwx).Value = Strings.Left(CType(range.Cells(rCnt, 52), Excel.Range).Value, 1)
                                '.Item(51, rwx).Value = Strings.Left(CType(range.Cells(rCnt, 53), Excel.Range).Value, 1)
                                '.Item(52, rwx).Value = Strings.Left(CType(range.Cells(rCnt, 54), Excel.Range).Value, 1)
                                '.Item(53, rwx).Value = Strings.Left(CType(range.Cells(rCnt, 55), Excel.Range).Value, 1)
                                '.Item(54, rwx).Value = Strings.Left(CType(range.Cells(rCnt, 56), Excel.Range).Value, 1)
                                '.Item(55, rwx).Value = Strings.Left(CType(range.Cells(rCnt, 57), Excel.Range).Value, 1)
                                '.Item(56, rwx).Value = Strings.Left(CType(range.Cells(rCnt, 58), Excel.Range).Value, 1)
                                '.Item(57, rwx).Value = Strings.Left(CType(range.Cells(rCnt, 59), Excel.Range).Value, 1)
                                '.Item(58, rwx).Value = Strings.Left(CType(range.Cells(rCnt, 60), Excel.Range).Value, 1)
                                '.Item(59, rwx).Value = Strings.Left(CType(range.Cells(rCnt, 61), Excel.Range).Value, 1)
                                '.Item(60, rwx).Value = Strings.Left(CType(range.Cells(rCnt, 62), Excel.Range).Value, 1)
                                '.Item(61, rwx).Value = Strings.Left(CType(range.Cells(rCnt, 63), Excel.Range).Value, 1)
                                '.Item(62, rwx).Value = Strings.Left(CType(range.Cells(rCnt, 64), Excel.Range).Value, 1)
                                '.Item(63, rwx).Value = Strings.Left(CType(range.Cells(rCnt, 65), Excel.Range).Value, 1)
                                '.Item(64, rwx).Value = Strings.Left(CType(range.Cells(rCnt, 66), Excel.Range).Value, 1)
                                '.Item(65, rwx).Value = Strings.Left(CType(range.Cells(rCnt, 67), Excel.Range).Value, 1)
                                '.Item(66, rwx).Value = Strings.Left(CType(range.Cells(rCnt, 68), Excel.Range).Value, 1)
                                '.Item(67, rwx).Value = Strings.Left(CType(range.Cells(rCnt, 69), Excel.Range).Value, 1)
                                '.Item(68, rwx).Value = Strings.Left(CType(range.Cells(rCnt, 70), Excel.Range).Value, 1)
                                '.Item(69, rwx).Value = Strings.Left(CType(range.Cells(rCnt, 71), Excel.Range).Value, 1)
                                '.Item(70, rwx).Value = Strings.Left(CType(range.Cells(rCnt, 72), Excel.Range).Value, 1)
                                '.Item(71, rwx).Value = Strings.Left(CType(range.Cells(rCnt, 73), Excel.Range).Value, 1)
                                '.Item(72, rwx).Value = Strings.Left(CType(range.Cells(rCnt, 74), Excel.Range).Value, 1)
                                '.Item(73, rwx).Value = Strings.Left(CType(range.Cells(rCnt, 75), Excel.Range).Value, 1)
                                '.Item(74, rwx).Value = Strings.Left(CType(range.Cells(rCnt, 76), Excel.Range).Value, 1)
                                '.Item(75, rwx).Value = Strings.Left(CType(range.Cells(rCnt, 77), Excel.Range).Value, 1)
                                '.Item(76, rwx).Value = Strings.Left(CType(range.Cells(rCnt, 78), Excel.Range).Value, 1)
                                ''ps10
                            End With
                        End If

                    End If

                    totres += 1
                    pbar1.Value += 1
                    pbar1.Refresh()

                Next
                lbltotres.Text = totres
                MsgBox("Importing template data complete!", 64, "Imported Successfully")
            Else
                MsgBox("Invalid Template Format/ Insufficient Rows!", 48, "Please check format")
            End If


            xlWorkBook.Close()
            xlApp.Quit()

            releaseObject(xlApp)
            releaseObject(xlWorkBook)
            releaseObject(xlWorkSheet)
        End If
    End Sub
    Private Sub releaseObject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub


    Private Sub btnxcel_Click(sender As Object, e As EventArgs) Handles btnxcel.Click
        With dgdc
            If .RowCount > 0 Then
                If dgdc.Rows.Count > 0 Then

                    pbar1.Refresh()
                    pbar1.Maximum = dgdc.RowCount
                    Me.Cursor = Cursors.AppStarting
                    '  dgv1.Visible = True
                    t1 = Nothing
                    If t1 Is Nothing Then
                        t1 = New Thread(Sub() startexport())
                        t1.IsBackground = True
                        t1.Start()
                    End If

                Else
                    t1 = Nothing
                End If

            End If

        End With
    End Sub
    Sub startexport()
        With dgdc
            pbar1.Value = 0
            pbar1.Refresh()
            rpath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) & "\pilotdata"
            If (Not System.IO.Directory.Exists(rpath)) Then
                System.IO.Directory.CreateDirectory(rpath)
            End If

            Dim xlApp As Excel.Application
            Dim xlWorkBook As Excel.Workbook
            Dim xlWorkSheet As Excel.Worksheet
            Dim misValue As Object = System.Reflection.Missing.Value
            Dim nstrfilename As String = Nothing

            nstrfilename = Application.StartupPath + "\pilotdata.xlsx"
            xlApp = New Excel.ApplicationClass
            xlWorkBook = xlApp.Workbooks.Open(nstrfilename)
            xlWorkSheet = xlWorkBook.Worksheets(1)

            Dim x As Integer = 0
            Dim r As Integer = 0

            Dim jcols As Integer = 1
            Dim jrow As Integer = 2

            For i = 0 To .RowCount - 1
                jcols = 1
                For j = 0 To .Columns.Count - 1
                    .CurrentCell = .Rows(i).Cells(0)
                    xlWorkSheet.Cells(jrow, jcols).value = .Item(j, i).Value
                    jcols += 1
                Next
                jrow += 1
                pbar1.Value += 1
                pbar1.Refresh()
                x += 1
            Next
            Dim newpath As String = Nothing

            newpath = rpath + "\DataCounter.xlsx"

            xlWorkBook.SaveAs(newpath)
            xlWorkBook.Close(True, misValue, misValue)
            xlApp.Quit()
            MsgBox("Export Complete", 64, "Exporting Data Completed")
            Process.Start(newpath)

            releaseObject(xlWorkSheet)
            releaseObject(xlWorkBook)
            releaseObject(xlApp)

            Me.Cursor = Cursors.Default
            'xlBook.Close()
            'xlApp.Workbooks.Close()
            'xlApp.Quit()
        End With

    End Sub
    Private Sub btngo_Click(sender As Object, e As EventArgs) Handles btngo.Click
        Dim xlApp As Excel.Application
        Dim xlWorkBook As Excel.Workbook
        Dim xlWorkSheet As Excel.Worksheet
        Dim range As Excel.Range
        Dim rCnt As Integer
        Dim cCnt As Integer = 0
        Dim objdets As Object = Nothing
        lbltotres.Text = Nothing
        strFileName = Nothing
        dgdc.Rows.Clear()

        If lblfile.Text <> Nothing Then
            strFileName = lblfile.Text
            xlApp = New Excel.Application
            xlWorkBook = xlApp.Workbooks.Open(lblfile.Text)
            Dim wkname As String = Nothing
            wkname = xlWorkBook.Sheets(1).name

            xlWorkSheet = xlWorkBook.Worksheets(wkname)

            range = xlWorkSheet.UsedRange

            pbar1.Value = 0
            pbar1.Maximum = range.Rows.Count
            Dim totres As Integer = 0


            If range.Rows.Count > 1 Then
                Dim pono As String = Nothing
                Dim accno As String = Nothing
                Dim ship2acc As String = Nothing
                Dim rmk As String = Nothing

                Dim rw As Integer = 0
                Dim xc As Integer = 0


                Dim mygurd As String = Nothing
                Dim myawardno As String = Nothing
                Dim encoder As String = Nothing
                Dim dtres As String = Nothing
                Dim rwx As Integer = 0
                setcolumns()

                For rCnt = 2 To range.Rows.Count
                    xc += 1

                    If CType(range.Cells(rCnt, 1), Excel.Range).Value <> Nothing Then
                        If Trim(CType(range.Cells(rCnt, 1), Excel.Range).Value) <> " " And CType(range.Cells(rCnt, 1), Excel.Range).Value <> Nothing Then
            
                            With dgdc
                                .RowCount += 1
                                rwx = .RowCount - 1

                                .Item(0, rwx).Value = xc 'CType(range.Cells(rCnt, 2), Excel.Range).Value
                                .Item(1, rwx).Value = CType(range.Cells(rCnt, 2), Excel.Range).Value 'AGE
                                .Item(2, rwx).Value = ggend(CType(range.Cells(rCnt, 3), Excel.Range).Value)
                                .Item(3, rwx).Value = gyr(CType(range.Cells(rCnt, 4), Excel.Range).Value)
                                'se1
                                .Item(4, rwx).Value = se1(CType(range.Cells(rCnt, 5), Excel.Range).Value)
                                .Item(5, rwx).Value = se1(CType(range.Cells(rCnt, 6), Excel.Range).Value)
                                .Item(6, rwx).Value = se2(CType(range.Cells(rCnt, 7), Excel.Range).Value) 'reverse
                                .Item(7, rwx).Value = se1(CType(range.Cells(rCnt, 8), Excel.Range).Value)
                                .Item(8, rwx).Value = se2(CType(range.Cells(rCnt, 9), Excel.Range).Value) 'reverse
                                .Item(9, rwx).Value = se1(CType(range.Cells(rCnt, 10), Excel.Range).Value)
                                .Item(10, rwx).Value = se1(CType(range.Cells(rCnt, 11), Excel.Range).Value)
                                .Item(11, rwx).Value = se2(CType(range.Cells(rCnt, 12), Excel.Range).Value) 'reverse
                                .Item(12, rwx).Value = se2(CType(range.Cells(rCnt, 13), Excel.Range).Value) 'reverse
                                .Item(13, rwx).Value = se2(CType(range.Cells(rCnt, 14), Excel.Range).Value) 'reverse
                                'se10
                                'sf1
                                .Item(14, rwx).Value = sf1(CType(range.Cells(rCnt, 15), Excel.Range).Value)
                                .Item(15, rwx).Value = sf1(CType(range.Cells(rCnt, 16), Excel.Range).Value)
                                .Item(16, rwx).Value = sf1(CType(range.Cells(rCnt, 17), Excel.Range).Value)
                                .Item(17, rwx).Value = sf1(CType(range.Cells(rCnt, 18), Excel.Range).Value)
                                .Item(18, rwx).Value = sf1(CType(range.Cells(rCnt, 19), Excel.Range).Value)
                                .Item(19, rwx).Value = sf1(CType(range.Cells(rCnt, 20), Excel.Range).Value)
                                'sf6
                                'p1
                                .Item(20, rwx).Value = p1(CType(range.Cells(rCnt, 21), Excel.Range).Value)
                                .Item(21, rwx).Value = p1(CType(range.Cells(rCnt, 23), Excel.Range).Value)
                                .Item(22, rwx).Value = p1(CType(range.Cells(rCnt, 25), Excel.Range).Value)
                                .Item(23, rwx).Value = p1(CType(range.Cells(rCnt, 27), Excel.Range).Value)
                                .Item(24, rwx).Value = p1(CType(range.Cells(rCnt, 29), Excel.Range).Value)
                                'p5
                                'ptime1
                                .Item(25, rwx).Value = CType(range.Cells(rCnt, 22), Excel.Range).Value
                                .Item(26, rwx).Value = CType(range.Cells(rCnt, 24), Excel.Range).Value
                                .Item(27, rwx).Value = CType(range.Cells(rCnt, 26), Excel.Range).Value
                                .Item(28, rwx).Value = CType(range.Cells(rCnt, 28), Excel.Range).Value
                                .Item(29, rwx).Value = CType(range.Cells(rCnt, 30), Excel.Range).Value
                                'ptime5
                                'freq
                                .Item(30, rwx).Value = socfreq(Strings.Left(CType(range.Cells(rCnt, 31), Excel.Range).Value, 6))
                                'purpose1
                                .Item(31, rwx).Value = p1(CType(range.Cells(rCnt, 32), Excel.Range).Value)
                                .Item(32, rwx).Value = p1(CType(range.Cells(rCnt, 33), Excel.Range).Value)
                                .Item(33, rwx).Value = p1(CType(range.Cells(rCnt, 34), Excel.Range).Value)
                                .Item(34, rwx).Value = p1(CType(range.Cells(rCnt, 35), Excel.Range).Value)
                                .Item(35, rwx).Value = p1(CType(range.Cells(rCnt, 36), Excel.Range).Value)
                                .Item(36, rwx).Value = p1(CType(range.Cells(rCnt, 37), Excel.Range).Value)
                                'purpose5
                                'as1 
                                .Item(37, rwx).Value = Strings.Left(CType(range.Cells(rCnt, 38), Excel.Range).Value, 1)
                                .Item(38, rwx).Value = Strings.Left(CType(range.Cells(rCnt, 39), Excel.Range).Value, 1)
                                .Item(39, rwx).Value = Strings.Left(CType(range.Cells(rCnt, 40), Excel.Range).Value, 1)
                                .Item(40, rwx).Value = Strings.Left(CType(range.Cells(rCnt, 41), Excel.Range).Value, 1)
                                .Item(41, rwx).Value = Strings.Left(CType(range.Cells(rCnt, 42), Excel.Range).Value, 1)
                                .Item(42, rwx).Value = Strings.Left(CType(range.Cells(rCnt, 43), Excel.Range).Value, 1)
                                .Item(43, rwx).Value = Strings.Left(CType(range.Cells(rCnt, 44), Excel.Range).Value, 1)
                                .Item(44, rwx).Value = Strings.Left(CType(range.Cells(rCnt, 45), Excel.Range).Value, 1)
                                .Item(45, rwx).Value = Strings.Left(CType(range.Cells(rCnt, 46), Excel.Range).Value, 1)
                                .Item(46, rwx).Value = Strings.Left(CType(range.Cells(rCnt, 47), Excel.Range).Value, 1)
                                .Item(47, rwx).Value = Strings.Left(CType(range.Cells(rCnt, 48), Excel.Range).Value, 1)
                                .Item(48, rwx).Value = Strings.Left(CType(range.Cells(rCnt, 49), Excel.Range).Value, 1)
                                .Item(49, rwx).Value = Strings.Left(CType(range.Cells(rCnt, 50), Excel.Range).Value, 1)
                                .Item(50, rwx).Value = Strings.Left(CType(range.Cells(rCnt, 51), Excel.Range).Value, 1)
                                .Item(51, rwx).Value = Strings.Left(CType(range.Cells(rCnt, 52), Excel.Range).Value, 1)
                                .Item(52, rwx).Value = Strings.Left(CType(range.Cells(rCnt, 53), Excel.Range).Value, 1)
                                .Item(53, rwx).Value = Strings.Left(CType(range.Cells(rCnt, 54), Excel.Range).Value, 1)
                                .Item(54, rwx).Value = Strings.Left(CType(range.Cells(rCnt, 55), Excel.Range).Value, 1)
                                .Item(55, rwx).Value = Strings.Left(CType(range.Cells(rCnt, 56), Excel.Range).Value, 1)
                                .Item(56, rwx).Value = Strings.Left(CType(range.Cells(rCnt, 57), Excel.Range).Value, 1)
                                .Item(57, rwx).Value = Strings.Left(CType(range.Cells(rCnt, 58), Excel.Range).Value, 1)
                                .Item(58, rwx).Value = Strings.Left(CType(range.Cells(rCnt, 59), Excel.Range).Value, 1)
                                .Item(59, rwx).Value = Strings.Left(CType(range.Cells(rCnt, 60), Excel.Range).Value, 1)
                                .Item(60, rwx).Value = Strings.Left(CType(range.Cells(rCnt, 61), Excel.Range).Value, 1)
                                .Item(61, rwx).Value = Strings.Left(CType(range.Cells(rCnt, 62), Excel.Range).Value, 1)
                                .Item(62, rwx).Value = Strings.Left(CType(range.Cells(rCnt, 63), Excel.Range).Value, 1)
                                .Item(63, rwx).Value = Strings.Left(CType(range.Cells(rCnt, 64), Excel.Range).Value, 1)
                                .Item(64, rwx).Value = Strings.Left(CType(range.Cells(rCnt, 65), Excel.Range).Value, 1)
                                .Item(65, rwx).Value = Strings.Left(CType(range.Cells(rCnt, 66), Excel.Range).Value, 1)
                                .Item(66, rwx).Value = Strings.Left(CType(range.Cells(rCnt, 67), Excel.Range).Value, 1)
                                .Item(67, rwx).Value = Strings.Left(CType(range.Cells(rCnt, 68), Excel.Range).Value, 1)
                                .Item(68, rwx).Value = Strings.Left(CType(range.Cells(rCnt, 69), Excel.Range).Value, 1)
                                .Item(69, rwx).Value = Strings.Left(CType(range.Cells(rCnt, 70), Excel.Range).Value, 1)
                                .Item(70, rwx).Value = Strings.Left(CType(range.Cells(rCnt, 71), Excel.Range).Value, 1)
                                .Item(71, rwx).Value = Strings.Left(CType(range.Cells(rCnt, 72), Excel.Range).Value, 1)
                                .Item(72, rwx).Value = Strings.Left(CType(range.Cells(rCnt, 73), Excel.Range).Value, 1)
                                .Item(73, rwx).Value = Strings.Left(CType(range.Cells(rCnt, 74), Excel.Range).Value, 1)
                                .Item(74, rwx).Value = Strings.Left(CType(range.Cells(rCnt, 75), Excel.Range).Value, 1)
                                .Item(75, rwx).Value = Strings.Left(CType(range.Cells(rCnt, 76), Excel.Range).Value, 1)
                                .Item(76, rwx).Value = Strings.Left(CType(range.Cells(rCnt, 77), Excel.Range).Value, 1)
                                'ps10
                            End With
                        End If

                    End If

                    totres += 1
                    pbar1.Value += 1
                    pbar1.Refresh()

                Next
                lbltotres.Text = totres
                MsgBox("Importing template data complete!", 64, "Imported Successfully")
            Else
                MsgBox("Invalid Template Format/ Insufficient Rows!", 48, "Please check format")
            End If


            xlWorkBook.Close()
            xlApp.Quit()

            releaseObject(xlApp)
            releaseObject(xlWorkBook)
            releaseObject(xlWorkSheet)
        End If
    End Sub
End Class
