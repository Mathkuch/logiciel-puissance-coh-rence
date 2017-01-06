Imports Microsoft.VisualBasic
Imports System
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Media
Imports System.Windows.Shapes
Imports System.Threading
Imports System.Windows.Threading
Imports Excel = Microsoft.Office.Interop.Excel
Imports System.ComponentModel

Class MainWindow
    Dim pause, echelle, r, Item As Int32
    Dim count As Int32
    Dim mySolidColorBrush1, mySolidColorBrush2, mySolidColorBrush3, mySolidColorBrush4 As New SolidColorBrush()
    'puissances
    Dim nO1Array, nT5Array, nC3Array, nF7Array, nFp1Array, nCzArray, nO2array, nT6Array, nC4Array, nF8Array, nFp2Array, Toppic As New List(Of Double)
    'Cohérence
    Dim Fp2F8, Fp2C4, Fp2T6, Fp2O2, Fp2Cz, Fp2Fp1, Fp2F7, Fp2C3, Fp2T5, Fp2O1, F8C4, F8T6, F8O2, F8Cz, F8Fp1, F8F7, F8C3, F8T5, F8O1 As New List(Of Double)
    Dim C4T6, C4O2, C4Cz, C4Fp1, C4F7, C4C3, C4T5, C4O1, T6O2, T6Cz, T6Fp1, T6F7, T6C3, T6T5, T6O1, O2Cz, O2Fp1, O2F7, O2C3, O2T5, O2O1, CzFp1, CzF7, CzC3, CzT5, CzO1 As New List(Of Double)
    Dim Fp1F7, Fp1C3, Fp1T5, Fp1O1, F7C3, F7T5, F7O1, C3T5, C3O1, T5O1 As New List(Of Double)
    Dim Coh As Double
    'list des canvas
    Dim canvasList, CanvasList1, CanvasList2, CanvasList3, CanvasList4, CanvaslistP, CanvaslistP1, CanvaslistP2, CanvaslistP3, CanvaslistP4, ListCanPic As New List(Of Canvas)
    'rond puissance
    Dim ellipseList As New List(Of Ellipse)
    'trait cohérence
    Dim LineList, lineList2, LineList3, lineList4, LineList1 As New List(Of Line)
    Dim Listline As New List(Of List(Of Line))
    Dim LineEEG As New Line
    Dim arrayList As New List(Of List(Of Double))
    Dim arraylist2 As New List(Of List(Of Double))
    'position des électrodes
    Dim centerList, centerList1, centerlist2, centerList3, centerList4 As New List(Of Point)
    Dim center, PointAlist, PointBlist As New List(Of List(Of Point))
    Dim nPoint, nPoint2, nPoint3 As New Point
    Dim dt As DispatcherTimer = New DispatcherTimer()
    Dim Picture, Picture1, Picture2, Picture3, Picture4, PictureEEG As New Image
    Dim PicList As New List(Of Image)
    Dim Canvaspicture, Canvaspicture1, Canvaspicture2, Canvaspicture3, Canvaspicture4, CanvaspictureEEG As New Canvas
    Dim ListCanvas As New List(Of List(Of Canvas))
    Dim Leftpic As Double
    Dim A, B, C, leftP, D As Integer
    Dim PathEEG, Infile As String
    Private _listCanPic As Object
    Dim YEEG As Double
    Dim CanvasligneEEG As New Canvas
    Private Sub Bouton1_Click(sender As Object, e As RoutedEventArgs) Handles Bouton1.Click
        dt.Interval = New TimeSpan(0, 0, 0, 0, 300)
        Select Case pause
            Case 0
                echelle = 2
                count = 0
                Dim nOFD As New Microsoft.Win32.OpenFileDialog()
                Textbox1.Text = "Bonjour"
                nOFD.DefaultExt = ".xlsm"
                nOFD.Filter = "Document Excel (*.xlsm)|*.xlsm"
                nOFD.Title = "Document Excel pour analyse?"
                Dim nResultOFD As Nullable(Of Boolean) = nOFD.ShowDialog()
                If nResultOFD = True Then
                    Textbox1.Text = nOFD.FileName
                    readExcelFile()
                End If
                AddHandler dt.Tick, AddressOf dispatcherTimer_Tick
                Scroll1.Maximum = nFp1Array.Count - 1
                For i As Integer = 0 To 10
                    Dim nCanvas As New Canvas()
                    Dim nEllipse As New Ellipse()
                    CanvaslistP.Add(nCanvas)
                    ellipseList.Add(nEllipse)
                    CanvaslistP(i).Children.Add(nEllipse)
                Next
                For j = 0 To 54
                    Dim cCanvas As New Canvas()
                    Dim Line As New Line()
                    canvasList.Add(cCanvas)
                    CanvasList1.Add(cCanvas)
                    CanvasList2.Add(cCanvas)
                    Canvaslist3.Add(cCanvas)
                    CanvasList4.Add(cCanvas)
                    LineList.Add(Line)
                    LineList1.Add(Line)
                    lineList2.Add(Line)
                    LineList3.Add(Line)
                    lineList4.Add(Line)
                Next
                Listline.Add(LineList)
                Listline.Add(LineList1)
                Listline.Add(lineList2)
                Listline.Add(LineList3)
                Listline.Add(lineList4)
                ListCanvas.Add(canvasList)
                ListCanvas.Add(CanvasList1)
                ListCanvas.Add(CanvasList2)
                ListCanvas.Add(CanvasList3)
                ListCanvas.Add(CanvasList4)

                Cbbx1.Items.Add("Tracé précritique")
                Cbbx1.Items.Add("Départ en frontal gauche")
                Cbbx1.Items.Add("Bifrontale")
                Cbbx1.Items.Add("Bifrontale prédominant à G+ Temporal G")
                Cbbx1.Items.Add("Bifrontale plutôt gauche")
                Cbbx1.Items.Add("Explosition occipitotemporal gauche")
                Cbbx1.Items.Add("Hémipshérique gauche frontocentral et temporo-occipital gauche")
                Cbbx1.Items.Add("Départ en fronto-temporal droit avec crise toujours à gauche")
                Cbbx1.Items.Add("3 foyers: fronto-temporal droit, central droit et temporo-occipital gauche")
                Cbbx1.Items.Add("Crise exclusivement droite")
                Cbbx1.Items.Add("Tracé postcritique")
                dt.Start()
                pause = 1
                Bouton1.Content = "Pause"
            Case 1
                dt.Stop()
                pause = 2
                Bouton1.Content = "Play"
            Case 2
                dt.Start()
                pause = 1
                Bouton1.Content = "Pause"
        End Select
    End Sub
    Private Sub coordelectrodes()
        For i = 0 To 10
            centerList.Add(New Point)
            centerList1.Add(New Point)
            centerlist2.Add(New Point)
            centerList3.Add(New Point)
            centerList4.Add(New Point)
        Next
        center.Add(centerList)
        center.Add(centerList1)
        center.Add(centerlist2)
        center.Add(centerList3)
        center.Add(centerList4)
        For j As Integer = 0 To 4
            center(j)(0) = New Point(Leftpic + 0.57647 * PicList(0).ActualWidth, Toppic(j) + 0.169863 * PicList(0).ActualHeight)
            center(j)(1) = New Point(Leftpic + 0.724705882 * PicList(0).ActualWidth, Toppic(j) + 0.32055 * PicList(0).ActualHeight)
            center(j)(2) = New Point(Leftpic + 0.635294118 * PicList(0).ActualWidth, Toppic(j) + 0.52603 * PicList(0).ActualHeight)
            center(j)(3) = New Point(Leftpic + 0.729411765 * PicList(0).ActualWidth, Toppic(j) + 0.75342 * PicList(0).ActualHeight)
            center(j)(4) = New Point(Leftpic + 0.57647 * PicList(0).ActualWidth, Toppic(j) + 0.87397 * PicList(0).ActualHeight)
            center(j)(5) = New Point(Leftpic + 0.477647059 * PicList(0).ActualWidth, Toppic(j) + 0.52603 * PicList(0).ActualHeight)
            center(j)(6) = New Point(Leftpic + 0.376470588 * PicList(0).ActualWidth, Toppic(j) + 0.169863 * PicList(0).ActualHeight)
            center(j)(7) = New Point(Leftpic + 0.211764706 * PicList(0).ActualWidth, Toppic(j) + 0.32055 * PicList(0).ActualHeight)
            center(j)(8) = New Point(Leftpic + 0.322352941 * PicList(0).ActualWidth, Toppic(j) + 0.52603 * PicList(0).ActualHeight)
            center(j)(9) = New Point(Leftpic + 0.230588235 * PicList(0).ActualWidth, Toppic(j) + 0.7452 * PicList(0).ActualHeight)
            center(j)(10) = New Point(Leftpic + 0.385882353 * PicList(0).ActualWidth, Toppic(j) + 0.87397 * PicList(0).ActualHeight)
        Next
    End Sub
    Public Sub dispatcherTimer_Tick(ByVal sender As Object, ByVal e As EventArgs)
        If count < nFp1Array.Count - 1 Then
            count = count + 1
            coordelectrodes()
            Puissance()
            coherence()
            EEG()
            Scroll1.Value = count
        Else
            dt.Stop()
        End If
    End Sub
    Private Sub readExcelFile()
        Dim nApp As Excel.Application
        Dim nWorkbook As Excel.Workbook
        Dim nWorksheet As Excel.Worksheet
        nApp = New Excel.Application
        nWorkbook = nApp.Workbooks.Open(Textbox1.Text)
        nWorksheet = nWorkbook.Worksheets("P T")
        Dim nRange As Excel.Range = nWorksheet.UsedRange
        Dim nArray(,) As Object = nRange.Value(Excel.XlRangeValueDataType.xlRangeValueDefault)
        Dim nSize As Integer = nArray.GetUpperBound(0)
        For i As Integer = 2 To nSize
            nO1Array.Add(nArray(i, 1))
            nT5Array.Add(nArray(i, 2))
            nC3Array.Add(nArray(i, 3))
            nF7Array.Add(nArray(i, 4))
            nFp1Array.Add(nArray(i, 5))
            nCzArray.Add(nArray(i, 6))
            nO2array.Add(nArray(i, 7))
            nT6Array.Add(nArray(i, 8))
            nC4Array.Add(nArray(i, 9))
            nF8Array.Add(nArray(i, 10))
            nFp2Array.Add(nArray(i, 11))
        Next
        arrayList.Add(nFp2Array)
        arrayList.Add(nF8Array)
        arrayList.Add(nC4Array)
        arrayList.Add(nT6Array)
        arrayList.Add(nO2array)
        arrayList.Add(nCzArray)
        arrayList.Add(nFp1Array)
        arrayList.Add(nC3Array)
        arrayList.Add(nF7Array)
        arrayList.Add(nT5Array)
        arrayList.Add(nO1Array)
        nWorksheet = nWorkbook.Worksheets("C T")
        Dim nRange2 As Excel.Range = nWorksheet.UsedRange
        Dim nArray2(,) As Object = nRange2.Value(Excel.XlRangeValueDataType.xlRangeValueDefault)
        Dim nSize2 As Integer = nArray2.GetUpperBound(0)
        For i As Integer = 2 To nSize2
            Fp2F8.Add(nArray2(i, 1))
            Fp2C4.Add(nArray2(i, 2))
            Fp2T6.Add(nArray2(i, 3))
            Fp2O2.Add(nArray2(i, 4))
            Fp2Cz.Add(nArray2(i, 5))
            Fp2Fp1.Add(nArray2(i, 6))
            Fp2F7.Add(nArray2(i, 7))
            Fp2C3.Add(nArray2(i, 8))
            Fp2T5.Add(nArray2(i, 9))
            Fp2O1.Add(nArray2(i, 10))
            F8C4.Add(nArray2(i, 11))
            F8T6.Add(nArray2(i, 12))
            F8O2.Add(nArray2(i, 13))
            F8Cz.Add(nArray2(i, 14))
            F8Fp1.Add(nArray2(i, 15))
            F8F7.Add(nArray2(i, 16))
            F8C3.Add(nArray2(i, 17))
            F8T5.Add(nArray2(i, 18))
            F8O1.Add(nArray2(i, 19))
            C4T6.Add(nArray2(i, 20))
            C4O2.Add(nArray2(i, 21))
            C4Cz.Add(nArray2(i, 22))
            C4Fp1.Add(nArray2(i, 23))
            C4F7.Add(nArray2(i, 24))
            C4C3.Add(nArray2(i, 25))
            C4T5.Add(nArray2(i, 26))
            C4O1.Add(nArray2(i, 27))
            T6O2.Add(nArray2(i, 28))
            T6Cz.Add(nArray2(i, 29))
            T6Fp1.Add(nArray2(i, 30))
            T6F7.Add(nArray2(i, 31))
            T6C3.Add(nArray2(i, 32))
            T6T5.Add(nArray2(i, 33))
            T6O1.Add(nArray2(i, 34))
            O2Cz.Add(nArray2(i, 35))
            O2Fp1.Add(nArray2(i, 36))
            O2F7.Add(nArray2(i, 37))
            O2C3.Add(nArray2(i, 38))
            O2T5.Add(nArray2(i, 39))
            O2O1.Add(nArray2(i, 40))
            CzFp1.Add(nArray2(i, 41))
            CzF7.Add(nArray2(i, 42))
            CzC3.Add(nArray2(i, 43))
            CzT5.Add(nArray2(i, 44))
            CzO1.Add(nArray2(i, 45))
            Fp1F7.Add(nArray2(i, 46))
            Fp1C3.Add(nArray2(i, 47))
            Fp1T5.Add(nArray2(i, 48))
            Fp1O1.Add(nArray2(i, 49))
            F7C3.Add(nArray2(i, 50))
            F7T5.Add(nArray2(i, 51))
            F7O1.Add(nArray2(i, 52))
            C3T5.Add(nArray2(i, 53))
            C3O1.Add(nArray2(i, 54))
            T5O1.Add(nArray2(i, 55))
        Next
        arraylist2.Add(Fp2F8)
        arraylist2.Add(Fp2C4)
        arraylist2.Add(Fp2T6)
        arraylist2.Add(Fp2O2)
        arraylist2.Add(Fp2Cz)
        arraylist2.Add(Fp2Fp1)
        arraylist2.Add(Fp2F7)
        arraylist2.Add(Fp2C3)
        arraylist2.Add(Fp2T5)
        arraylist2.Add(Fp2O1)

        arraylist2.Add(F8C4)
        arraylist2.Add(F8T6)
        arraylist2.Add(F8O2)
        arraylist2.Add(F8Cz)
        arraylist2.Add(F8Fp1)
        arraylist2.Add(F8F7)
        arraylist2.Add(F8C3)
        arraylist2.Add(F8T5)
        arraylist2.Add(F8O1)

        arraylist2.Add(C4T6)
        arraylist2.Add(C4O2)
        arraylist2.Add(C4Cz)
        arraylist2.Add(C4Fp1)
        arraylist2.Add(C4F7)
        arraylist2.Add(C4C3)
        arraylist2.Add(C4T5)
        arraylist2.Add(C4O1)

        arraylist2.Add(T6O2)
        arraylist2.Add(T6Cz)
        arraylist2.Add(T6Fp1)
        arraylist2.Add(T6F7)
        arraylist2.Add(T6C3)
        arraylist2.Add(T6T5)
        arraylist2.Add(T6O1)

        arraylist2.Add(O2Cz)
        arraylist2.Add(O2Fp1)
        arraylist2.Add(O2F7)
        arraylist2.Add(O2C3)
        arraylist2.Add(O2T5)
        arraylist2.Add(O2O1)

        arraylist2.Add(CzFp1)
        arraylist2.Add(CzF7)
        arraylist2.Add(CzC3)
        arraylist2.Add(CzT5)
        arraylist2.Add(CzO1)

        arraylist2.Add(Fp1F7)
        arraylist2.Add(Fp1C3)
        arraylist2.Add(Fp1T5)
        arraylist2.Add(Fp1O1)

        arraylist2.Add(F7C3)
        arraylist2.Add(F7T5)
        arraylist2.Add(F7O1)

        arraylist2.Add(C3T5)
        arraylist2.Add(C3O1)

        arraylist2.Add(T5O1)
        MsgBox("Chargement fait")
    End Sub
    Private Sub Coordonnee(i As Integer)
        Dim E As Integer = 0
        If i < 10 Then
            LineList(i).X1 = center(E)(0).X
            LineList(i).Y1 = center(E)(0).Y
            Select Case i
                Case 0
                    LineList(i).X2 = center(E)(2).X
                    LineList(i).Y2 = center(E)(5).Y
                Case 1
                    LineList(i).X2 = center(E)(2).X
                    LineList(i).Y2 = center(E)(5).Y
                Case 2
                    LineList(i).X2 = center(E)(3).X
                    LineList(i).Y2 = center(E)(5).Y
                Case 3
                    LineList(i).X2 = center(E)(4).X
                    LineList(i).Y2 = center(E)(5).Y
                Case 4
                    LineList(i).X2 = center(E)(5).X
                    LineList(i).Y2 = center(E)(5).Y
                Case 5
                    LineList(i).X2 = center(E)(6).X
                    LineList(i).Y2 = center(E)(6).Y
                Case 6
                    LineList(i).X2 = center(E)(7).X
                    LineList(i).Y2 = center(E)(7).Y
                Case 7
                    LineList(i).X2 = center(E)(8).X
                    LineList(i).Y2 = center(E)(8).Y
                Case 8
                    LineList(i).X2 = center(E)(9).X
                    LineList(i).Y2 = center(E)(9).Y
                Case 9
                    LineList(i).X2 = center(E)(10).X
                    LineList(i).Y2 = center(E)(10).Y
            End Select
        ElseIf i < 19 And i > 9 Then
            LineList(i).X1 = center(E)(1).X
            LineList(i).Y1 = center(E)(1).Y
            Select Case i
                Case 10
                    LineList(i).X2 = center(E)(2).X
                    LineList(i).Y2 = center(E)(5).Y
                Case 11
                    LineList(i).X2 = center(E)(3).X
                    LineList(i).Y2 = center(E)(5).Y
                Case 12
                    LineList(i).X2 = center(E)(4).X
                    LineList(i).Y2 = center(E)(5).Y
                Case 13
                    LineList(i).X2 = center(E)(5).X
                    LineList(i).Y2 = center(E)(5).Y
                Case 14
                    LineList(i).X2 = center(E)(6).X
                    LineList(i).Y2 = center(E)(6).Y
                Case 15
                    LineList(i).X2 = center(E)(7).X
                    LineList(i).Y2 = center(E)(7).Y
                Case 16
                    LineList(i).X2 = center(E)(8).X
                    LineList(i).Y2 = center(E)(8).Y
                Case 17
                    LineList(i).X2 = center(E)(9).X
                    LineList(i).Y2 = center(E)(9).Y
                Case 18
                    LineList(i).X2 = center(E)(10).X
                    LineList(i).Y2 = center(E)(10).Y
            End Select
        ElseIf i < 27 And i > 18 Then
            LineList(i).X1 = center(E)(2).X
            LineList(i).Y1 = center(E)(2).Y
            Select Case i
                Case 19
                    LineList(i).X2 = center(E)(3).X
                    LineList(i).Y2 = center(E)(5).Y
                Case 20
                    LineList(i).X2 = center(E)(4).X
                    LineList(i).Y2 = center(E)(5).Y
                Case 21
                    LineList(i).X2 = center(E)(5).X
                    LineList(i).Y2 = center(E)(5).Y
                Case 22
                    LineList(i).X2 = center(E)(6).X
                    LineList(i).Y2 = center(E)(6).Y
                Case 23
                    LineList(i).X2 = center(E)(7).X
                    LineList(i).Y2 = center(E)(7).Y
                Case 24
                    LineList(i).X2 = center(E)(8).X
                    LineList(i).Y2 = center(E)(8).Y
                Case 25
                    LineList(i).X2 = center(E)(9).X
                    LineList(i).Y2 = center(E)(9).Y
                Case 26
                    LineList(i).X2 = center(E)(10).X
                    LineList(i).Y2 = center(E)(10).Y
            End Select
        ElseIf i < 34 And i > 26 Then
            LineList(i).X1 = center(E)(3).X
            LineList(i).Y1 = center(E)(3).Y
            Select Case i
                Case 27
                    LineList(i).X2 = center(E)(4).X
                    LineList(i).Y2 = center(E)(5).Y
                Case 28
                    LineList(i).X2 = center(E)(5).X
                    LineList(i).Y2 = center(E)(5).Y
                Case 29
                    LineList(i).X2 = center(E)(6).X
                    LineList(i).Y2 = center(E)(6).Y
                Case 30
                    LineList(i).X2 = center(E)(7).X
                    LineList(i).Y2 = center(E)(7).Y
                Case 31
                    LineList(i).X2 = center(E)(8).X
                    LineList(i).Y2 = center(E)(8).Y
                Case 32
                    LineList(i).X2 = center(E)(9).X
                    LineList(i).Y2 = center(E)(9).Y
                Case 33
                    LineList(i).X2 = center(E)(10).X
                    LineList(i).Y2 = center(E)(10).Y
            End Select
        ElseIf i < 40 And i > 33 Then
            LineList(i).X1 = center(E)(4).X
            LineList(i).Y1 = center(E)(4).Y
            Select Case i
                Case 34
                    LineList(i).X2 = center(E)(5).X
                    LineList(i).Y2 = center(E)(5).Y
                Case 35
                    LineList(i).X2 = center(E)(6).X
                    LineList(i).Y2 = center(E)(6).Y
                Case 36
                    LineList(i).X2 = center(E)(7).X
                    LineList(i).Y2 = center(E)(7).Y
                Case 37
                    LineList(i).X2 = center(E)(8).X
                    LineList(i).Y2 = center(E)(8).Y
                Case 38
                    LineList(i).X2 = center(E)(9).X
                    LineList(i).Y2 = center(E)(9).Y
                Case 39
                    LineList(i).X2 = center(E)(10).X
                    LineList(i).Y2 = center(E)(10).Y
            End Select
        ElseIf i < 45 And i > 39 Then
            LineList(i).X1 = center(E)(5).X
            LineList(i).Y1 = center(E)(5).Y
            Select Case i
                Case 40
                    LineList(i).X2 = center(E)(6).X
                    LineList(i).Y2 = center(E)(6).Y
                Case 41
                    LineList(i).X2 = center(E)(7).X
                    LineList(i).Y2 = center(E)(7).Y
                Case 42
                    LineList(i).X2 = center(E)(8).X
                    LineList(i).Y2 = center(E)(8).Y
                Case 43
                    LineList(i).X2 = center(E)(9).X
                    LineList(i).Y2 = center(E)(9).Y
                Case 44
                    LineList(i).X2 = center(E)(10).X
                    LineList(i).Y2 = center(E)(10).Y
            End Select
        ElseIf i < 49 And i > 44 Then
            LineList(i).X1 = center(E)(6).X
            LineList(i).Y1 = center(E)(6).Y
            Select Case i
                Case 45
                    LineList(i).X2 = center(E)(7).X
                    LineList(i).Y2 = center(E)(7).Y
                Case 46
                    LineList(i).X2 = center(E)(8).X
                    LineList(i).Y2 = center(E)(8).Y
                Case 47
                    LineList(i).X2 = center(E)(9).X
                    LineList(i).Y2 = center(E)(9).Y
                Case 48
                    LineList(i).X2 = center(E)(10).X
                    LineList(i).Y2 = center(E)(10).Y
            End Select
        ElseIf i < 52 And i > 48 Then
            LineList(i).X1 = center(E)(7).X
            LineList(i).Y1 = center(E)(7).Y
            Select Case i
                Case 49
                    LineList(i).X2 = center(E)(8).X
                    LineList(i).Y2 = center(E)(8).Y
                Case 50
                    LineList(i).X2 = center(E)(9).X
                    LineList(i).Y2 = center(E)(9).Y
                Case 51
                    LineList(i).X2 = center(E)(10).X
                    LineList(i).Y2 = center(E)(10).Y
            End Select
        ElseIf i < 54 And i > 51 Then
            LineList(i).X1 = center(E)(8).X
            LineList(i).Y1 = center(E)(8).Y
            Select Case i
                Case 52
                    LineList(i).X2 = center(E)(9).X
                    LineList(i).Y2 = center(E)(9).Y
                Case 53
                    LineList(i).X2 = center(E)(10).X
                    LineList(i).Y2 = center(E)(10).Y
            End Select
        Else
            LineList(i).X1 = center(E)(9).X
            LineList(i).Y1 = center(E)(9).Y
            LineList(i).X2 = center(E)(10).X
            LineList(i).Y2 = center(E)(10).Y
        End If
    End Sub
    Private Sub coherence()
        mySolidColorBrush1.Color = Color.FromRgb(0, 0, 153)
        For i As Integer = 0 To 54
            Coh = arraylist2(i)(count)
            If Coh > 0.2 Then
                Coordonnee(i)
                Coordonnee(i)
                If Coh < 0.2 Then
                    LineList(i).Stroke = mySolidColorBrush1
                ElseIf Coh < 0.3 And Coh > 0.199999 Then
                    LineList(i).Stroke = mySolidColorBrush2
                ElseIf Coh < 0.4 And Coh < 0.29999 Then
                    LineList(i).Stroke = mySolidColorBrush3
                Else
                    LineList(i).Stroke = mySolidColorBrush4
                End If
                canvasList(i).Children.Remove(LineList(i))
                Canvas1.Children.Remove(canvasList(i))
                canvasList(i).Children.Add(LineList(i))
                Canvas1.Children.Add(canvasList(i))

                'ListCanvas(E)(i).Children.Remove(linelist(i))
                'Canvas1.Children.Remove(ListCanvas(E)(i))
                'ListCanvas(E)(i).Children.Add(linelist(i))
                'Canvas1.Children.Add(ListCanvas(E)(i))
            End If
            LineList(i).StrokeThickness = 5
        Next
        'MsgBox("changement de scalp")
        Me.Content = Canvas1
        Scroll1.Value = count
    End Sub
    Private Sub Puissance()
        mySolidColorBrush1.Color = Color.FromRgb(0, 0, 153)
        mySolidColorBrush2.Color = Color.FromRgb(51, 204, 0)
        mySolidColorBrush3.Color = Color.FromRgb(255, 204, 51)
        mySolidColorBrush4.Color = Color.FromRgb(255, 0, 0)
        For i As Integer = 0 To 10
            Item = arrayList(i)(count)
            If Item < 5 Then
                r = 5 * echelle
                ellipseList(i).Fill = mySolidColorBrush1
            ElseIf Item > 4 And Item < 11 Then
                r = Item * echelle
                ellipseList(i).Fill = mySolidColorBrush2
            ElseIf Item > 10 And Item < 16 Then
                r = Item * echelle
                ellipseList(i).Fill = mySolidColorBrush3
            Else
                r = Item * echelle
                ellipseList(i).Fill = mySolidColorBrush4
            End If
            ellipseList(i).StrokeThickness = 5
            ellipseList(i).Width = r / B
            ellipseList(i).Height = r / B
            CanvaslistP(i).Height = 100
            CanvaslistP(i).Width = 100
            nPoint = center(0)(i)
            Canvas.SetTop(CanvaslistP(i), nPoint.X - (r / (2 * B)))
            Canvas.SetLeft(CanvaslistP(i), nPoint.Y - (r / (2 * B)))
            CanvaslistP(i).Children.Remove(ellipseList(i))
            Canvas1.Children.Remove(CanvaslistP(i))
            CanvaslistP(i).Children.Add(ellipseList(i))
            Canvas1.Children.Add(CanvaslistP(i))
        Next
        Textbox2.Text = count
        If count < 105 Then
            Cbbx1.Text = "Tracé précritique"
        ElseIf count < 120 And count > 104 Then
            Cbbx1.Text = "Départ en frontal gauche"
        ElseIf count < 135 And count > 119 Then
            Cbbx1.Text = "Bifrontale prédominant à G+ Temporal G"
        ElseIf count < 155 And count > 134 Then
            Cbbx1.Text = "Bifrontale"
        ElseIf count < 180 And count > 154 Then
            Cbbx1.Text = "Bifrontale prédominant à G+ Temporal G"
        ElseIf count < 320 And count > 179 Then
            Cbbx1.Text = "Explosition occipitotemporal gauche"
        ElseIf count < 398 And count > 319 Then
            Cbbx1.Text = "Hémipshérique gauche frontocentral et temporo-occipital gauche"
        ElseIf count < 470 And count > 397 Then
            Cbbx1.Text = "Départ en fronto-temporal droit avec crise toujours à gauche"
        ElseIf count < 555 And count > 469 Then
            Cbbx1.Text = "3 foyers: fronto-temporal droit, central droit et temporo-occipital gauche"
        ElseIf count < 850 And count > 554 Then
            Cbbx1.Text = "Crise exclusivement droite"
        Else
            Cbbx1.Text = "Tracé postcritique"
        End If
        Me.Content = Canvas1
    End Sub
    Private Sub EEG()
        LineEEG.Y1 = 150
        LineEEG.Y2 = PictureEEG.ActualHeight + 150
        If C = 1 Then
            PathEEG = "C:\Users\Kuchenbuch\Documents\Visual Studio 2013\Projects\MPSI\MPSI\Resources\EEG1\"
        Else
            PathEEG = "C:\Users\Kuchenbuch\Documents\Visual Studio 2013\Projects\MPSI\MPSI\Resources\EEG2\9"
        End If
        If count = 1 Then
            PictureEEG.Source = New BitmapImage(New Uri("C:\Users\Kuchenbuch\Documents\Visual Studio 2013\Projects\MPSI\MPSI\Resources\EEG2\90.jpg"))
            YEEG = 150 + 64
            A = 0
            LineEEG.Stroke = mySolidColorBrush1
            LineEEG.StrokeThickness = 5
        ElseIf (count - 1) / 30 = Int((count - 1) / 30) Then
            CanvasligneEEG.Children.Remove(LineEEG)
            Canvas1.Children.Remove(CanvasligneEEG)
            YEEG = 150 + 64 + ((PictureEEG.ActualWidth - 64) / 30)
            A = 0
            CanvasligneEEG.Children.Add(LineEEG)
            Canvas1.Children.Add(CanvasligneEEG)
        Else
            CanvasligneEEG.Children.Remove(LineEEG)
            Canvas1.Children.Remove(CanvasligneEEG)
            A = count - Int((count - 1) / 30) * 30
            YEEG = 150 + (64) + A * ((PictureEEG.ActualWidth - 64) / 30)
            CanvasligneEEG.Children.Add(LineEEG)
            Canvas1.Children.Add(CanvasligneEEG)
        End If
        PictureEEG.Source = New BitmapImage(New Uri(PathEEG & Int((count - 1) / 30) & ".jpg"))
        LineEEG.X1 = YEEG
        LineEEG.X2 = YEEG
        LineEEG.Stroke = mySolidColorBrush1
        LineEEG.StrokeThickness = 5
    End Sub
    Private Sub Scroll1_Scroll(sender As Object, e As Primitives.ScrollEventArgs) Handles Scroll1.Scroll
        If pause = 1 Then
            dt.Stop()
            pause = 2
            Bouton1.Content = "Play"
        End If
        count = Scroll1.Value
        Puissance()
        coherence()
        EEG()
    End Sub
    Private Sub Textbox2_KeyUp(sender As Object, e As KeyEventArgs) Handles Textbox2.KeyUp
        On Error Resume Next
        count = CType(Textbox2.Text, Integer)
        Puissance()
        coherence()
        Scroll1.Value = count
    End Sub
    Private Sub Textbox2_MouseEnter(sender As Object, e As MouseEventArgs) Handles Textbox2.MouseEnter
        If pause = 1 Then
            dt.Stop()
            pause = 2
            Bouton1.Content = "Play"
        End If
        Textbox2.Text = ""
    End Sub
    Private Sub Cbbx1_GotFocus(sender As Object, e As RoutedEventArgs) Handles Cbbx1.GotFocus
        dt.Stop()
        pause = 2
        Bouton1.Content = "Play"
    End Sub
    Private Sub Cbbx1_LostFocus(sender As Object, e As RoutedEventArgs) Handles Cbbx1.LostFocus
        Select Case Cbbx1.SelectedIndex
            Case 0
                count = 2
            Case 1
                count = 105
            Case 2
                count = 120
            Case 3
                count = 135
            Case 4
                count = 155
            Case 5
                count = 180
            Case 6
                count = 320
            Case 7
                count = 398
            Case 8
                count = 470
            Case 9
                count = 555
            Case 10
                count = 850
        End Select
        Puissance()
        coherence()
        EEG()
        Scroll1.Value = count
        Textbox2.Text = count
    End Sub
    Private Sub windows1_Loaded(sender As Object, e As RoutedEventArgs) Handles windows1.Loaded
        pause = 0
        For i As Integer = 0 To 4
            PicList(i).Source = New BitmapImage(New Uri("C:\Users\Kuchenbuch\Documents\Visual Studio 2013\Projects\MPSI\MPSI\Resources\scalp2.jpg"))
        Next
        B = 2
        Me.Content = Canvas1
        C = 0
    End Sub
    Public Sub windows1_SizeChanged(sender As Object, e As SizeChangedEventArgs) Handles windows1.SizeChanged
        PicList.Clear()
        ListCanPic.Clear()
        PicList.Add(Picture)
        PicList.Add(Picture1)
        PicList.Add(Picture2)
        PicList.Add(Picture3)
        PicList.Add(Picture4)
        ListCanPic.Add(Canvaspicture)
        ListCanPic.Add(Canvaspicture1)
        ListCanPic.Add(Canvaspicture2)
        ListCanPic.Add(Canvaspicture3)
        ListCanPic.Add(Canvaspicture4)
        modif()
        CanvaspictureEEG.Children.Remove(PictureEEG)
        Canvas1.Children.Remove(CanvaspictureEEG)
        If C = 1 Then
            PathEEG = "C:\Users\Kuchenbuch\Documents\Visual Studio 2013\Projects\MPSI\MPSI\Resources\EEG1\"
            LineEEG.Y2 = 750 + 150
            Scroll1.Margin = New Thickness(150, 750 + 150, Scroll1.Margin.Right, Scroll1.Margin.Bottom)
            Scroll1.Width = 1278
        Else
            PathEEG = "C:\Users\Kuchenbuch\Documents\Visual Studio 2013\Projects\MPSI\MPSI\Resources\EEG2\9"
            LineEEG.Y2 = 450 + 150
            Scroll1.Margin = New Thickness(150, 450 + 150, Scroll1.Margin.Right, Scroll1.Margin.Bottom)
            Scroll1.Width = 766
        End If
        Canvas.SetLeft(CanvaspictureEEG, 150)
        Canvas.SetTop(CanvaspictureEEG, 150)
        If count = 0 Then
            count = 1
        End If
        PictureEEG.Source = New BitmapImage(New Uri(PathEEG & Int((count - 1) / 30) & ".jpg"))
        CanvaspictureEEG.Children.Add(PictureEEG)
        Canvas1.Children.Add(CanvaspictureEEG)
        Me.Content = Canvas1
    End Sub
    Private Sub modif()
        For i As Integer = 0 To 4
            ListCanPic(i).Children.Remove(PicList(i))
            Canvas1.Children.Remove(ListCanPic(i))
            If windows1.ActualWidth > 1500 Then
                PicList(i).Source = New BitmapImage(New Uri("C:\Users\Kuchenbuch\Documents\Visual Studio 2013\Projects\MPSI\MPSI\Resources\Scalp.jpg"))
                leftP = 280
                C = 1
                B = 1
            Else
                PicList(i).Source = New BitmapImage(New Uri("C:\Users\Kuchenbuch\Documents\Visual Studio 2013\Projects\MPSI\MPSI\Resources\Scalp2.jpg"))
                leftP = 120
                C = 0
                B = 2
            End If
            Leftpic = windows1.ActualWidth - 280
            Toppic.Add(windows1.ActualHeight - (windows1.ActualHeight - 30) + 200 * i)
            Canvas.SetTop(ListCanPic(i), Toppic(i))
            Canvas.SetLeft(ListCanPic(i), Leftpic)
            ListCanPic(i).Children.Add(PicList(i))
            Canvas1.Children.Add(ListCanPic(i))
        Next
        CanvasligneEEG.Children.Remove(LineEEG)
        Canvas1.Children.Remove(CanvasligneEEG)
        CanvasligneEEG.Children.Add(LineEEG)
        Canvas1.Children.Add(CanvasligneEEG)
    End Sub

End Class
