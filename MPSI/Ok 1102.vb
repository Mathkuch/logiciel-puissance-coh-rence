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
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.Windows.Data
Imports System.Windows.Documents
Imports System.Windows.Input
Imports System.Windows.Media.Imaging
Imports System.Windows.Navigation

Public Class MaskedTextBox
    Inherits Xceed.Wpf.Toolkit.Primitives.ValueRangeTextBox
End Class
Class MainWindow
    Dim MTXTbox As MaskedTextBox
    Dim pause, echelle, r, Item, Slide As Int32
    Dim count As Int32
    Dim mySolidColorBrush1, mySolidColorBrush2, mySolidColorBrush3, mySolidColorBrush4 As New SolidColorBrush()
    Dim dt As DispatcherTimer = New DispatcherTimer()

    'EEG
    Dim LineEEG As New Line
    '''images des EEG
    Dim PictureEEG As New Image
    Dim TopPic As New List(Of Double)
    '''list des canvas pour EEG 
    Dim canvasligneEEG As New Canvas
    Dim CanvaspictureEEG As New Canvas
    '''position barre sur EEG
    Dim YEEG As Integer
    'Scalp
    '''image du scalp
    Dim Picture, Picture1, Picture2, Picture3, Picture4 As New Image
    Dim PicList As New List(Of Image)
    '''List canvas picture Scalp
    Dim Canvaspicture, Canvaspicture1, Canvaspicture2, Canvaspicture3, Canvaspicture4 As New Canvas
    Dim listcanpic As New List(Of Canvas)
    '''position des électrodes sur les scalps
    Dim centerList, centerList1, centerlist2, centerList3, centerList4 As New List(Of Point)
    Dim center As New List(Of List(Of Point))
    Dim TxtboxD, TxtBoxT, TxtBoxA, TxtBoxB, TxtBoxG As New TextBlock
    Dim Txtlist As New List(Of TextBlock)
    Dim leftTxt As Integer


    'Coherence
    '''données Excel
    Dim Fp2F8, Fp2C4, Fp2T6, Fp2O2, Fp2Cz, Fp2Fp1, Fp2F7, Fp2C3, Fp2T5, Fp2O1, F8C4, F8T6, F8O2, F8Cz, F8Fp1, F8F7, F8C3, F8T5, F8O1 As New List(Of Double)
    Dim C4T6, C4O2, C4Cz, C4Fp1, C4F7, C4C3, C4T5, C4O1, T6O2, T6Cz, T6Fp1, T6F7, T6C3, T6T5, T6O1, O2Cz, O2Fp1, O2F7, O2C3, O2T5, O2O1, CzFp1, CzF7, CzC3, CzT5, CzO1 As New List(Of Double)
    Dim Fp1F7, Fp1C3, Fp1T5, Fp1O1, F7C3, F7T5, F7O1, C3T5, C3O1, T5O1 As New List(Of Double)

    Dim Fp2F81, Fp2C41, Fp2T61, Fp2O21, Fp2Cz1, Fp2Fp11, Fp2F71, Fp2C31, Fp2T51, Fp2O11, F8C41, F8T61, F8O21, F8Cz1, F8Fp11, F8F71, F8C31, F8T51, F8O11 As New List(Of Double)
    Dim C4T61, C4O21, C4Cz1, C4Fp11, C4F71, C4C31, C4T51, C4O11, T6O21, T6Cz1, T6Fp11, T6F71, T6C31, T6T51, T6O11, O2Cz1, O2Fp11, O2F71, O2C31, O2T51, O2O11, CzFp11, CzF71, CzC31, CzT51, CzO11 As New List(Of Double)
    Dim Fp1F71, Fp1C31, Fp1T51, Fp1O11, F7C31, F7T51, F7O11, C3T51, C3O11, T5O11 As New List(Of Double)

    Dim Fp2F82, Fp2C42, Fp2T62, Fp2O22, Fp2Cz2, Fp2Fp12, Fp2F72, Fp2C32, Fp2T52, Fp2O12, F8C42, F8T62, F8O22, F8Cz2, F8Fp12, F8F72, F8C32, F8T52, F8O12 As New List(Of Double)
    Dim C4T62, C4O22, C4Cz2, C4Fp12, C4F72, C4C32, C4T52, C4O12, T6O22, T6Cz2, T6Fp12, T6F72, T6C32, T6T52, T6O12, O2Cz2, O2Fp12, O2F72, O2C32, O2T52, O2O12, CzFp12, CzF72, CzC32, CzT52, CzO12 As New List(Of Double)
    Dim Fp1F72, Fp1C32, Fp1T52, Fp1O12, F7C32, F7T52, F7O12, C3T52, C3O12, T5O12 As New List(Of Double)

    Dim Fp2F83, Fp2C43, Fp2T63, Fp2O23, Fp2Cz3, Fp2Fp13, Fp2F73, Fp2C33, Fp2T53, Fp2O13, F8C43, F8T63, F8O23, F8Cz3, F8Fp13, F8F73, F8C33, F8T53, F8O13 As New List(Of Double)
    Dim C4T63, C4O23, C4Cz3, C4Fp13, C4F73, C4C33, C4T53, C4O13, T6O23, T6Cz3, T6Fp13, T6F73, T6C33, T6T53, T6O13, O2Cz3, O2Fp13, O2F73, O2C33, O2T53, O2O13, CzFp13, CzF73, CzC33, CzT53, CzO13 As New List(Of Double)
    Dim Fp1F73, Fp1C33, Fp1T53, Fp1O13, F7C33, F7T53, F7O13, C3T53, C3O13, T5O13 As New List(Of Double)

    Dim Fp2F84, Fp2C44, Fp2T64, Fp2O24, Fp2Cz4, Fp2Fp14, Fp2F74, Fp2C34, Fp2T54, Fp2O14, F8C44, F8T64, F8O24, F8Cz4, F8Fp14, F8F74, F8C34, F8T54, F8O14 As New List(Of Double)
    Dim C4T64, C4O24, C4Cz4, C4Fp14, C4F74, C4C34, C4T54, C4O14, T6O24, T6Cz4, T6Fp14, T6F74, T6C34, T6T54, T6O14, O2Cz4, O2Fp14, O2F74, O2C34, O2T54, O2O14, CzFp14, CzF74, CzC34, CzT54, CzO14 As New List(Of Double)
    Dim Fp1F74, Fp1C34, Fp1T54, Fp1O14, F7C34, F7T54, F7O14, C3T54, C3O14, T5O14 As New List(Of Double)
    '''valeur
    Dim CarrayList, CarrayList1, CarrayList2, CarrayList3, CarrayList4 As New List(Of List(Of Double))
    Dim CListofArray As New List(Of List(Of List(Of Double)))
    Dim Coh As Double
    Dim SeuilCoh As Double
    '''list des canvas cohérence
    Dim CanvasList, CanvasList1, CanvasList2, CanvasList3, CanvasList4 As New List(Of Canvas)
    Dim ListCanvasC As New List(Of List(Of Canvas))
    ''''trait cohérence
    Dim LineList, LineList1, LineList2, LineList3, LineList4 As New List(Of Line)
    Dim ListofLine As New List(Of List(Of Line))

    'Puissance
    '''données Excel
    Dim nO1Array, nT5Array, nC3Array, nF7Array, nFp1Array, nCzArray, nO2array, nT6Array, nC4Array, nF8Array, nFp2Array As New List(Of Double)
    Dim nO1Array1, nT5Array1, nC3Array1, nF7Array1, nFp1Array1, nCzArray1, nO2array1, nT6Array1, nC4Array1, nF8Array1, nFp2Array1 As New List(Of Double)
    Dim nO1Array2, nT5Array2, nC3Array2, nF7Array2, nFp1Array2, nCzArray2, nO2array2, nT6Array2, nC4Array2, nF8Array2, nFp2Array2 As New List(Of Double)
    Dim nO1Array3, nT5Array3, nC3Array3, nF7Array3, nFp1Array3, nCzArray3, nO2array3, nT6Array3, nC4Array3, nF8Array3, nFp2Array3 As New List(Of Double)
    Dim nO1Array4, nT5Array4, nC3Array4, nF7Array4, nFp1Array4, nCzArray4, nO2array4, nT6Array4, nC4Array4, nF8Array4, nFp2Array4 As New List(Of Double)
    Dim arrayList, arrayList1, arraylist2, arraylist3, arraylist4 As New List(Of List(Of Double))
    Dim ListofArray As New List(Of List(Of List(Of Double)))

    '''rond puissance
    Dim ellipseList, ellipseList1, ellipseList2, ellipseList3, ellipseList4 As New List(Of Ellipse)
    Dim ListEl As New List(Of List(Of Ellipse))
    '''liste des canvas puissance
    Dim canvasListP, canvasListP1, canvasListP2, canvasListP3, canvasListP4 As New List(Of Canvas)
    Dim ListCanvasP As New List(Of List(Of Canvas))



    Dim nPoint2, nPoint3 As New Point
    Dim Leftpic As Double
    Dim A, B, C, leftP As Integer
    Dim PathEEG, Infile As String
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
                    Dim nCanvas1 As New Canvas()
                    Dim nCanvas2 As New Canvas()
                    Dim nCanvas3 As New Canvas()
                    Dim nCanvas4 As New Canvas()
                    Dim nEllipse As New Ellipse()
                    Dim nEllipse1 As New Ellipse()
                    Dim nEllipse2 As New Ellipse()
                    Dim nEllipse3 As New Ellipse()
                    Dim nEllipse4 As New Ellipse()
                    canvasListP.Add(nCanvas)
                    canvasListP1.Add(nCanvas1)
                    canvasListP2.Add(nCanvas2)
                    canvasListP3.Add(nCanvas3)
                    canvasListP4.Add(nCanvas4)
                    ellipseList.Add(nEllipse)
                    ellipseList1.Add(nEllipse1)
                    ellipseList2.Add(nEllipse2)
                    ellipseList3.Add(nEllipse3)
                    ellipseList4.Add(nEllipse4)
                Next
                ListCanvasP.Add(canvasListP)
                ListCanvasP.Add(canvasListP1)
                ListCanvasP.Add(canvasListP2)
                ListCanvasP.Add(canvasListP3)
                ListCanvasP.Add(canvasListP4)
                ListEl.Add(ellipseList)
                ListEl.Add(ellipseList1)
                ListEl.Add(ellipseList2)
                ListEl.Add(ellipseList3)
                ListEl.Add(ellipseList4)
                For j = 0 To 54
                    Dim cCanvas As New Canvas()
                    Dim cCanvas1 As New Canvas()
                    Dim cCanvas2 As New Canvas()
                    Dim cCanvas3 As New Canvas()
                    Dim cCanvas4 As New Canvas()
                    Dim Line As New Line()
                    Dim Line1 As New Line()
                    Dim Line2 As New Line()
                    Dim Line3 As New Line()
                    Dim Line4 As New Line()
                    CanvasList.Add(cCanvas)
                    CanvasList1.Add(cCanvas1)
                    CanvasList2.Add(cCanvas2)
                    CanvasList3.Add(cCanvas3)
                    CanvasList4.Add(cCanvas4)
                    LineList.Add(Line)
                    LineList1.Add(Line1)
                    LineList2.Add(Line2)
                    LineList3.Add(Line3)
                    LineList4.Add(Line4)
                Next
                ListofLine.Add(LineList)
                ListofLine.Add(LineList1)
                ListofLine.Add(LineList2)
                ListofLine.Add(LineList3)
                ListofLine.Add(LineList4)
                ListCanvasC.Add(CanvasList)
                ListCanvasC.Add(CanvasList1)
                ListCanvasC.Add(CanvasList2)
                ListCanvasC.Add(CanvasList3)
                ListCanvasC.Add(CanvasList4)
                Cbbx1.Items.Add("Tracé précritique")
                Cbbx1.Items.Add("Départ temporo-occipital gauche")
                Cbbx1.Items.Add("F7")
                Cbbx1.Items.Add("Fp1 F7")
                Cbbx1.Items.Add("Quasi-hemisphérique gauche")
                Cbbx1.Items.Add("Départ fronto-temporo-occipital droit")
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
    Public Sub dispatcherTimer_Tick(ByVal sender As Object, ByVal e As EventArgs)
        If count < nFp1Array.Count - 1 Then
            count = count + 1
            coordelectrodes()
            Puissance()
            coherence()
            EEG()
            Scroll1.Value = count
            Textbox2.Text = count
        Else
            dt.Stop()
        End If
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
            '
            center(j)(0) = New Point(Leftpic + 0.57647 * PicList(0).ActualWidth, TopPic(j) + 0.169863 * PicList(0).ActualHeight)
            center(j)(1) = New Point(Leftpic + 0.724705882 * PicList(0).ActualWidth, TopPic(j) + 0.32055 * PicList(0).ActualHeight)
            center(j)(2) = New Point(Leftpic + 0.635294118 * PicList(0).ActualWidth, TopPic(j) + 0.52603 * PicList(0).ActualHeight)
            center(j)(3) = New Point(Leftpic + 0.729411765 * PicList(0).ActualWidth, TopPic(j) + 0.75342 * PicList(0).ActualHeight)
            center(j)(4) = New Point(Leftpic + 0.57647 * PicList(0).ActualWidth, TopPic(j) + 0.87397 * PicList(0).ActualHeight)
            center(j)(5) = New Point(Leftpic + 0.477647059 * PicList(0).ActualWidth, TopPic(j) + 0.52603 * PicList(0).ActualHeight)
            center(j)(6) = New Point(Leftpic + 0.376470588 * PicList(0).ActualWidth, TopPic(j) + 0.169863 * PicList(0).ActualHeight)
            center(j)(7) = New Point(Leftpic + 0.322352941 * PicList(0).ActualWidth, TopPic(j) + 0.52603 * PicList(0).ActualHeight)
            center(j)(8) = New Point(Leftpic + 0.211764706 * PicList(0).ActualWidth, TopPic(j) + 0.32055 * PicList(0).ActualHeight)
            center(j)(9) = New Point(Leftpic + 0.230588235 * PicList(0).ActualWidth, TopPic(j) + 0.7452 * PicList(0).ActualHeight)
            center(j)(10) = New Point(Leftpic + 0.385882353 * PicList(0).ActualWidth, TopPic(j) + 0.87397 * PicList(0).ActualHeight)
        Next
    End Sub
    Private Sub readExcelFile()
        Dim nApp As Excel.Application
        Dim nWorkbook As Excel.Workbook
        Dim nWorksheet As Excel.Worksheet
        nApp = New Excel.Application
        nWorkbook = nApp.Workbooks.Open(Textbox1.Text)
        nWorksheet = nWorkbook.Worksheets("P D")
        Dim nRange As Excel.Range = nWorksheet.UsedRange
        Dim nArray(,) As Object = nRange.Value(Excel.XlRangeValueDataType.xlRangeValueDefault)
        Dim nSize As Integer = nArray.GetUpperBound(0)
        For i As Integer = 2 To nSize
            nO1Array.Add(nArray(i, 15))
            nT5Array.Add(nArray(i, 16))
            nC3Array.Add(nArray(i, 17))
            nF7Array.Add(nArray(i, 18))
            nFp1Array.Add(nArray(i, 19))
            nCzArray.Add(nArray(i, 20))
            nO2array.Add(nArray(i, 21))
            nT6Array.Add(nArray(i, 22))
            nC4Array.Add(nArray(i, 23))
            nF8Array.Add(nArray(i, 24))
            nFp2Array.Add(nArray(i, 25))
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

        nWorksheet = nWorkbook.Worksheets("P T")
        Dim nRange1 As Excel.Range = nWorksheet.UsedRange
        Dim nArray1(,) As Object = nRange1.Value(Excel.XlRangeValueDataType.xlRangeValueDefault)
        Dim nSize1 As Integer = nArray.GetUpperBound(0)
        For i As Integer = 2 To nSize1
            nO1Array1.Add(nArray1(i, 15))
            nT5Array1.Add(nArray1(i, 16))
            nC3Array1.Add(nArray1(i, 17))
            nF7Array1.Add(nArray1(i, 18))
            nFp1Array1.Add(nArray1(i, 19))
            nCzArray1.Add(nArray1(i, 20))
            nO2array1.Add(nArray1(i, 21))
            nT6Array1.Add(nArray1(i, 22))
            nC4Array1.Add(nArray1(i, 23))
            nF8Array1.Add(nArray1(i, 24))
            nFp2Array1.Add(nArray1(i, 25))
        Next
        arrayList1.Add(nFp2Array1)
        arrayList1.Add(nF8Array1)
        arrayList1.Add(nC4Array1)
        arrayList1.Add(nT6Array1)
        arrayList1.Add(nO2array1)
        arrayList1.Add(nCzArray1)
        arrayList1.Add(nFp1Array1)
        arrayList1.Add(nC3Array1)
        arrayList1.Add(nF7Array1)
        arrayList1.Add(nT5Array1)
        arrayList1.Add(nO1Array1)

        nWorksheet = nWorkbook.Worksheets("P A")
        Dim nRange2 As Excel.Range = nWorksheet.UsedRange
        Dim nArray2(,) As Object = nRange2.Value(Excel.XlRangeValueDataType.xlRangeValueDefault)
        Dim nSize2 As Integer = nArray.GetUpperBound(0)
        For i As Integer = 2 To nSize2
            nO1Array2.Add(nArray2(i, 15))
            nT5Array2.Add(nArray2(i, 16))
            nC3Array2.Add(nArray2(i, 17))
            nF7Array2.Add(nArray2(i, 18))
            nFp1Array2.Add(nArray2(i, 19))
            nCzArray2.Add(nArray2(i, 20))
            nO2array2.Add(nArray2(i, 21))
            nT6Array2.Add(nArray2(i, 22))
            nC4Array2.Add(nArray2(i, 23))
            nF8Array2.Add(nArray2(i, 24))
            nFp2Array2.Add(nArray2(i, 25))
        Next
        arraylist2.Add(nFp2Array2)
        arraylist2.Add(nF8Array2)
        arraylist2.Add(nC4Array2)
        arraylist2.Add(nT6Array2)
        arraylist2.Add(nO2array2)
        arraylist2.Add(nCzArray2)
        arraylist2.Add(nFp1Array2)
        arraylist2.Add(nC3Array2)
        arraylist2.Add(nF7Array2)
        arraylist2.Add(nT5Array2)
        arraylist2.Add(nO1Array2)

        nWorksheet = nWorkbook.Worksheets("P B")
        Dim nRange3 As Excel.Range = nWorksheet.UsedRange
        Dim nArray3(,) As Object = nRange3.Value(Excel.XlRangeValueDataType.xlRangeValueDefault)
        Dim nSize3 As Integer = nArray.GetUpperBound(0)
        For i As Integer = 2 To nSize3
            nO1Array3.Add(nArray3(i, 15))
            nT5Array3.Add(nArray3(i, 16))
            nC3Array3.Add(nArray3(i, 17))
            nF7Array3.Add(nArray3(i, 18))
            nFp1Array3.Add(nArray3(i, 19))
            nCzArray3.Add(nArray3(i, 20))
            nO2array3.Add(nArray3(i, 21))
            nT6Array3.Add(nArray3(i, 22))
            nC4Array3.Add(nArray3(i, 23))
            nF8Array3.Add(nArray3(i, 24))
            nFp2Array3.Add(nArray3(i, 25))
        Next
        arraylist3.Add(nFp2Array3)
        arraylist3.Add(nF8Array3)
        arraylist3.Add(nC4Array3)
        arraylist3.Add(nT6Array3)
        arraylist3.Add(nO2array3)
        arraylist3.Add(nCzArray3)
        arraylist3.Add(nFp1Array3)
        arraylist3.Add(nC3Array3)
        arraylist3.Add(nF7Array3)
        arraylist3.Add(nT5Array3)
        arraylist3.Add(nO1Array3)

        nWorksheet = nWorkbook.Worksheets("P G")
        Dim nRange4 As Excel.Range = nWorksheet.UsedRange
        Dim narray4(,) As Object = nRange4.Value(Excel.XlRangeValueDataType.xlRangeValueDefault)
        Dim nSize4 As Integer = nArray.GetUpperBound(0)
        For i As Integer = 2 To nSize4
            nO1Array4.Add(narray4(i, 15))
            nT5Array4.Add(narray4(i, 16))
            nC3Array4.Add(narray4(i, 17))
            nF7Array4.Add(narray4(i, 18))
            nFp1Array4.Add(narray4(i, 19))
            nCzArray4.Add(narray4(i, 20))
            nO2array4.Add(narray4(i, 21))
            nT6Array4.Add(narray4(i, 22))
            nC4Array4.Add(narray4(i, 23))
            nF8Array4.Add(narray4(i, 24))
            nFp2Array4.Add(narray4(i, 25))
        Next
        arraylist4.Add(nFp2Array4)
        arraylist4.Add(nF8Array4)
        arraylist4.Add(nC4Array4)
        arraylist4.Add(nT6Array4)
        arraylist4.Add(nO2array4)
        arraylist4.Add(nCzArray4)
        arraylist4.Add(nFp1Array4)
        arraylist4.Add(nC3Array4)
        arraylist4.Add(nF7Array4)
        arraylist4.Add(nT5Array4)
        arraylist4.Add(nO1Array4)

        ListofArray.Add(arrayList)
        ListofArray.Add(arrayList1)
        ListofArray.Add(arraylist2)
        ListofArray.Add(arraylist3)
        ListofArray.Add(arraylist4)


        nWorksheet = nWorkbook.Worksheets("C D")
        Dim CRange As Excel.Range = nWorksheet.UsedRange
        Dim CArray(,) As Object = CRange.Value(Excel.XlRangeValueDataType.xlRangeValueDefault)
        Dim cSize As Integer = CArray.GetUpperBound(0)
        For i As Integer = 2 To cSize
            Fp2F8.Add(CArray(i, 1))
            Fp2C4.Add(CArray(i, 2))
            Fp2T6.Add(CArray(i, 3))
            Fp2O2.Add(CArray(i, 4))
            Fp2Cz.Add(CArray(i, 5))
            Fp2Fp1.Add(CArray(i, 6))
            Fp2F7.Add(CArray(i, 7))
            Fp2C3.Add(CArray(i, 8))
            Fp2T5.Add(CArray(i, 9))
            Fp2O1.Add(CArray(i, 10))
            F8C4.Add(CArray(i, 11))
            F8T6.Add(CArray(i, 12))
            F8O2.Add(CArray(i, 13))
            F8Cz.Add(CArray(i, 14))
            F8Fp1.Add(CArray(i, 15))
            F8F7.Add(CArray(i, 16))
            F8C3.Add(CArray(i, 17))
            F8T5.Add(CArray(i, 18))
            F8O1.Add(CArray(i, 19))
            C4T6.Add(CArray(i, 20))
            C4O2.Add(CArray(i, 21))
            C4Cz.Add(CArray(i, 22))
            C4Fp1.Add(CArray(i, 23))
            C4F7.Add(CArray(i, 24))
            C4C3.Add(CArray(i, 25))
            C4T5.Add(CArray(i, 26))
            C4O1.Add(CArray(i, 27))
            T6O2.Add(CArray(i, 28))
            T6Cz.Add(CArray(i, 29))
            T6Fp1.Add(CArray(i, 30))
            T6F7.Add(CArray(i, 31))
            T6C3.Add(CArray(i, 32))
            T6T5.Add(CArray(i, 33))
            T6O1.Add(CArray(i, 34))
            O2Cz.Add(CArray(i, 35))
            O2Fp1.Add(CArray(i, 36))
            O2F7.Add(CArray(i, 37))
            O2C3.Add(CArray(i, 38))
            O2T5.Add(CArray(i, 39))
            O2O1.Add(CArray(i, 40))
            CzFp1.Add(CArray(i, 41))
            CzF7.Add(CArray(i, 42))
            CzC3.Add(CArray(i, 43))
            CzT5.Add(CArray(i, 44))
            CzO1.Add(CArray(i, 45))
            Fp1F7.Add(CArray(i, 46))
            Fp1C3.Add(CArray(i, 47))
            Fp1T5.Add(CArray(i, 48))
            Fp1O1.Add(CArray(i, 49))
            F7C3.Add(CArray(i, 50))
            F7T5.Add(CArray(i, 51))
            F7O1.Add(CArray(i, 52))
            C3T5.Add(CArray(i, 53))
            C3O1.Add(CArray(i, 54))
            T5O1.Add(CArray(i, 55))
        Next
        CarrayList.Add(Fp2F8)
        CarrayList.Add(Fp2C4)
        CarrayList.Add(Fp2T6)
        CarrayList.Add(Fp2O2)
        CarrayList.Add(Fp2Cz)
        CarrayList.Add(Fp2Fp1)
        CarrayList.Add(Fp2F7)
        CarrayList.Add(Fp2C3)
        CarrayList.Add(Fp2T5)
        CarrayList.Add(Fp2O1)

        CarrayList.Add(F8C4)
        CarrayList.Add(F8T6)
        CarrayList.Add(F8O2)
        CarrayList.Add(F8Cz)
        CarrayList.Add(F8Fp1)
        CarrayList.Add(F8F7)
        CarrayList.Add(F8C3)
        CarrayList.Add(F8T5)
        CarrayList.Add(F8O1)

        CarrayList.Add(C4T6)
        CarrayList.Add(C4O2)
        CarrayList.Add(C4Cz)
        CarrayList.Add(C4Fp1)
        CarrayList.Add(C4F7)
        CarrayList.Add(C4C3)
        CarrayList.Add(C4T5)
        CarrayList.Add(C4O1)

        CarrayList.Add(T6O2)
        CarrayList.Add(T6Cz)
        CarrayList.Add(T6Fp1)
        CarrayList.Add(T6F7)
        CarrayList.Add(T6C3)
        CarrayList.Add(T6T5)
        CarrayList.Add(T6O1)

        CarrayList.Add(O2Cz)
        CarrayList.Add(O2Fp1)
        CarrayList.Add(O2F7)
        CarrayList.Add(O2C3)
        CarrayList.Add(O2T5)
        CarrayList.Add(O2O1)

        CarrayList.Add(CzFp1)
        CarrayList.Add(CzF7)
        CarrayList.Add(CzC3)
        CarrayList.Add(CzT5)
        CarrayList.Add(CzO1)

        CarrayList.Add(Fp1F7)
        CarrayList.Add(Fp1C3)
        CarrayList.Add(Fp1T5)
        CarrayList.Add(Fp1O1)

        CarrayList.Add(F7C3)
        CarrayList.Add(F7T5)
        CarrayList.Add(F7O1)

        CarrayList.Add(C3T5)
        CarrayList.Add(C3O1)

        CarrayList.Add(T5O1)

        nWorksheet = nWorkbook.Worksheets("C A")
        Dim CRange2 As Excel.Range = nWorksheet.UsedRange
        Dim CArray2(,) As Object = CRange2.Value(Excel.XlRangeValueDataType.xlRangeValueDefault)
        Dim cSize2 As Integer = CArray2.GetUpperBound(0)
        For i As Integer = 2 To cSize2
            Fp2F82.Add(CArray2(i, 1))
            Fp2C42.Add(CArray2(i, 2))
            Fp2T62.Add(CArray2(i, 3))
            Fp2O22.Add(CArray2(i, 4))
            Fp2Cz2.Add(CArray2(i, 5))
            Fp2Fp12.Add(CArray2(i, 6))
            Fp2F72.Add(CArray2(i, 7))
            Fp2C32.Add(CArray2(i, 8))
            Fp2T52.Add(CArray2(i, 9))
            Fp2O12.Add(CArray2(i, 10))
            F8C42.Add(CArray2(i, 11))
            F8T62.Add(CArray2(i, 12))
            F8O22.Add(CArray2(i, 13))
            F8Cz2.Add(CArray2(i, 14))
            F8Fp12.Add(CArray2(i, 15))
            F8F72.Add(CArray2(i, 16))
            F8C32.Add(CArray2(i, 17))
            F8T52.Add(CArray2(i, 18))
            F8O12.Add(CArray2(i, 19))
            C4T62.Add(CArray2(i, 20))
            C4O22.Add(CArray2(i, 21))
            C4Cz2.Add(CArray2(i, 22))
            C4Fp12.Add(CArray2(i, 23))
            C4F72.Add(CArray2(i, 24))
            C4C32.Add(CArray2(i, 25))
            C4T52.Add(CArray2(i, 26))
            C4O12.Add(CArray2(i, 27))
            T6O22.Add(CArray2(i, 28))
            T6Cz2.Add(CArray2(i, 29))
            T6Fp12.Add(CArray2(i, 30))
            T6F72.Add(CArray2(i, 31))
            T6C32.Add(CArray2(i, 32))
            T6T52.Add(CArray2(i, 33))
            T6O12.Add(CArray2(i, 34))
            O2Cz2.Add(CArray2(i, 35))
            O2Fp12.Add(CArray2(i, 36))
            O2F72.Add(CArray2(i, 37))
            O2C32.Add(CArray2(i, 38))
            O2T52.Add(CArray2(i, 39))
            O2O12.Add(CArray2(i, 40))
            CzFp12.Add(CArray2(i, 41))
            CzF72.Add(CArray2(i, 42))
            CzC32.Add(CArray2(i, 43))
            CzT52.Add(CArray2(i, 44))
            CzO12.Add(CArray2(i, 45))
            Fp1F72.Add(CArray2(i, 46))
            Fp1C32.Add(CArray2(i, 47))
            Fp1T52.Add(CArray2(i, 48))
            Fp1O12.Add(CArray2(i, 49))
            F7C32.Add(CArray2(i, 50))
            F7T52.Add(CArray2(i, 51))
            F7O12.Add(CArray2(i, 52))
            C3T52.Add(CArray2(i, 53))
            C3O12.Add(CArray2(i, 54))
            T5O12.Add(CArray2(i, 55))
        Next
        CarrayList2.Add(Fp2F82)
        CarrayList2.Add(Fp2C42)
        CarrayList2.Add(Fp2T62)
        CarrayList2.Add(Fp2O22)
        CarrayList2.Add(Fp2Cz2)
        CarrayList2.Add(Fp2Fp12)
        CarrayList2.Add(Fp2F72)
        CarrayList2.Add(Fp2C32)
        CarrayList2.Add(Fp2T52)
        CarrayList2.Add(Fp2O12)

        CarrayList2.Add(F8C42)
        CarrayList2.Add(F8T62)
        CarrayList2.Add(F8O22)
        CarrayList2.Add(F8Cz2)
        CarrayList2.Add(F8Fp12)
        CarrayList2.Add(F8F72)
        CarrayList2.Add(F8C32)
        CarrayList2.Add(F8T52)
        CarrayList2.Add(F8O12)

        CarrayList2.Add(C4T62)
        CarrayList2.Add(C4O22)
        CarrayList2.Add(C4Cz2)
        CarrayList2.Add(C4Fp12)
        CarrayList2.Add(C4F72)
        CarrayList2.Add(C4C32)
        CarrayList2.Add(C4T52)
        CarrayList2.Add(C4O12)

        CarrayList2.Add(T6O22)
        CarrayList2.Add(T6Cz2)
        CarrayList2.Add(T6Fp12)
        CarrayList2.Add(T6F72)
        CarrayList2.Add(T6C32)
        CarrayList2.Add(T6T52)
        CarrayList2.Add(T6O12)

        CarrayList2.Add(O2Cz2)
        CarrayList2.Add(O2Fp12)
        CarrayList2.Add(O2F72)
        CarrayList2.Add(O2C32)
        CarrayList2.Add(O2T52)
        CarrayList2.Add(O2O12)

        CarrayList2.Add(CzFp12)
        CarrayList2.Add(CzF72)
        CarrayList2.Add(CzC32)
        CarrayList2.Add(CzT52)
        CarrayList2.Add(CzO12)

        CarrayList2.Add(Fp1F72)
        CarrayList2.Add(Fp1C32)
        CarrayList2.Add(Fp1T52)
        CarrayList2.Add(Fp1O12)

        CarrayList2.Add(F7C32)
        CarrayList2.Add(F7T52)
        CarrayList2.Add(F7O12)

        CarrayList2.Add(C3T52)
        CarrayList2.Add(C3O12)

        CarrayList2.Add(T5O12)


        nWorksheet = nWorkbook.Worksheets("C B")
        Dim CRange3 As Excel.Range = nWorksheet.UsedRange
        Dim CArray3(,) As Object = CRange3.Value(Excel.XlRangeValueDataType.xlRangeValueDefault)
        Dim cSize3 As Integer = CArray3.GetUpperBound(0)
        For i As Integer = 2 To cSize3
            Fp2F83.Add(CArray3(i, 1))
            Fp2C43.Add(CArray3(i, 2))
            Fp2T63.Add(CArray3(i, 3))
            Fp2O23.Add(CArray3(i, 4))
            Fp2Cz3.Add(CArray3(i, 5))
            Fp2Fp13.Add(CArray3(i, 6))
            Fp2F73.Add(CArray3(i, 7))
            Fp2C33.Add(CArray3(i, 8))
            Fp2T53.Add(CArray3(i, 9))
            Fp2O13.Add(CArray3(i, 10))
            F8C43.Add(CArray3(i, 11))
            F8T63.Add(CArray3(i, 12))
            F8O23.Add(CArray3(i, 13))
            F8Cz3.Add(CArray3(i, 14))
            F8Fp13.Add(CArray3(i, 15))
            F8F73.Add(CArray3(i, 16))
            F8C33.Add(CArray3(i, 17))
            F8T53.Add(CArray3(i, 18))
            F8O13.Add(CArray3(i, 19))
            C4T63.Add(CArray3(i, 20))
            C4O23.Add(CArray3(i, 21))
            C4Cz3.Add(CArray3(i, 22))
            C4Fp13.Add(CArray3(i, 23))
            C4F73.Add(CArray3(i, 24))
            C4C33.Add(CArray3(i, 25))
            C4T53.Add(CArray3(i, 26))
            C4O13.Add(CArray3(i, 27))
            T6O23.Add(CArray3(i, 28))
            T6Cz3.Add(CArray3(i, 29))
            T6Fp13.Add(CArray3(i, 30))
            T6F73.Add(CArray3(i, 31))
            T6C33.Add(CArray3(i, 32))
            T6T53.Add(CArray3(i, 33))
            T6O13.Add(CArray3(i, 34))
            O2Cz3.Add(CArray3(i, 35))
            O2Fp13.Add(CArray3(i, 36))
            O2F73.Add(CArray3(i, 37))
            O2C33.Add(CArray3(i, 38))
            O2T53.Add(CArray3(i, 39))
            O2O13.Add(CArray3(i, 40))
            CzFp13.Add(CArray3(i, 41))
            CzF73.Add(CArray3(i, 42))
            CzC33.Add(CArray3(i, 43))
            CzT53.Add(CArray3(i, 44))
            CzO13.Add(CArray3(i, 45))
            Fp1F73.Add(CArray3(i, 46))
            Fp1C33.Add(CArray3(i, 47))
            Fp1T53.Add(CArray3(i, 48))
            Fp1O13.Add(CArray3(i, 49))
            F7C33.Add(CArray3(i, 50))
            F7T53.Add(CArray3(i, 51))
            F7O13.Add(CArray3(i, 52))
            C3T53.Add(CArray3(i, 53))
            C3O13.Add(CArray3(i, 54))
            T5O13.Add(CArray3(i, 55))
        Next
        CarrayList3.Add(Fp2F83)
        CarrayList3.Add(Fp2C43)
        CarrayList3.Add(Fp2T63)
        CarrayList3.Add(Fp2O23)
        CarrayList3.Add(Fp2Cz3)
        CarrayList3.Add(Fp2Fp13)
        CarrayList3.Add(Fp2F73)
        CarrayList3.Add(Fp2C33)
        CarrayList3.Add(Fp2T53)
        CarrayList3.Add(Fp2O13)

        CarrayList3.Add(F8C43)
        CarrayList3.Add(F8T63)
        CarrayList3.Add(F8O23)
        CarrayList3.Add(F8Cz3)
        CarrayList3.Add(F8Fp13)
        CarrayList3.Add(F8F73)
        CarrayList3.Add(F8C33)
        CarrayList3.Add(F8T53)
        CarrayList3.Add(F8O13)

        CarrayList3.Add(C4T63)
        CarrayList3.Add(C4O23)
        CarrayList3.Add(C4Cz3)
        CarrayList3.Add(C4Fp13)
        CarrayList3.Add(C4F73)
        CarrayList3.Add(C4C33)
        CarrayList3.Add(C4T53)
        CarrayList3.Add(C4O13)

        CarrayList3.Add(T6O23)
        CarrayList3.Add(T6Cz3)
        CarrayList3.Add(T6Fp13)
        CarrayList3.Add(T6F73)
        CarrayList3.Add(T6C33)
        CarrayList3.Add(T6T53)
        CarrayList3.Add(T6O13)

        CarrayList3.Add(O2Cz3)
        CarrayList3.Add(O2Fp13)
        CarrayList3.Add(O2F73)
        CarrayList3.Add(O2C33)
        CarrayList3.Add(O2T53)
        CarrayList3.Add(O2O13)

        CarrayList3.Add(CzFp13)
        CarrayList3.Add(CzF73)
        CarrayList3.Add(CzC33)
        CarrayList3.Add(CzT53)
        CarrayList3.Add(CzO13)

        CarrayList3.Add(Fp1F73)
        CarrayList3.Add(Fp1C33)
        CarrayList3.Add(Fp1T53)
        CarrayList3.Add(Fp1O13)

        CarrayList3.Add(F7C33)
        CarrayList3.Add(F7T53)
        CarrayList3.Add(F7O13)

        CarrayList3.Add(C3T53)
        CarrayList3.Add(C3O13)

        CarrayList3.Add(T5O13)

        nWorksheet = nWorkbook.Worksheets("C G")
        Dim CRange4 As Excel.Range = nWorksheet.UsedRange
        Dim CArray4(,) As Object = CRange4.Value(Excel.XlRangeValueDataType.xlRangeValueDefault)
        Dim cSize4 As Integer = CArray4.GetUpperBound(0)
        For i As Integer = 2 To cSize4
            Fp2F84.Add(CArray4(i, 1))
            Fp2C44.Add(CArray4(i, 2))
            Fp2T64.Add(CArray4(i, 3))
            Fp2O24.Add(CArray4(i, 4))
            Fp2Cz4.Add(CArray4(i, 5))
            Fp2Fp14.Add(CArray4(i, 6))
            Fp2F74.Add(CArray4(i, 7))
            Fp2C34.Add(CArray4(i, 8))
            Fp2T54.Add(CArray4(i, 9))
            Fp2O14.Add(CArray4(i, 10))
            F8C44.Add(CArray4(i, 11))
            F8T64.Add(CArray4(i, 12))
            F8O24.Add(CArray4(i, 13))
            F8Cz4.Add(CArray4(i, 14))
            F8Fp14.Add(CArray4(i, 15))
            F8F74.Add(CArray4(i, 16))
            F8C34.Add(CArray4(i, 17))
            F8T54.Add(CArray4(i, 18))
            F8O14.Add(CArray4(i, 19))
            C4T64.Add(CArray4(i, 20))
            C4O24.Add(CArray4(i, 21))
            C4Cz4.Add(CArray4(i, 22))
            C4Fp14.Add(CArray4(i, 23))
            C4F74.Add(CArray4(i, 24))
            C4C34.Add(CArray4(i, 25))
            C4T54.Add(CArray4(i, 26))
            C4O14.Add(CArray4(i, 27))
            T6O24.Add(CArray4(i, 28))
            T6Cz4.Add(CArray4(i, 29))
            T6Fp14.Add(CArray4(i, 30))
            T6F74.Add(CArray4(i, 31))
            T6C34.Add(CArray4(i, 32))
            T6T54.Add(CArray4(i, 33))
            T6O14.Add(CArray4(i, 34))
            O2Cz4.Add(CArray4(i, 35))
            O2Fp14.Add(CArray4(i, 36))
            O2F74.Add(CArray4(i, 37))
            O2C34.Add(CArray4(i, 38))
            O2T54.Add(CArray4(i, 39))
            O2O14.Add(CArray4(i, 40))
            CzFp14.Add(CArray4(i, 41))
            CzF74.Add(CArray4(i, 42))
            CzC34.Add(CArray4(i, 43))
            CzT54.Add(CArray4(i, 44))
            CzO14.Add(CArray4(i, 45))
            Fp1F74.Add(CArray4(i, 46))
            Fp1C34.Add(CArray4(i, 47))
            Fp1T54.Add(CArray4(i, 48))
            Fp1O14.Add(CArray4(i, 49))
            F7C34.Add(CArray4(i, 50))
            F7T54.Add(CArray4(i, 51))
            F7O14.Add(CArray4(i, 52))
            C3T54.Add(CArray4(i, 53))
            C3O14.Add(CArray4(i, 54))
            T5O14.Add(CArray4(i, 55))
        Next
        CarrayList4.Add(Fp2F84)
        CarrayList4.Add(Fp2C44)
        CarrayList4.Add(Fp2T64)
        CarrayList4.Add(Fp2O24)
        CarrayList4.Add(Fp2Cz4)
        CarrayList4.Add(Fp2Fp14)
        CarrayList4.Add(Fp2F74)
        CarrayList4.Add(Fp2C34)
        CarrayList4.Add(Fp2T54)
        CarrayList4.Add(Fp2O14)

        CarrayList4.Add(F8C44)
        CarrayList4.Add(F8T64)
        CarrayList4.Add(F8O24)
        CarrayList4.Add(F8Cz4)
        CarrayList4.Add(F8Fp14)
        CarrayList4.Add(F8F74)
        CarrayList4.Add(F8C34)
        CarrayList4.Add(F8T54)
        CarrayList4.Add(F8O14)

        CarrayList4.Add(C4T64)
        CarrayList4.Add(C4O24)
        CarrayList4.Add(C4Cz4)
        CarrayList4.Add(C4Fp14)
        CarrayList4.Add(C4F74)
        CarrayList4.Add(C4C34)
        CarrayList4.Add(C4T54)
        CarrayList4.Add(C4O14)

        CarrayList4.Add(T6O24)
        CarrayList4.Add(T6Cz4)
        CarrayList4.Add(T6Fp14)
        CarrayList4.Add(T6F74)
        CarrayList4.Add(T6C34)
        CarrayList4.Add(T6T54)
        CarrayList4.Add(T6O14)

        CarrayList4.Add(O2Cz4)
        CarrayList4.Add(O2Fp14)
        CarrayList4.Add(O2F74)
        CarrayList4.Add(O2C34)
        CarrayList4.Add(O2T54)
        CarrayList4.Add(O2O14)

        CarrayList4.Add(CzFp14)
        CarrayList4.Add(CzF74)
        CarrayList4.Add(CzC34)
        CarrayList4.Add(CzT54)
        CarrayList4.Add(CzO14)

        CarrayList4.Add(Fp1F74)
        CarrayList4.Add(Fp1C34)
        CarrayList4.Add(Fp1T54)
        CarrayList4.Add(Fp1O14)

        CarrayList4.Add(F7C34)
        CarrayList4.Add(F7T54)
        CarrayList4.Add(F7O14)

        CarrayList4.Add(C3T54)
        CarrayList4.Add(C3O14)

        CarrayList4.Add(T5O14)

        nWorksheet = nWorkbook.Worksheets("C T")
        Dim CRange1 As Excel.Range = nWorksheet.UsedRange
        Dim CArray1(,) As Object = CRange1.Value(Excel.XlRangeValueDataType.xlRangeValueDefault)
        Dim cSize1 As Integer = CArray1.GetUpperBound(0)
        For i As Integer = 2 To cSize1
            Fp2F81.Add(CArray1(i, 1))
            Fp2C41.Add(CArray1(i, 2))
            Fp2T61.Add(CArray1(i, 3))
            Fp2O21.Add(CArray1(i, 4))
            Fp2Cz1.Add(CArray1(i, 5))
            Fp2Fp11.Add(CArray1(i, 6))
            Fp2F71.Add(CArray1(i, 7))
            Fp2C31.Add(CArray1(i, 8))
            Fp2T51.Add(CArray1(i, 9))
            Fp2O11.Add(CArray1(i, 10))
            F8C41.Add(CArray1(i, 11))
            F8T61.Add(CArray1(i, 12))
            F8O21.Add(CArray1(i, 13))
            F8Cz1.Add(CArray1(i, 14))
            F8Fp11.Add(CArray1(i, 15))
            F8F71.Add(CArray1(i, 16))
            F8C31.Add(CArray1(i, 17))
            F8T51.Add(CArray1(i, 18))
            F8O11.Add(CArray1(i, 19))
            C4T61.Add(CArray1(i, 20))
            C4O21.Add(CArray1(i, 21))
            C4Cz1.Add(CArray1(i, 22))
            C4Fp11.Add(CArray1(i, 23))
            C4F71.Add(CArray1(i, 24))
            C4C31.Add(CArray1(i, 25))
            C4T51.Add(CArray1(i, 26))
            C4O11.Add(CArray1(i, 27))
            T6O21.Add(CArray1(i, 28))
            T6Cz1.Add(CArray1(i, 29))
            T6Fp11.Add(CArray1(i, 30))
            T6F71.Add(CArray1(i, 31))
            T6C31.Add(CArray1(i, 32))
            T6T51.Add(CArray1(i, 33))
            T6O11.Add(CArray1(i, 34))
            O2Cz1.Add(CArray1(i, 35))
            O2Fp11.Add(CArray1(i, 36))
            O2F71.Add(CArray1(i, 37))
            O2C31.Add(CArray1(i, 38))
            O2T51.Add(CArray1(i, 39))
            O2O11.Add(CArray1(i, 40))
            CzFp11.Add(CArray1(i, 41))
            CzF71.Add(CArray1(i, 42))
            CzC31.Add(CArray1(i, 43))
            CzT51.Add(CArray1(i, 44))
            CzO11.Add(CArray1(i, 45))
            Fp1F71.Add(CArray1(i, 46))
            Fp1C31.Add(CArray1(i, 47))
            Fp1T51.Add(CArray1(i, 48))
            Fp1O11.Add(CArray1(i, 49))
            F7C31.Add(CArray1(i, 50))
            F7T51.Add(CArray1(i, 51))
            F7O11.Add(CArray1(i, 52))
            C3T51.Add(CArray1(i, 53))
            C3O11.Add(CArray1(i, 54))
            T5O11.Add(CArray1(i, 55))
        Next
        CarrayList1.Add(Fp2F81)
        CarrayList1.Add(Fp2C41)
        CarrayList1.Add(Fp2T61)
        CarrayList1.Add(Fp2O21)
        CarrayList1.Add(Fp2Cz1)
        CarrayList1.Add(Fp2Fp11)
        CarrayList1.Add(Fp2F71)
        CarrayList1.Add(Fp2C31)
        CarrayList1.Add(Fp2T51)
        CarrayList1.Add(Fp2O11)

        CarrayList1.Add(F8C41)
        CarrayList1.Add(F8T61)
        CarrayList1.Add(F8O21)
        CarrayList1.Add(F8Cz1)
        CarrayList1.Add(F8Fp11)
        CarrayList1.Add(F8F71)
        CarrayList1.Add(F8C31)
        CarrayList1.Add(F8T51)
        CarrayList1.Add(F8O11)

        CarrayList1.Add(C4T61)
        CarrayList1.Add(C4O21)
        CarrayList1.Add(C4Cz1)
        CarrayList1.Add(C4Fp11)
        CarrayList1.Add(C4F71)
        CarrayList1.Add(C4C31)
        CarrayList1.Add(C4T51)
        CarrayList1.Add(C4O11)

        CarrayList1.Add(T6O21)
        CarrayList1.Add(T6Cz1)
        CarrayList1.Add(T6Fp11)
        CarrayList1.Add(T6F71)
        CarrayList1.Add(T6C31)
        CarrayList1.Add(T6T51)
        CarrayList1.Add(T6O11)

        CarrayList1.Add(O2Cz1)
        CarrayList1.Add(O2Fp11)
        CarrayList1.Add(O2F71)
        CarrayList1.Add(O2C31)
        CarrayList1.Add(O2T51)
        CarrayList1.Add(O2O11)

        CarrayList1.Add(CzFp11)
        CarrayList1.Add(CzF71)
        CarrayList1.Add(CzC31)
        CarrayList1.Add(CzT51)
        CarrayList1.Add(CzO11)

        CarrayList1.Add(Fp1F71)
        CarrayList1.Add(Fp1C31)
        CarrayList1.Add(Fp1T51)
        CarrayList1.Add(Fp1O11)

        CarrayList1.Add(F7C31)
        CarrayList1.Add(F7T51)
        CarrayList1.Add(F7O11)

        CarrayList1.Add(C3T51)
        CarrayList1.Add(C3O11)

        CarrayList1.Add(T5O11)



        CListofArray.Add(CarrayList)
        CListofArray.Add(CarrayList1)
        CListofArray.Add(CarrayList2)
        CListofArray.Add(CarrayList3)
        CListofArray.Add(CarrayList4)

    End Sub
    Private Sub Coordonnee(i As Integer, E As Integer)
        If i < 10 Then
            ListofLine(E)(i).X1 = center(E)(0).X
            ListofLine(E)(i).Y1 = center(E)(0).Y
            Select Case i
                Case 0
                    ListofLine(E)(i).X2 = center(E)(1).X
                    ListofLine(E)(i).Y2 = center(E)(1).Y
                Case 1
                    ListofLine(E)(i).X2 = center(E)(2).X
                    ListofLine(E)(i).Y2 = center(E)(2).Y
                Case 2
                    ListofLine(E)(i).X2 = center(E)(3).X
                    ListofLine(E)(i).Y2 = center(E)(3).Y
                Case 3
                    ListofLine(E)(i).X2 = center(E)(4).X
                    ListofLine(E)(i).Y2 = center(E)(4).Y
                Case 4
                    ListofLine(E)(i).X2 = center(E)(5).X
                    ListofLine(E)(i).Y2 = center(E)(5).Y
                Case 5
                    ListofLine(E)(i).X2 = center(E)(6).X
                    ListofLine(E)(i).Y2 = center(E)(6).Y
                Case 6
                    ListofLine(E)(i).X2 = center(E)(7).X
                    ListofLine(E)(i).Y2 = center(E)(7).Y
                Case 7
                    ListofLine(E)(i).X2 = center(E)(8).X
                    ListofLine(E)(i).Y2 = center(E)(8).Y
                Case 8
                    ListofLine(E)(i).X2 = center(E)(9).X
                    ListofLine(E)(i).Y2 = center(E)(9).Y
                Case 9
                    ListofLine(E)(i).X2 = center(E)(10).X
                    ListofLine(E)(i).Y2 = center(E)(10).Y
            End Select
        ElseIf i < 19 And i > 9 Then
            ListofLine(E)(i).X1 = center(E)(1).X
            ListofLine(E)(i).Y1 = center(E)(1).Y
            Select Case i
                Case 10
                    ListofLine(E)(i).X2 = center(E)(2).X
                    ListofLine(E)(i).Y2 = center(E)(2).Y
                Case 11
                    ListofLine(E)(i).X2 = center(E)(3).X
                    ListofLine(E)(i).Y2 = center(E)(3).Y
                Case 12
                    ListofLine(E)(i).X2 = center(E)(4).X
                    ListofLine(E)(i).Y2 = center(E)(4).Y
                Case 13
                    ListofLine(E)(i).X2 = center(E)(5).X
                    ListofLine(E)(i).Y2 = center(E)(5).Y
                Case 14
                    ListofLine(E)(i).X2 = center(E)(6).X
                    ListofLine(E)(i).Y2 = center(E)(6).Y
                Case 15
                    ListofLine(E)(i).X2 = center(E)(7).X
                    ListofLine(E)(i).Y2 = center(E)(7).Y
                Case 16
                    ListofLine(E)(i).X2 = center(E)(8).X
                    ListofLine(E)(i).Y2 = center(E)(8).Y
                Case 17
                    ListofLine(E)(i).X2 = center(E)(9).X
                    ListofLine(E)(i).Y2 = center(E)(9).Y
                Case 18
                    ListofLine(E)(i).X2 = center(E)(10).X
                    ListofLine(E)(i).Y2 = center(E)(10).Y
            End Select
        ElseIf i < 27 And i > 18 Then
            ListofLine(E)(i).X1 = center(E)(2).X
            ListofLine(E)(i).Y1 = center(E)(2).Y
            Select Case i
                Case 19
                    ListofLine(E)(i).X2 = center(E)(3).X
                    ListofLine(E)(i).Y2 = center(E)(3).Y
                Case 20
                    ListofLine(E)(i).X2 = center(E)(4).X
                    ListofLine(E)(i).Y2 = center(E)(4).Y
                Case 21
                    ListofLine(E)(i).X2 = center(E)(5).X
                    ListofLine(E)(i).Y2 = center(E)(5).Y
                Case 22
                    ListofLine(E)(i).X2 = center(E)(6).X
                    ListofLine(E)(i).Y2 = center(E)(6).Y
                Case 23
                    ListofLine(E)(i).X2 = center(E)(7).X
                    ListofLine(E)(i).Y2 = center(E)(7).Y
                Case 24
                    ListofLine(E)(i).X2 = center(E)(8).X
                    ListofLine(E)(i).Y2 = center(E)(8).Y
                Case 25
                    ListofLine(E)(i).X2 = center(E)(9).X
                    ListofLine(E)(i).Y2 = center(E)(9).Y
                Case 26
                    ListofLine(E)(i).X2 = center(E)(10).X
                    ListofLine(E)(i).Y2 = center(E)(10).Y
            End Select
        ElseIf i < 34 And i > 26 Then
            ListofLine(E)(i).X1 = center(E)(3).X
            ListofLine(E)(i).Y1 = center(E)(3).Y
            Select Case i
                Case 27
                    ListofLine(E)(i).X2 = center(E)(4).X
                    ListofLine(E)(i).Y2 = center(E)(4).Y
                Case 28
                    ListofLine(E)(i).X2 = center(E)(5).X
                    ListofLine(E)(i).Y2 = center(E)(5).Y
                Case 29
                    ListofLine(E)(i).X2 = center(E)(6).X
                    ListofLine(E)(i).Y2 = center(E)(6).Y
                Case 30
                    ListofLine(E)(i).X2 = center(E)(7).X
                    ListofLine(E)(i).Y2 = center(E)(7).Y
                Case 31
                    ListofLine(E)(i).X2 = center(E)(8).X
                    ListofLine(E)(i).Y2 = center(E)(8).Y
                Case 32
                    ListofLine(E)(i).X2 = center(E)(9).X
                    ListofLine(E)(i).Y2 = center(E)(9).Y
                Case 33
                    ListofLine(E)(i).X2 = center(E)(10).X
                    ListofLine(E)(i).Y2 = center(E)(10).Y
            End Select
        ElseIf i < 40 And i > 33 Then
            ListofLine(E)(i).X1 = center(E)(4).X
            ListofLine(E)(i).Y1 = center(E)(4).Y
            Select Case i
                Case 34
                    ListofLine(E)(i).X2 = center(E)(5).X
                    ListofLine(E)(i).Y2 = center(E)(5).Y
                Case 35
                    ListofLine(E)(i).X2 = center(E)(6).X
                    ListofLine(E)(i).Y2 = center(E)(6).Y
                Case 36
                    ListofLine(E)(i).X2 = center(E)(7).X
                    ListofLine(E)(i).Y2 = center(E)(7).Y
                Case 37
                    ListofLine(E)(i).X2 = center(E)(8).X
                    ListofLine(E)(i).Y2 = center(E)(8).Y
                Case 38
                    ListofLine(E)(i).X2 = center(E)(9).X
                    ListofLine(E)(i).Y2 = center(E)(9).Y
                Case 39
                    ListofLine(E)(i).X2 = center(E)(10).X
                    ListofLine(E)(i).Y2 = center(E)(10).Y
            End Select
        ElseIf i < 45 And i > 39 Then
            ListofLine(E)(i).X1 = center(E)(5).X
            ListofLine(E)(i).Y1 = center(E)(5).Y
            Select Case i
                Case 40
                    ListofLine(E)(i).X2 = center(E)(6).X
                    ListofLine(E)(i).Y2 = center(E)(6).Y
                Case 41
                    ListofLine(E)(i).X2 = center(E)(7).X
                    ListofLine(E)(i).Y2 = center(E)(7).Y
                Case 42
                    ListofLine(E)(i).X2 = center(E)(8).X
                    ListofLine(E)(i).Y2 = center(E)(8).Y
                Case 43
                    ListofLine(E)(i).X2 = center(E)(9).X
                    ListofLine(E)(i).Y2 = center(E)(9).Y
                Case 44
                    ListofLine(E)(i).X2 = center(E)(10).X
                    ListofLine(E)(i).Y2 = center(E)(10).Y
            End Select
        ElseIf i < 49 And i > 44 Then
            ListofLine(E)(i).X1 = center(E)(6).X
            ListofLine(E)(i).Y1 = center(E)(6).Y
            Select Case i
                Case 45
                    ListofLine(E)(i).X2 = center(E)(7).X
                    ListofLine(E)(i).Y2 = center(E)(7).Y
                Case 46
                    ListofLine(E)(i).X2 = center(E)(8).X
                    ListofLine(E)(i).Y2 = center(E)(8).Y
                Case 47
                    ListofLine(E)(i).X2 = center(E)(9).X
                    ListofLine(E)(i).Y2 = center(E)(9).Y
                Case 48
                    ListofLine(E)(i).X2 = center(E)(10).X
                    ListofLine(E)(i).Y2 = center(E)(10).Y
            End Select
        ElseIf i < 52 And i > 48 Then
            ListofLine(E)(i).X1 = center(E)(7).X
            ListofLine(E)(i).Y1 = center(E)(7).Y
            Select Case i
                Case 49
                    ListofLine(E)(i).X2 = center(E)(8).X
                    ListofLine(E)(i).Y2 = center(E)(8).Y
                Case 50
                    ListofLine(E)(i).X2 = center(E)(9).X
                    ListofLine(E)(i).Y2 = center(E)(9).Y
                Case 51
                    ListofLine(E)(i).X2 = center(E)(10).X
                    ListofLine(E)(i).Y2 = center(E)(10).Y
            End Select
        ElseIf i < 54 And i > 51 Then
            ListofLine(E)(i).X1 = center(E)(8).X
            ListofLine(E)(i).Y1 = center(E)(8).Y
            Select Case i
                Case 52
                    ListofLine(E)(i).X2 = center(E)(9).X
                    ListofLine(E)(i).Y2 = center(E)(9).Y
                Case 53
                    ListofLine(E)(i).X2 = center(E)(10).X
                    ListofLine(E)(i).Y2 = center(E)(10).Y
            End Select
        Else
            ListofLine(E)(i).X1 = center(E)(9).X
            ListofLine(E)(i).Y1 = center(E)(9).Y
            ListofLine(E)(i).X2 = center(E)(10).X
            ListofLine(E)(i).Y2 = center(E)(10).Y
        End If
    End Sub
    Private Sub coherence()
        mySolidColorBrush1.Color = Color.FromRgb(0, 0, 153)
        For e = 0 To 4
            For i As Integer = 0 To 54
                Canvas1.Children.Remove(ListCanvasC(e)(i))
                ListCanvasC(e)(i).Children.Remove(ListofLine(e)(i))
                Coh = CListofArray(e)(i)(count)
                If Coh > SeuilCoh Then
                    Coordonnee(i, e)
                    If Coh < (1 - SeuilCoh) / 4 Then
                        ListofLine(e)(i).Stroke = mySolidColorBrush1
                    ElseIf Coh < (2 * (1 - SeuilCoh)) / 4 And Coh > ((1 - SeuilCoh) / 4) - 0.000001 Then
                        ListofLine(e)(i).Stroke = mySolidColorBrush2
                    ElseIf Coh < (3 * (1 - SeuilCoh)) / 4 And Coh > (2 * (1 - SeuilCoh) / 4) - 0.000001 Then
                        ListofLine(e)(i).Stroke = mySolidColorBrush3
                    Else
                        ListofLine(e)(i).Stroke = mySolidColorBrush4
                    End If
                    ListofLine(e)(i).StrokeThickness = 5
                    ListCanvasC(e)(i).Children.Add(ListofLine(e)(i))
                    Canvas1.Children.Add(ListCanvasC(e)(i))
                End If
            Next
        Next
        Me.Content = Canvas1
        Scroll1.Value = count
    End Sub
    Private Sub Puissance()
        For T = 0 To 4
            mySolidColorBrush1.Color = Color.FromRgb(0, 0, 153)
            mySolidColorBrush2.Color = Color.FromRgb(51, 204, 0)
            mySolidColorBrush3.Color = Color.FromRgb(255, 204, 51)
            mySolidColorBrush4.Color = Color.FromRgb(255, 0, 0)
            For i As Integer = 0 To 10
                Canvas1.Children.Remove(ListCanvasP(T)(i))
                ListCanvasP(T)(i).Children.Remove(ListEl(T)(i))
                Item = ListofArray(T)(i)(count)
                If Item < 5 Then
                    r = 5 * echelle
                    ListEl(T)(i).Fill = mySolidColorBrush1
                ElseIf Item > 4 And Item < 11 Then
                    r = Item * echelle
                    ListEl(T)(i).Fill = mySolidColorBrush2
                ElseIf Item > 10 And Item < 16 Then
                    r = Item * echelle
                    ListEl(T)(i).Fill = mySolidColorBrush3
                Else
                    r = Item * echelle
                    ListEl(T)(i).Fill = mySolidColorBrush4
                End If
                ListEl(T)(i).StrokeThickness = 5
                ListEl(T)(i).Width = r / B
                ListEl(T)(i).Height = r / B
                ListCanvasP(T)(i).Height = 100
                ListCanvasP(T)(i).Width = 100
                Canvas.SetTop(ListCanvasP(T)(i), center(T)(i).Y - (r / (2 * B)))
                Canvas.SetLeft(ListCanvasP(T)(i), center(T)(i).X - (r / (2 * B)))
                ListCanvasP(T)(i).Children.Add(ListEl(T)(i))
                Canvas1.Children.Add(ListCanvasP(T)(i))
            Next
        Next
        Textbox2.Text = count
        Select Case count
            Case Is < 38
                Cbbx1.Text = "Tracé précritique"
            Case Is < 60
                Cbbx1.Text = "Départ temporo-occipital gauche"
            Case Is < 100
                Cbbx1.Text = "F7"
            Case Is < 170
                Cbbx1.Text = "Fp1 F7"
            Case Is < 360
                Cbbx1.Text = "Quasi-hemisphérique gauche"
            Case Is < 530
                Cbbx1.Text = "Départ fronto-temporo-occipital droit"
            Case Is < 773
                Cbbx1.Text = "Crise exclusivement droite"
            Case Else
                Cbbx1.Text = "Tracé postcritique"
        End select
        Me.Content = Canvas1

    End Sub
    Private Sub EEG()
        LineEEG.Y1 = Canvas.GetTop(CanvaspictureEEG)
        LineEEG.Y2 = PictureEEG.ActualHeight + Canvas.GetTop(CanvaspictureEEG)
        If C = 1 Then
            PathEEG = "..\..\Resources\EEG1\"
        Else
            PathEEG = "..\..\Resources\EEG2\9"
        End If
        If count = 1 Then
            PictureEEG.Source = New BitmapImage(New Uri("..\..\Resources\EEG2\90.jpg", UriKind.Relative))
            YEEG = Canvas.GetLeft(CanvaspictureEEG) + 70
            A = 0
            LineEEG.Stroke = mySolidColorBrush1
            LineEEG.StrokeThickness = 5
        ElseIf (count) / 30 = Int((count) / 30) Then
            canvasligneEEG.Children.Remove(LineEEG)
            Canvas1.Children.Remove(canvasligneEEG)
            YEEG = Canvas.GetLeft(CanvaspictureEEG) + PictureEEG.ActualWidth
            A = 0
            Canvas.SetLeft(LineEEG, YEEG)
            canvasligneEEG.Children.Add(LineEEG)
            Canvas1.Children.Add(canvasligneEEG)
        Else
            canvasligneEEG.Children.Remove(LineEEG)
            Canvas1.Children.Remove(canvasligneEEG)
            A = count - Int((count) / 30) * 30
            YEEG = Canvas.GetLeft(CanvaspictureEEG) + A * ((PictureEEG.ActualWidth - 70) / 30)
            Canvas.SetLeft(LineEEG, YEEG)
            canvasligneEEG.Children.Add(LineEEG)
            Canvas1.Children.Add(canvasligneEEG)
        End If
        PictureEEG.Source = New BitmapImage(New Uri(PathEEG & Int((count - 1) / 30) & ".jpg", UriKind.Relative))
        LineEEG.X1 = 25
        LineEEG.X2 = 25
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
                count = 38
            Case 2
                count = 60
            Case 3
                count = 320
            Case 4
                count = 350
            Case 5
                count = 530
            Case 6
                count = 653
            Case 7
                count = 773
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
            PicList(i).Source = New BitmapImage(New Uri("..\..\Resources\scalp.jpg", UriKind.Relative))
        Next
        B = 2
        Me.Content = Canvas1
        C = 0
        SeuilCoh = 0.3
        Slide = 0
        For k As Integer = 0 To 4
            Canvas.SetLeft(Txtlist(k), Leftpic + PicList(k).ActualWidth)
            Canvas.SetTop(Txtlist(k), TopPic(k))
            Canvas1.Children.Add(Txtlist(k))
        Next
    End Sub
    Public Sub windows1_SizeChanged(sender As Object, e As SizeChangedEventArgs) Handles windows1.SizeChanged
        PicList.Clear()
        listcanpic.Clear()
        PicList.Add(Picture)
        PicList.Add(Picture1)
        PicList.Add(Picture2)
        PicList.Add(Picture3)
        PicList.Add(Picture4)
        listcanpic.Add(Canvaspicture)
        listcanpic.Add(Canvaspicture1)
        listcanpic.Add(Canvaspicture2)
        listcanpic.Add(Canvaspicture3)
        listcanpic.Add(Canvaspicture4)
        TxtboxD.Text = "0.5-4 Hz"
        TxtBoxT.Text = "4-8 Hz"
        TxtBoxA.Text = "8-12 Hz"
        TxtBoxB.Text = "12-30 Hz"
        TxtBoxG.Text = "30-60 Hz"
        Txtlist.Add(TxtboxD)
        Txtlist.Add(TxtBoxT)
        Txtlist.Add(TxtBoxA)
        Txtlist.Add(TxtBoxB)
        Txtlist.Add(TxtBoxG)
        modif()

        CanvaspictureEEG.Children.Remove(PictureEEG)
        Canvas1.Children.Remove(CanvaspictureEEG)
        If C = 1 Then
            Dim uriAdress As New Uri("..\..\Resources\EEG1\", UriKind.Relative)
            PathEEG = uriAdress.ToString
            LineEEG.Y2 = 750 + 150
            Scroll1.Margin = New Thickness(25, 750 + 75, Scroll1.Margin.Right, Scroll1.Margin.Bottom)
            Scroll1.Width = 1278
        Else
            Dim uriAdress2 As New Uri("..\..\Resources\EEG2\9", UriKind.Relative)
            PathEEG = uriAdress2.ToString
            LineEEG.Y2 = 450 + 150
            Scroll1.Margin = New Thickness(25, 450 + 75, Scroll1.Margin.Right, Scroll1.Margin.Bottom)
            Scroll1.Width = 766
        End If
        Canvas.SetLeft(CanvaspictureEEG, 25)
        Canvas.SetTop(CanvaspictureEEG, 75)
        If count = 0 Then
            count = 1
        End If
        PictureEEG.Source = New BitmapImage(New Uri(PathEEG & Int((count - 1) / 30) & ".jpg", UriKind.Relative))
        CanvaspictureEEG.Children.Add(PictureEEG)
        Canvas1.Children.Add(CanvaspictureEEG)
        Me.Content = Canvas1
    End Sub
    Private Sub modif()
        For i As Integer = 0 To 4
            listcanpic(i).Children.Remove(PicList(i))
            Canvas1.Children.Remove(listcanpic(i))
            If windows1.ActualWidth > 1500 Then
                PicList(i).Source = New BitmapImage(New Uri("..\..\Resources\scalp.jpg", UriKind.Relative))
                leftP = 280
                C = 1
                B = 1
                Canvas.SetTop(Panneau, 764 + 180)
            Else
                PicList(i).Source = New BitmapImage(New Uri("..\..\Resources\scalp.jpg", UriKind.Relative))
                leftP = 120
                C = 0
                B = 2
                Canvas.SetTop(Panneau, 764)
            End If
            Leftpic = windows1.ActualWidth - 280
            TopPic.Add(windows1.ActualHeight - (windows1.ActualHeight - 30) + 200 * i)
            Canvas.SetTop(listcanpic(i), TopPic(i))
            Canvas.SetLeft(listcanpic(i), Leftpic)
            listcanpic(i).Children.Add(PicList(i))
            Canvas1.Children.Add(listcanpic(i))
        Next
        For k As Integer = 0 To 4
            Canvas1.Children.Remove(Txtlist(k))
            Canvas.SetLeft(Txtlist(k), Leftpic + PicList(k).ActualWidth)
            Canvas.SetTop(Txtlist(k), TopPic(k))
            Canvas1.Children.Add(Txtlist(k))
        Next
        canvasligneEEG.Children.Remove(LineEEG)
        Canvas1.Children.Remove(canvasligneEEG)
        canvasligneEEG.Children.Add(LineEEG)
        Canvas1.Children.Add(canvasligneEEG)
    End Sub
    Private Sub MTXTBox2_LostFocus(sender As Object, e As RoutedEventArgs) Handles MTXTBox2.LostFocus
        SeuilCoh = Convert.ToDouble(MTXTBox2.Text)
    End Sub
End Class
