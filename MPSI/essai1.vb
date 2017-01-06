Imports Microsoft.VisualBasic
Imports System
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Media
Imports System.Windows.Shapes
Class MainWindow

    Private Sub Bouton1_Click(sender As Object, e As RoutedEventArgs) Handles Bouton1.Click
        Dim myEllipse As New Ellipse()
        'Create a Canvas as the root Panel
        Dim myParentCanvas As New Canvas()
        myParentCanvas.Width = 400
        myParentCanvas.Height = 400

        ' Define child Canvas elements
        Dim myCanvas1 As New Canvas()
        myCanvas1.Background = Brushes.Red
        myCanvas1.Height = 100
        myCanvas1.Width = 100
        Canvas.SetTop(myCanvas1, 0)
        Canvas.SetLeft(myCanvas1, 0)

        Dim myCanvas2 As New Canvas()
        myCanvas2.Background = Brushes.Green
        myCanvas2.Height = 100
        myCanvas2.Width = 100
        Canvas.SetTop(myCanvas2, 100)
        Canvas.SetLeft(myCanvas2, 100)

        Dim myCanvas3 As New Canvas()
        Dim mySolidColorBrush As New SolidColorBrush()
        ' Describes the brush's color using RGB values. 
        ' Each value has a range of 0-255.
        mySolidColorBrush.Color = Color.FromArgb(255, 255, 255, 0)
        myEllipse.Fill = mySolidColorBrush
        myEllipse.StrokeThickness = 5
        ' Set the width and height of the Ellipse.
        myEllipse.Width = 20
        myEllipse.Height = 20
        myCanvas1.Children.Add(myEllipse)
        myCanvas3.Height = 100
        myCanvas3.Width = 100
        Canvas.SetTop(myCanvas3, 50)
        Canvas.SetLeft(myCanvas3, 50)

        Dim myCanvas4 As New Canvas()
        Canvas.SetTop(myCanvas3, 10)
        Canvas.SetLeft(myCanvas3, 10)
        ' Add child elements to the Canvas' Children collection
        myParentCanvas.Children.Add(myCanvas1)
        myParentCanvas.Children.Add(myCanvas2)
        myParentCanvas.Children.Add(myCanvas3)
        myParentCanvas.Children.Add(myCanvas4)
        ' Add the parent Canvas as the Content of the Window Object
        Me.Content = myParentCanvas
    End Sub
End Class
