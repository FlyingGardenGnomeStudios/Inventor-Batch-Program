Imports System.Runtime.InteropServices
Public Class Trial
    Dim _invApp As Inventor.Application
    Dim Main As Main
    Public Sub New()
        ' This call is required by the designer.
        InitializeComponent()
        _invApp = Marshal.GetActiveObject("Inventor.Application")

        ' Add any initialization after the InitializeComponent() call.  

    End Sub


    Public Function PopMain(CalledFunction As Main)
        Main = CalledFunction
        Return Nothing
    End Function

End Class