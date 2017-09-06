
#Region "Namespaces"

Imports System.Text
Imports System.Linq
Imports System.Xml
Imports System.Reflection
Imports System.ComponentModel
Imports System.Collections
Imports System.Collections.Generic
Imports System.Windows
Imports System.Windows.Media.Imaging
Imports System.Windows.Forms
Imports System.Drawing
Imports System.IO
Imports System.Runtime.InteropServices
Imports Microsoft.Win32
Imports Inventor
#End Region

'Namespace Batch_Program
Public Class ButtonActions

    Public Shared Sub Button1_Execute()
        ' Dim main As New Test_Batch.Main
        Dim Main As New Main
        'TODO: add code below for the button click callback.
        Main.ShowDialog()
    End Sub
End Class
'End Namespace


