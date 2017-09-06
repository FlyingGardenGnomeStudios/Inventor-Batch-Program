
Imports System.Windows.Forms
Imports System.Drawing

Imports Inventor

'Namespace Batch_Program
''' <summary>
''' The class wrapps up Inventor Button creation stuffs and is easy to use.
''' No need to derive. Create an instance using either constructor and assign the Action.
''' </summary>
Public Class InventorButton
#Region "Fields & Properties"

    Private mButtonDef As ButtonDefinition
    Public Property ButtonDef() As ButtonDefinition
        Get
            Return mButtonDef
        End Get
        Set
            mButtonDef = Value
        End Set
    End Property

#End Region

#Region "Constructors"

    ''' <summary>
    ''' The most comprehensive signature.
    ''' </summary>
    Public Sub New(displayName As String, internalName As String, description As String, tooltip As String, standardIcon As Icon, largeIcon As Icon,
            commandType As CommandTypesEnum, buttonDisplayType As ButtonDisplayEnum)
        Create(displayName, internalName, description, tooltip, AddinGlobal.ClassId, standardIcon,
                largeIcon, commandType, buttonDisplayType)
    End Sub

    ''' <summary>
    ''' The signature does not care about Command Type (always editing) and Button Display (always with text).
    ''' </summary>
    Public Sub New(displayName As String, internalName As String, description As String, tooltip As String, standardIcon As Icon, largeIcon As Icon)
        Create(displayName, internalName, description, tooltip, AddinGlobal.ClassId, Nothing,
                Nothing, CommandTypesEnum.kEditMaskCmdType, ButtonDisplayEnum.kAlwaysDisplayText)
    End Sub

    ''' <summary>
    ''' The signature does not care about icons.
    ''' </summary>
    Public Sub New(displayName As String, internalName As String, description As String, tooltip As String, commandType As CommandTypesEnum, buttonDisplayType As ButtonDisplayEnum)
        Create(displayName, internalName, description, tooltip, AddinGlobal.ClassId, Nothing,
                Nothing, commandType, buttonDisplayType)
    End Sub

    ''' <summary>
    ''' This signature only cares about display name and icons.
    ''' </summary>
    ''' <param name="displayName"></param>
    ''' <param name="standardIcon"></param>
    ''' <param name="largeIcon"></param>
    Public Sub New(displayName As String, standardIcon As Icon, largeIcon As Icon)
        Create(displayName, displayName, displayName, displayName, AddinGlobal.ClassId, standardIcon,
                largeIcon, CommandTypesEnum.kEditMaskCmdType, ButtonDisplayEnum.kAlwaysDisplayText)
    End Sub

    ''' <summary>
    ''' The simplest signature, which can be good for prototyping.
    ''' </summary>
    Public Sub New(displayName As String)
        Create(displayName, displayName, displayName, displayName, AddinGlobal.ClassId, Nothing,
                Nothing, CommandTypesEnum.kEditMaskCmdType, ButtonDisplayEnum.kAlwaysDisplayText)
    End Sub

    ''' <summary>
    ''' The helper method for constructors to call to avoid duplicate code.
    ''' </summary>
    Public Sub Create(displayName As String, internalName As String, description As String, tooltip As String, clientId As String, standardIcon As Icon,
            largeIcon As Icon, commandType As CommandTypesEnum, buttonDisplayType As ButtonDisplayEnum)
        If String.IsNullOrEmpty(clientId) Then
            clientId = AddinGlobal.ClassId
        End If

        Dim standardIconIPictureDisp As stdole.IPictureDisp = Nothing
        Dim largeIconIPictureDisp As stdole.IPictureDisp = Nothing
        If standardIcon IsNot Nothing Then
            standardIconIPictureDisp = IconToPicture(standardIcon)
            largeIconIPictureDisp = IconToPicture(largeIcon)
        End If

        mButtonDef = AddinGlobal.InventorApp.CommandManager.ControlDefinitions.AddButtonDefinition(displayName, internalName, commandType, clientId, description, tooltip,
                standardIconIPictureDisp, largeIconIPictureDisp, buttonDisplayType)

        mButtonDef.Enabled = True
        AddHandler mButtonDef.OnExecute, AddressOf ButtonDefinition_OnExecute

        DisplayText = True

        AddinGlobal.ButtonList.Add(Me)
    End Sub

#End Region

#Region "Behavior"

    Public Property DisplayBigIcon() As Boolean
        Get
            Return m_DisplayBigIcon
        End Get
        Set
            m_DisplayBigIcon = Value
        End Set
    End Property
    Private m_DisplayBigIcon As Boolean
    Public Property DisplayText() As Boolean
        Get
            Return m_DisplayText
        End Get
        Set
            m_DisplayText = Value
        End Set
    End Property
    Private m_DisplayText As Boolean
    Public Property InsertBeforeTarget() As Boolean
        Get
            Return m_InsertBeforeTarget
        End Get
        Set
            m_InsertBeforeTarget = Value
        End Set
    End Property
    Private m_InsertBeforeTarget As Boolean

    Public Sub SetBehavior(displayBigIcon__1 As Boolean, displayText__2 As Boolean, insertBeforeTarget__3 As Boolean)
        DisplayBigIcon = displayBigIcon__1
        DisplayText = displayText__2
        InsertBeforeTarget = insertBeforeTarget__3
    End Sub

    Public Sub CopyBehaviorFrom(button As InventorButton)
        Me.DisplayBigIcon = button.DisplayBigIcon
        Me.DisplayText = button.DisplayText
        Me.InsertBeforeTarget = Me.InsertBeforeTarget
    End Sub

#End Region

#Region "Actions"

    ''' <summary>
    ''' The button callback method.
    ''' </summary>
    ''' <param name="context"></param>
    Private Sub ButtonDefinition_OnExecute(context As NameValueMap)
        If Execute IsNot Nothing Then
            Execute()
        Else
            MessageBox.Show("Nothing to execute.")
        End If
    End Sub

    ''' <summary>
    ''' The button action to be assigned from anywhere outside.
    ''' </summary>
    Public Execute As Action

#End Region

#Region "Image Converters"

    Public Shared Function ImageToPicture(image As Image) As stdole.IPictureDisp
        Return ImageConverter.ImageToPicture(image)
    End Function

    Public Shared Function IconToPicture(icon As Icon) As stdole.IPictureDisp
        Return ImageConverter.ImageToPicture(icon.ToBitmap())
    End Function

    Public Shared Function PictureToImage(picture As stdole.IPictureDisp) As Image
        Return ImageConverter.PictureToImage(picture)
    End Function

    Public Shared Function PictureToIcon(picture As stdole.IPictureDisp) As Icon
        Return ImageConverter.PictureToIcon(picture)
    End Function

    Private Class ImageConverter
        Inherits AxHost
        Public Sub New()
            MyBase.New(String.Empty)
        End Sub

        Public Shared Function ImageToPicture(image As Image) As stdole.IPictureDisp
            Return DirectCast(GetIPictureDispFromPicture(image), stdole.IPictureDisp)
        End Function

        Public Shared Function IconToPicture(icon As Icon) As stdole.IPictureDisp
            Return ImageToPicture(icon.ToBitmap())
        End Function

        Public Shared Function PictureToImage(picture As stdole.IPictureDisp) As Image
            Return GetPictureFromIPicture(picture)
        End Function

        Public Shared Function PictureToIcon(picture As stdole.IPictureDisp) As Icon
            Dim bitmap As New Bitmap(PictureToImage(picture))
            Return System.Drawing.Icon.FromHandle(bitmap.GetHicon())
        End Function
    End Class

#End Region

End Class
'End Namespace


