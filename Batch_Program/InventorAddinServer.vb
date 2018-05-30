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
<GuidAttribute("97c796fe-4ecd-47ca-9736-c4ba1ba68a1b")>
Public Class InventorAddinServer
    Implements Inventor.ApplicationAddInServer
    Public Sub New()
    End Sub

#Region "Implementations of the Interface Members"

    ''' <summary>
    ''' Do initializations in it such as caching the application object, registering event handlers, and adding ribbon buttons.
    ''' </summary>
    ''' <param name="siteObj">The entry object for the addin.</param>
    ''' <param name="loaded1stTime">Indicating whether the addin is loaded for the 1st time.</param>
    Public Sub Activate(siteObj As Inventor.ApplicationAddInSite, loaded1stTime As Boolean) Implements ApplicationAddInServer.Activate
        AddinGlobal.InventorApp = siteObj.Application

        Try
            AddinGlobal.GetAddinClassId(Me.GetType())

            Dim icon1 As New Icon(Assembly.GetExecutingAssembly().GetManifestResourceStream("addin.ico"))
            'Change it if necessary but make sure it's embedded.
            'Dim button1 As New InventorButton("Batch Program", "InventorAddinServer.Button_" & Guid.NewGuid().ToString(), "", "", icon1, icon1, _
            'CommandTypesEnum.kShapeEditCmdType, ButtonDisplayEnum.kDisplayTextInLearningMode)
            '
            'button1.Execute = AddressOf ButtonActions.Button1_Execute
            Dim button1 As New InventorButton("Run Batch", "InventorAddinServer.Button_" & Guid.NewGuid().ToString(),
                                              "Modify attributes of related documents open in Inventor", "Modify attributes of related documents open in Inventor", icon1, icon1,
                                                CommandTypesEnum.kShapeEditCmdType, ButtonDisplayEnum.kDisplayTextInLearningMode)
            button1.SetBehavior(True, True, True)
            button1.Execute = AddressOf ButtonActions.Button1_Execute

            If loaded1stTime = True Then
                Dim uiMan As UserInterfaceManager = AddinGlobal.InventorApp.UserInterfaceManager
                If uiMan.InterfaceStyle = InterfaceStyleEnum.kRibbonInterface Then
                    'kClassicInterface support can be added if necessary.
                    Dim ribbon As Inventor.Ribbon = uiMan.Ribbons("Assembly")
                    Dim tab As RibbonTab = ribbon.RibbonTabs("id_TabTools")
                    'Change it if necessary.
                    AddinGlobal.RibbonPanelId = "{b9d9b9eb-d9b9-46e5-8551-a87982cdd3cf}"
                    AddinGlobal.RibbonPanel = tab.RibbonPanels.Add("Batch Program", "InventorAddinServer.RibbonPanel_" & Guid.NewGuid().ToString(), AddinGlobal.RibbonPanelId, String.Empty, True)

                    Dim cmdCtrls As CommandControls = AddinGlobal.RibbonPanel.CommandControls
                    cmdCtrls.AddButton(button1.ButtonDef, button1.DisplayBigIcon, button1.DisplayText, "", button1.InsertBeforeTarget)
                End If
            End If

            'Change it if necessary but make sure it's embedded.
            Dim button3 As New InventorButton("Run Batch", "InventorAddinServer.Button_" & Guid.NewGuid().ToString(),
                                              "Modify attributes of related documents open in Inventor", "Modify attributes of related documents open in Inventor", icon1, icon1,
                                                CommandTypesEnum.kShapeEditCmdType, ButtonDisplayEnum.kDisplayTextInLearningMode)
            button1.SetBehavior(True, True, True)
            button1.Execute = AddressOf ButtonActions.Button1_Execute

            If loaded1stTime = True Then
                Dim uiMan As UserInterfaceManager = AddinGlobal.InventorApp.UserInterfaceManager
                If uiMan.InterfaceStyle = InterfaceStyleEnum.kRibbonInterface Then
                    'kClassicInterface support can be added if necessary.
                    Dim ribbon As Inventor.Ribbon = uiMan.Ribbons("Part")
                    Dim tab As RibbonTab = ribbon.RibbonTabs("id_TabTools")
                    'Change it if necessary.
                    AddinGlobal.RibbonPanelId = "{b9d9b9eb-d9b9-46e5-8551-a87982cdd3cf}"
                    AddinGlobal.RibbonPanel = tab.RibbonPanels.Add("Batch Program", "InventorAddinServer.RibbonPanel_" & Guid.NewGuid().ToString(), AddinGlobal.RibbonPanelId, String.Empty, True)

                    Dim cmdCtrls As CommandControls = AddinGlobal.RibbonPanel.CommandControls
                    cmdCtrls.AddButton(button1.ButtonDef, button1.DisplayBigIcon, button1.DisplayText, "", button1.InsertBeforeTarget)
                End If
            End If

            'Change it if necessary but make sure it's embedded.
            Dim button2 As New InventorButton("Run Batch", "InventorAddinServer.Button_" & Guid.NewGuid().ToString(),
                                              "Modify attributes of related documents open in Inventor", "Modify attributes of related documents open in Inventor", icon1, icon1,
                                                CommandTypesEnum.kShapeEditCmdType, ButtonDisplayEnum.kDisplayTextInLearningMode)
            button1.SetBehavior(True, True, True)
            button1.Execute = AddressOf ButtonActions.Button1_Execute

            If loaded1stTime = True Then
                Dim uiMan As UserInterfaceManager = AddinGlobal.InventorApp.UserInterfaceManager
                If uiMan.InterfaceStyle = InterfaceStyleEnum.kRibbonInterface Then
                    'kClassicInterface support can be added if necessary.
                    Dim ribbon As Inventor.Ribbon = uiMan.Ribbons("Drawing")
                    Dim tab As RibbonTab = ribbon.RibbonTabs("id_TabTools")
                    'Change it if necessary.
                    AddinGlobal.RibbonPanelId = "{b9d9b9eb-d9b9-46e5-8551-a87982cdd3cf}"
                    AddinGlobal.RibbonPanel = tab.RibbonPanels.Add("Batch Program", "InventorAddinServer.RibbonPanel_" & Guid.NewGuid().ToString(), AddinGlobal.RibbonPanelId, String.Empty, True)

                    Dim cmdCtrls As CommandControls = AddinGlobal.RibbonPanel.CommandControls
                    cmdCtrls.AddButton(button1.ButtonDef, button1.DisplayBigIcon, button1.DisplayText, "", button1.InsertBeforeTarget)
                End If
            End If
        Catch e As Exception
            MessageBox.Show(e.ToString())
        End Try

        ' TODO: Add more initialization code below.

    End Sub

    ''' <summary>
    ''' Do cleanups in it such as releasing COM objects or forcing the GC to Collect when necessary.
    ''' </summary>
    Public Sub Deactivate() Implements ApplicationAddInServer.Deactivate
        ' Add more cleanup work below if necessary, e.g. remove even handlers, flush and close log files, etc.


        ' Release COM objects
        '

        ' Add more FinalReleaseComObject() calls for COM objects you know, e.g.
        'if (comObj != null) Marshal.FinalReleaseComObject(comObj);

        ' Release the COM objects maintained by InventorNetAddinWizard.
        For Each button As InventorButton In AddinGlobal.ButtonList
            If button.ButtonDef IsNot Nothing Then
                Marshal.FinalReleaseComObject(button.ButtonDef)
            End If
        Next
        If AddinGlobal.RibbonPanel IsNot Nothing Then
            Marshal.FinalReleaseComObject(AddinGlobal.RibbonPanel)
        End If
        If AddinGlobal.InventorApp IsNot Nothing Then
            Marshal.FinalReleaseComObject(AddinGlobal.InventorApp)
        End If
    End Sub

    ''' <summary>
    ''' Deprecated. Use the ControlDefinition instead to execute commands.
    ''' </summary>
    ''' <param name="commandID"></param>
    Public Sub ExecuteCommand(commandID As Integer) Implements ApplicationAddInServer.ExecuteCommand
    End Sub

    ''' <summary>
    ''' Implement it if wanting to expose your own automation interface. 
    ''' </summary>
    Public ReadOnly Property Automation() As Object Implements ApplicationAddInServer.Automation
        Get
            Return Nothing
        End Get
    End Property

#End Region


End Class
'End Namespace
