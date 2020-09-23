VERSION 5.00
Begin {AC0714F6-3D04-11D1-AE7D-00A0C90F26F4} Connect 
   ClientHeight    =   7605
   ClientLeft      =   1740
   ClientTop       =   1545
   ClientWidth     =   11235
   _ExtentX        =   19817
   _ExtentY        =   13414
   _Version        =   393216
   Description     =   "This add-in will restore the VB IDE to the previous windowstate it was in before a project is run"
   DisplayName     =   "Ed's Restore IDE"
   AppName         =   "Visual Basic"
   AppVer          =   "Visual Basic 6.0"
   LoadName        =   "Startup"
   LoadBehavior    =   1
   RegLocation     =   "HKEY_CURRENT_USER\Software\Microsoft\Visual Basic\6.0"
   CmdLineSafe     =   -1  'True
   CmdLineSupport  =   -1  'True
End
Attribute VB_Name = "Connect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IDTExtensibility
Public VBInstance           As VBIDE.VBE 'our IDE object
Private WithEvents VBBuild  As VBIDE.VBBuildEvents 'this detects run, stop and break events
Attribute VBBuild.VB_VarHelpID = -1
Dim PrevState               As vbext_VBAMode 'stores windowstate from before project was run

Private Sub AddinInstance_OnConnection(ByVal Application As Object, ByVal ConnectMode As AddInDesignerObjects.ext_ConnectMode, ByVal AddInInst As Object, custom() As Variant)

    On Error GoTo Error_Handler
    
    'save the vb instance
    Set VBInstance = Application
    
    'start the initialisation process
    IDTExtensibility_OnStartupComplete custom()
    
    'setting PrevState means that the code that
    'restores the IDE knows that it is not necessary
    'because the project has not been run yet
    PrevState = -1
    
    Exit Sub
    
Error_Handler:
    
    'simple error handler
    MsgBox "Error in AddinInstance_OnConnection" & Chr(13) & Err.Number & ": " & Err.Description
    
End Sub

Private Sub IDTExtensibility_OnStartupComplete(custom() As Variant)

    'define the hidden Events2 object
    Dim VBEvents As Events2

    'initialise the VBEvents object
    Set VBEvents = VBInstance.Events
    'now initiliase the VBBuild object
    Set VBBuild = VBEvents.VBBuildEvents
    
End Sub

Private Sub VBBuild_EnterRunMode()

    Dim lPane   As Long
    
    'save the current windowstate of the IDE in a variable
    PrevState = VBInstance.MainWindow.WindowState
        
End Sub

Private Sub VBBuild_EnterDesignMode()

    Dim lPane   As Long
    
    'if PrevState is -1, then it means that the
    'project hasn't been run yet, this prevents
    'wierd things happening when you launch the
    'IDE
    If PrevState <> -1 Then
        'first we have to minimize it
        VBInstance.MainWindow.WindowState = vbext_ws_Minimize
        'then we can restore it to the
        'state that was saved earlier
        VBInstance.MainWindow.WindowState = PrevState
        'now make sure it has focus
        VBInstance.MainWindow.SetFocus
    End If
    
End Sub


'The following events are required for the project to work
'even though there is no code in them due to the nature of
'the Implements keyword
Private Sub IDTExtensibility_OnAddInsUpdate(custom() As Variant)

End Sub

Private Sub IDTExtensibility_OnDisconnection(ByVal RemoveMode As VBIDE.vbext_DisconnectMode, custom() As Variant)

End Sub

Private Sub IDTExtensibility_OnConnection(ByVal VBInst As Object, ByVal ConnectMode As VBIDE.vbext_ConnectMode, ByVal AddInInst As VBIDE.AddIn, custom() As Variant)

End Sub

