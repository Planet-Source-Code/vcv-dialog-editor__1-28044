VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form mainMenu 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Dialog Editor"
   ClientHeight    =   345
   ClientLeft      =   2625
   ClientTop       =   2880
   ClientWidth     =   12780
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   23
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   852
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3600
      Top             =   -75
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":038A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":049C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":05AE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar toolbar 
      Height          =   330
      Left            =   30
      TabIndex        =   0
      Top             =   0
      Width           =   8040
      _ExtentX        =   14182
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   3
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnu_File 
      Caption         =   "&File"
      Begin VB.Menu mnu_File_NewDialog 
         Caption         =   "&New Dialog"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnu_File_OpenDialog 
         Caption         =   "&Open Dialog"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnu_File_Save 
         Caption         =   "&Save Dialog"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnu_File_LB01 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_File_Exit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnu_options 
      Caption         =   "&Options"
      Begin VB.Menu mnu_Op_ShowGrid 
         Caption         =   "&Show Grid"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnu_Op_Snap 
         Caption         =   "S&nap to Grid"
         Checked         =   -1  'True
      End
   End
End
Attribute VB_Name = "mainMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Unload ToolBox
    Unload Me
    End
End Sub

Private Sub Form_Resize()
    ToolBox.Show
End Sub


Private Sub mnu_File_Exit_Click()
    Unload Me
End Sub

Private Sub mnu_File_NewDialog_Click()
    Dim newDialog As New dialogTemplate
    newDialog.Show
End Sub


Private Sub mnu_Op_ShowGrid_Click()
    With mnu_Op_ShowGrid
        .Checked = Not .Checked
        ShowGrid = .Checked
    End With
    
    Dim x As Form
    For Each x In Forms
        If x.Name = "dialogTemplate" Then DrawTheGrid x
    Next
End Sub


Private Sub mnu_Op_Snap_Click()
    With mnu_Op_Snap
        .Checked = Not .Checked
        useGrid = .Checked
    End With
End Sub


Private Sub toolBar_ButtonClick(ByVal Button As MSComctlLib.Button)
    
    Select Case Button.Index
        Case 1
            Call mnu_File_NewDialog_Click
    End Select
            
End Sub


