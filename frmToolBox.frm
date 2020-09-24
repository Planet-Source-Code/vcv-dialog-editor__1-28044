VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form ToolBox 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Toolbox"
   ClientHeight    =   4020
   ClientLeft      =   2625
   ClientTop       =   3675
   ClientWidth     =   1560
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   268
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   104
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   630
      Top             =   2100
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   21
      ImageHeight     =   21
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmToolBox.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmToolBox.frx":06A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmToolBox.frx":0D4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmToolBox.frx":13F2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar toolBar 
      Height          =   810
      Left            =   45
      TabIndex        =   0
      Top             =   45
      Width           =   885
      _ExtentX        =   1561
      _ExtentY        =   1429
      ButtonWidth     =   741
      ButtonHeight    =   714
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   1
            Style           =   2
            Value           =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   2
            Style           =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   3
            Style           =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   4
            Style           =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "ToolBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    ToolBox.width = (68 * 15)
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Me.Tag = "" Then Cancel = 1
End Sub


Private Sub Form_Resize()
    toolBar.Move toolBar.Left, toolBar.Top, Me.ScaleWidth - 4, Me.ScaleHeight - 6
End Sub


Private Sub toolBar_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim x As Form
    For Each x In Forms
        If x.Name = "dialogTemplate" Then
            If Button.Index > 1 Then
                x.MousePointer = 2
            Else
                x.MousePointer = 0
            End If
        End If
    Next x
End Sub


