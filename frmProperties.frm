VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Properties 
   Caption         =   "Properties"
   ClientHeight    =   4575
   ClientLeft      =   11955
   ClientTop       =   3675
   ClientWidth     =   3450
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HasDC           =   0   'False
   Icon            =   "frmProperties.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4575
   ScaleWidth      =   3450
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cmbControls 
      Height          =   315
      Left            =   0
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   15
      Width           =   3450
   End
   Begin VB.PictureBox picTip 
      BackColor       =   &H80000018&
      Height          =   990
      Left            =   0
      ScaleHeight     =   930
      ScaleWidth      =   3390
      TabIndex        =   4
      Top             =   3585
      Width           =   3450
      Begin VB.Label lblPropDesc 
         BackStyle       =   0  'Transparent
         Caption         =   "This is where a description of the selected property is placed.  It contains information on what the property does."
         Height          =   630
         Left            =   105
         TabIndex        =   6
         Top             =   270
         Width           =   3240
      End
      Begin VB.Label lblPropName 
         BackStyle       =   0  'Transparent
         Caption         =   "Property Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   60
         TabIndex        =   5
         Top             =   45
         Width           =   3345
      End
   End
   Begin MSComctlLib.ListView lvProperties 
      CausesValidation=   0   'False
      Height          =   2865
      Left            =   -15
      TabIndex        =   3
      Top             =   345
      Width           =   3480
      _ExtentX        =   6138
      _ExtentY        =   5054
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      PictureAlignment=   5
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Name"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   1
         Text            =   "Value"
         Object.Width           =   3087
      EndProperty
      Picture         =   "frmProperties.frx":038A
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "r"
      BeginProperty Font 
         Name            =   "Marlett"
         Size            =   9.75
         Charset         =   2
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3135
      TabIndex        =   2
      Top             =   3225
      Width           =   315
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "a"
      BeginProperty Font 
         Name            =   "Marlett"
         Size            =   9.75
         Charset         =   2
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2805
      TabIndex        =   1
      Top             =   3225
      Width           =   315
   End
   Begin VB.TextBox txtSetProperty 
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Top             =   3225
      Width           =   2775
   End
   Begin VB.ComboBox cmbSetProperty 
      Height          =   315
      Left            =   0
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   3225
      Visible         =   0   'False
      Width           =   2790
   End
End
Attribute VB_Name = "Properties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub HideAllPropTypes()
    cmbSetProperty.Visible = False
    txtSetProperty.Visible = False
End Sub


Private Sub lstProperties_Click()

End Sub


Private Sub cmbControls_Click()
    On Error Resume Next
    Dim x As Control
    If cmbControls.List(cmbControls.ListIndex) = "Dialog" Then
        CallByName selectedForm, "FillProps", VbMethod
        CallByName selectedForm, "SetFoc", VbMethod
        'selectedForm.SetFocus
        Exit Sub
    End If
    For Each x In selectedForm
        If x.Tag = cmbControls.List(cmbControls.ListIndex) Then
            selectedControl = x
            'CallByName selectedControl, "Click", VbMethod
            CallByName selectedControl, "SetFocus", VbMethod
            Exit Sub
        End If
    Next x
End Sub



Private Sub cmbSetProperty_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call cmdOk_Click
    End If
End Sub


Private Sub cmdCancel_Click()
    lvProperties.SetFocus
    txtSetProperty.Text = ""
End Sub

Private Sub cmdOk_Click()
    On Error GoTo errorhand
    
    Select Case True
        Case txtSetProperty.Visible
            CallByName selectedForm.selectedControl, RealPropName(lblPropName), VbLet, txtSetProperty
        Case cmbSetProperty.Visible
            If InStr(cmbSetProperty.List(cmbSetProperty.ListIndex), "-") Then
                CallByName selectedForm.selectedControl, RealPropName(lblPropName), VbLet, Left(cmbSetProperty.List(cmbSetProperty.ListIndex), InStr(cmbSetProperty.List(cmbSetProperty.ListIndex), " - ") - 1)
            Else
                CallByName selectedForm.selectedControl, RealPropName(lblPropName), VbLet, cmbSetProperty.List(cmbSetProperty.ListIndex)
            End If
    End Select
    
    lvProperties.SetFocus
    lvProperties.ListItems.Item(lvProperties.SelectedItem.Index).SubItems(1) = CallByName(selectedForm.selectedControl, RealPropName(lblPropName), VbGet)
    txtSetProperty.Text = ""
    
    Exit Sub
errorhand:
    MsgBox "" & Error, vbCritical, "Error setting property"
    
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If Me.WindowState = vbMinimized Then Exit Sub
    
    'widths / left
    lvProperties.width = Me.width - 100
    cmbControls.width = Me.width - 100
    lvProperties.ColumnHeaders.Item(2).width = Me.width - 2050
    txtSetProperty.width = Me.width - 800
    cmbSetProperty.width = Me.width - 800
    picTip.width = Me.width - 120
    lblPropDesc.width = picTip.width - 220
    cmdOk.Left = txtSetProperty.Left + txtSetProperty.width + 30
    cmdCancel.Left = cmdOk.Left + cmdOk.width + 15
    
    'heights / top
    lvProperties.height = Me.height - 1770 - 345
    txtSetProperty.Top = lvProperties.height + lvProperties.Top + 10
    cmbSetProperty.Top = txtSetProperty.Top
    picTip.Top = txtSetProperty.Top + txtSetProperty.height + 30
    
    cmdOk.Top = txtSetProperty.Top
    cmdCancel.Top = cmdOk.Top
    
End Sub

Private Sub lvProperties_Click()

    Dim propType As String, i As Integer
    
    On Error Resume Next
    lblPropName.Caption = lvProperties.SelectedItem.Text
    lblPropDesc = GetPropertyTip(lblPropName.Caption)
    
    propType = GetPropertyType(lblPropName.Caption)
    HideAllPropTypes
    Select Case propType
        Case "TEXT"
            txtSetProperty.Visible = True
            txtSetProperty.Text = CallByName(selectedForm.selectedControl, RealPropName(lblPropName), VbGet)
            txtSetProperty.Tag = "TEXT"
        Case "BOOL"
            cmbSetProperty.Visible = True
            cmbSetProperty.Clear
            cmbSetProperty.AddItem "True"
            cmbSetProperty.AddItem "False"
            If CallByName(selectedForm.selectedControl, lblPropName, VbGet) = "True" Then
                cmbSetProperty.ListIndex = 0
            Else
                cmbSetProperty.ListIndex = 1
            End If
        Case "INT"
            txtSetProperty.Visible = True
            txtSetProperty.Text = CallByName(selectedForm.selectedControl, lblPropName, VbGet)
            txtSetProperty.Tag = "INT"
        Case "ENUM"
            cmbSetProperty.Visible = True
            cmbSetProperty.Clear
            For i = 1 To GetPropertyENums(lblPropName)
                cmbSetProperty.AddItem GetPropertyENum(lblPropName, i)
                If GetPropertyENum(lblPropName, i) Like CallByName(selectedForm.selectedControl, RealPropName(lblPropName), VbGet) & " - *" Then
                    cmbSetProperty.ListIndex = i - 1
                End If
            Next i
            
           
    End Select
        
End Sub


Private Sub lvProperties_DblClick()
    Select Case Visible
        Case txtSetProperty.Visible
            txtSetProperty.SetFocus
        Case cmbSetProperty.Visible
            cmbSetProperty.SetFocus
    End Select
End Sub


Private Sub lvProperties_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    MsgBox "eh"
End Sub

Private Sub lvProperties_KeyPress(KeyAscii As Integer)
    Dim propType As String, i As Integer
    propType = GetPropertyType(lblPropName.Caption)
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        Select Case propType
            Case "TEXT"
                txtSetProperty.Visible = True
                txtSetProperty.Text = CallByName(selectedForm.selectedControl, lblPropName, VbGet)
                txtSetProperty.Tag = "TEXT"
                txtSetProperty.SetFocus
            Case "BOOL"
                cmbSetProperty.Visible = True
                cmbSetProperty.Clear
                cmbSetProperty.AddItem "True"
                cmbSetProperty.AddItem "False"
                If CallByName(selectedForm.selectedControl, lblPropName, VbGet) = "True" Then
                    cmbSetProperty.ListIndex = 0
                Else
                    cmbSetProperty.ListIndex = 1
                End If
                cmbSetProperty.SetFocus
            Case "INT"
                txtSetProperty.Visible = True
                txtSetProperty.Text = CallByName(selectedForm.selectedControl, lblPropName, VbGet)
                txtSetProperty.Tag = "INT"
                txtSetProperty.SetFocus
            Case "ENUM"
                cmbSetProperty.Visible = True
                cmbSetProperty.Clear
                For i = 1 To GetPropertyENums(lblPropName)
                    cmbSetProperty.AddItem GetPropertyENum(lblPropName, i)
                    If GetPropertyENum(lblPropName, i) Like CallByName(selectedForm.selectedControl, lblPropName, VbGet) & " - *" Then
                        cmbSetProperty.ListIndex = i - 1
                    End If
                Next i
                cmbSetProperty.SetFocus
        End Select
    End If
End Sub

Private Sub txtSetProperty_GotFocus()
    With txtSetProperty
        .SelStart = 0
        .SelLength = Len(txtSetProperty)
    End With
End Sub

Private Sub txtSetProperty_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call cmdOk_Click
    End If
    
    If txtSetProperty.Tag = "INT" Then
        If IsNumeric(Chr(KeyAscii)) = False Then
            KeyAscii = 0
        End If
    End If
End Sub


