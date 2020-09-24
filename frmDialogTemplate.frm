VERSION 5.00
Begin VB.Form dialogTemplate 
   AutoRedraw      =   -1  'True
   Caption         =   "Dialog"
   ClientHeight    =   2730
   ClientLeft      =   4710
   ClientTop       =   4035
   ClientWidth     =   4710
   DrawMode        =   6  'Mask Pen Not
   DrawWidth       =   3
   FillColor       =   &H8000000D&
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDialogTemplate.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MousePointer    =   1  'Arrow
   ScaleHeight     =   182
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   314
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picResize 
      BorderStyle     =   0  'None
      Height          =   105
      Left            =   2025
      MousePointer    =   8  'Size NW SE
      ScaleHeight     =   7
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   7
      TabIndex        =   3
      Top             =   1140
      Visible         =   0   'False
      Width           =   105
      Begin VB.Shape Shape1 
         BackColor       =   &H8000000D&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H8000000E&
         Height          =   105
         Left            =   0
         Top             =   0
         Width           =   105
      End
   End
   Begin VB.CommandButton btnTemplate 
      Caption         =   "Button"
      Height          =   465
      Index           =   0
      Left            =   330
      TabIndex        =   2
      Top             =   660
      Visible         =   0   'False
      Width           =   990
   End
   Begin VB.PictureBox picBuffer 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   120
      Left            =   0
      ScaleHeight     =   8
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.TextBox txtTemplate 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Index           =   0
      Left            =   765
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   0
      Top             =   240
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.Label lblTemplate 
      Caption         =   "Label"
      Height          =   240
      Index           =   0
      Left            =   2265
      TabIndex        =   4
      Top             =   255
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.Shape shpCreate 
      BorderWidth     =   2
      DrawMode        =   6  'Mask Pen Not
      Height          =   600
      Left            =   3120
      Top             =   1410
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Shape shpMove 
      BorderColor     =   &H008080FF&
      BorderWidth     =   2
      DrawMode        =   6  'Mask Pen Not
      FillColor       =   &H008080FF&
      Height          =   585
      Left            =   3120
      Top             =   780
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Shape shpResize 
      BackColor       =   &H8000000D&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000E&
      Height          =   105
      Left            =   2130
      Shape           =   1  'Square
      Top             =   1140
      Visible         =   0   'False
      Width           =   105
   End
End
Attribute VB_Name = "dialogTemplate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public StartX   As Integer
Public StartY   As Integer
Public EndX     As Integer
Public EndY     As Integer
Public XPos     As Integer
Public YPos     As Integer
Public keyShift As Integer
Public selectedControl  As Object


Public Sub FillProps()
'    FillProperties Me, Dialog_props
End Sub


Public Sub MoveResizeHandle()
    With selectedControl
        If picResize.Visible = True Then picResize.Move .Left + .width, .Top + .height
        If shpResize.Visible = True Then shpResize.Move .Left + .width, .Top + .height
    End With
End Sub


Public Sub SetFoc()
'    Set selectedControl = Me
End Sub

Private Sub btnTemplate_GotFocus(Index As Integer)

    FillProperties btnTemplate(Index), Button_props
    FillControlList btnTemplate(Index)
    
 
End Sub

Private Sub btnTemplate_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
        
    StartX = x
    StartY = Y
    XPos = btnTemplate(Index).Left
    YPos = btnTemplate(Index).Top
    
    Set selectedControl = btnTemplate(Index)
    With selectedControl
        picResize.Move .Left + .width, .Top + .height
        picResize.Visible = True
'        .Tag = Index
        shpMove.Move XPos, YPos, .width, .height
    End With
    
End Sub


Private Sub btnTemplate_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
    Dim yNew As Integer, xNew As Integer
    If Button = 1 Then
        With btnTemplate(Index)
            yNew = GridY(YPos - (Pixels(StartY - Y)))
            xNew = GridX(XPos - (Pixels(StartX - x)))
            
            shpMove.Move xNew, yNew
            If shpMove.Visible = False Then shpMove.Visible = True
            
        End With
    End If
End Sub


Private Sub btnTemplate_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
    With btnTemplate(Index)
        .Move GridX(XPos - (Pixels(StartX - x))), GridY(YPos - (Pixels(StartY - Y)))
        shpMove.Visible = False
    End With
    
    With selectedControl
        picResize.Move .Left + .width, .Top + .height
        picResize.Visible = True
        '.Tag = Index
    End With
    
    
    FillProperties btnTemplate(Index), Button_props
    FillControlList btnTemplate(Index)
End Sub


Private Sub Form_Activate()
    On Error Resume Next
    Set selectedForm = Me
    If ToolBox.toolBar.Buttons.Item(1).Value = tbrPressed Then
        Me.MousePointer = 1
    Else
        Me.MousePointer = 2
    End If
    
    'FillProperties Me, Dialog_props
    FillControlList Me
End Sub

Private Sub Form_Click()
    FillProperties Me, Dialog_props
    
    FillControlList Me
End Sub

Private Sub Form_GotFocus()
    'Set selectedControl = Me
    'FillProperties Me, Dialog_props
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    Dim GridSize2 As Integer
    If useGrid Then
        GridSize2 = GridSize
    Else
        GridSize2 = 1
    End If
    If Shift = 1 Then  'shift
        If KeyCode = 38 Then        'up
            selectedControl.height = selectedControl.height - GridSize2
        ElseIf KeyCode = 40 Then    'down
            selectedControl.height = selectedControl.height + GridSize2
        ElseIf KeyCode = 37 Then    'left
            selectedControl.width = selectedControl.width - GridSize2
        ElseIf KeyCode = 39 Then    'right
            selectedControl.width = selectedControl.width + GridSize2
        End If
        MoveResizeHandle
    ElseIf Shift = 2 Then  'control
        If KeyCode = 38 Then        'up
            selectedControl.Top = selectedControl.Top - GridSize2
        ElseIf KeyCode = 40 Then    'down
            selectedControl.Top = selectedControl.Top + GridSize2
        ElseIf KeyCode = 37 Then    'left
            selectedControl.Left = selectedControl.Left - GridSize2
        ElseIf KeyCode = 39 Then    'right
            selectedControl.Left = selectedControl.Left + GridSize2
        End If
        MoveResizeHandle
    End If
    
        
End Sub


Private Sub Form_Load()
    DrawTheGrid Me
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    StartX = (GridX(CInt(x)))
    StartY = (GridY(CInt(Y)))
    
    Set selectedControl = Me
    picResize.Visible = False
    shpResize.Visible = False
    MoveResizeHandle
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    EndX = GridX(CInt(x))
    EndY = GridY(CInt(Y))
    
    If Button = 1 And Me.MousePointer = 2 Then
        On Error Resume Next
        shpCreate.Move StartX, StartY, EndX - StartX, EndY - StartY
        If shpCreate.Visible = False Then shpCreate.Visible = True
    End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    On Error Resume Next
    
    shpCreate.Visible = False
    If Me.MousePointer = 2 Then
    
        With ToolBox.toolBar
        Select Case tbrPressed
        Case .Buttons(2).Value
            If txtTemplate(0).Tag = "" Then txtTemplate(0).Tag = 0
            txtTemplate(0).Tag = txtTemplate(0).Tag + 1
            Load txtTemplate(txtTemplate(0).Tag)
            'object(object(0).Tag).Move StartX, StartY, EndX - StartX, EndY - StartY
            txtTemplate(txtTemplate(0).Tag).Move StartX, StartY, EndX - StartX, EndY - StartY
            txtTemplate(txtTemplate(0).Tag).Visible = True
            txtTemplate(txtTemplate(0).Tag).Tag = "Text" & txtTemplate(0).Tag
            txtTemplate(txtTemplate(0).Tag).SetFocus
        Case .Buttons(3).Value
            If btnTemplate(0).Tag = "" Then btnTemplate(0).Tag = 0
            btnTemplate(0).Tag = btnTemplate(0).Tag + 1
            Load btnTemplate(btnTemplate(0).Tag)
            'object(object(0).Tag).Move StartX, StartY, EndX - StartX, EndY - StartY
            btnTemplate(btnTemplate(0).Tag).Move StartX, StartY, EndX - StartX, EndY - StartY
            btnTemplate(btnTemplate(0).Tag).Visible = True
            btnTemplate(btnTemplate(0).Tag).Tag = "Button" & btnTemplate(0).Tag
            btnTemplate(btnTemplate(0).Tag).SetFocus
        Case .Buttons(4).Value
            If lblTemplate(0).Tag = "" Then lblTemplate(0).Tag = 0
            lblTemplate(0).Tag = lblTemplate(0).Tag + 1
            Load lblTemplate(lblTemplate(0).Tag)
            'object(object(0).Tag).Move StartX, StartY, EndX - StartX, EndY - StartY
            lblTemplate(lblTemplate(0).Tag).Move StartX, StartY, EndX - StartX, EndY - StartY
            lblTemplate(lblTemplate(0).Tag).Visible = True
            lblTemplate(lblTemplate(0).Tag).Tag = "Label" & lblTemplate(0).Tag
        End Select
        End With
        
    End If
    
    ToolBox.toolBar.Buttons.Item(1).Value = tbrPressed
    Me.MousePointer = 1
End Sub



Private Sub Form_Resize()
    DrawTheGrid Me
End Sub


Private Sub lblTemplate_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
    StartX = x
    StartY = Y
    XPos = lblTemplate(Index).Left
    YPos = lblTemplate(Index).Top
    
    Set selectedControl = lblTemplate(Index)
    With selectedControl
        picResize.Move .Left + .width, .Top + .height
        picResize.Visible = True
        
        shpMove.Move XPos, YPos, .width, .height
    End With
    
End Sub


Private Sub lblTemplate_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
    Dim yNew As Integer, xNew As Integer
    If Button = 1 Then
        With lblTemplate(Index)
            yNew = GridY(YPos - (Pixels(StartY - Y)))
            xNew = GridX(XPos - (Pixels(StartX - x)))
            
            shpMove.Move xNew, yNew
            If shpMove.Visible = False Then shpMove.Visible = True
            
        End With
    End If
End Sub


Private Sub lblTemplate_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
    With lblTemplate(Index)
        .Move GridX(XPos - (Pixels(StartX - x))), GridY(YPos - (Pixels(StartY - Y)))
        shpMove.Visible = False
    End With
    
    With selectedControl
        picResize.Move .Left + .width, .Top + .height
        picResize.Visible = True
        '.Tag = Index
    End With
    
    FillProperties lblTemplate(Index), Label_props
    FillControlList lblTemplate(Index)
End Sub


Private Sub picResize_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = 1 Then
        
        StartX = x
        StartY = Y
        XPos = selectedControl.width
        YPos = selectedControl.height
        
        picResize.Visible = False
        shpResize.Visible = True
    End If
    FillControlList selectedControl
    
End Sub


Private Sub picResize_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = 1 Then
        With selectedControl
            On Error Resume Next
            .width = GridX(XPos - (StartX - x))
            .height = GridY(YPos - (StartY - Y))
            MoveResizeHandle
        End With
    End If
End Sub


Private Sub picResize_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = 1 Then
        FillControlList selectedControl
        shpResize.Visible = False
        picResize.Visible = True
        With selectedControl
            MoveResizeHandle
        End With
    End If
    
End Sub

Private Sub txtTemplate_GotFocus(Index As Integer)
        
    FillProperties txtTemplate(Index), Edit_props
    FillControlList txtTemplate(Index)
End Sub

Private Sub txtTemplate_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
    StartX = x
    StartY = Y
    XPos = txtTemplate(Index).Left
    YPos = txtTemplate(Index).Top
    
    Set selectedControl = txtTemplate(Index)
    With selectedControl
        picResize.Move .Left + .width, .Top + .height
        picResize.Visible = True
        '.Tag = Index
        shpMove.Move XPos, YPos, .width, .height
    End With
    
End Sub


Private Sub txtTemplate_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
    Dim yNew As Integer, xNew As Integer
    If Button = 1 Then
        With txtTemplate(Index)
            yNew = GridY(YPos - (Pixels(StartY - Y)))
            xNew = GridX(XPos - (Pixels(StartX - x)))
            
            shpMove.Move xNew, yNew
            If shpMove.Visible = False Then shpMove.Visible = True
        End With
    End If
End Sub


Private Sub txtTemplate_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
    With txtTemplate(Index)
        .Move GridX(XPos - (Pixels(StartX - x))), GridY(YPos - (Pixels(StartY - Y)))
        shpMove.Visible = False
    End With
    
    With selectedControl
        picResize.Move .Left + .width, .Top + .height
        picResize.Visible = True
       ' .Tag = Index
    End With
    
    
    FillProperties txtTemplate(Index), Edit_props
    FillControlList txtTemplate(Index)
End Sub



