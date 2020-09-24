Attribute VB_Name = "modFunctions"
Public selectedForm As Object

Public useGrid  As Boolean
Public GridSize As Integer
Public ShowGrid As Boolean
Public gridColor As Long

Public Declare Function SetPixelV Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal Y As Long, ByVal crColor As Long) As Long

'# Property Types
Public Type Property
    Name    As String
    Type    As String
    ENums   As Integer
    ENumItems() As String
    Help    As String
End Type

Public PropertiesA()    As Property
Public PropCount       As Integer
'# End properties

'# Property Info
Public Const Label_props = "Alignment Text,Appearance,AutoSize,BackColor,BackStyle,BorderStyle,Caption,Enabled" & _
                           "Font,ForeColor,Height,Left,MousePointer,Top,Visible,Width,WordWrap"
Public Const Dialog_props = "Appearance,BackColor,BorderStyle Form,Caption,Enabled,Font,ForeColor,Height,ID,Left," & _
                            "MousePointer,Picture,Top,Visible,Width,WindowState"
Public Const Button_props = "Backcolor,Cancel,Caption,Default,Enabled,Font,Height,ID,Left,MousePointer,Picture,Style," & _
                            "Top,Visible,Width"
Public Const Edit_props = "Alignment Text,Appearance,BackColor,BorderStyle,Enabled,Font,ForeColor,Height,ID,Locked," & _
                          "MaxLength,MousePointer,MultiLine,PasswordChar,ScrollBars,Text,Top,Visible,Width"
Public Const Props = _
    "Alignment Text:ENUM:0 - Left justified,1 - Right justified,2 - Center justified:Sets the alignment of the text in the control" & vbCrLf & _
    "Appearance:ENUM:0 - Flat,1 - 3D:Sets the appearance of the control." & vbCrLf & "AutoSize:BOOL::Sets whether or not the static label is resized upon caption changed." & vbCrLf & _
    "BackStyle:ENUM:0 - Transparent,1 - Opaque:Specifies whether or not the background is transparent." & vbCrLf & _
    "BackColor:COLOR::Sets the background color of the control." & vbCrLf & "BorderStyle:ENUM:0 - None,1 - Fixed Single:Sets the style of the control border." & vbCrLf & "BorderStyle Form:ENUM:0 - None,1 - Fixed Single,2 - Sizable,3 - Fixed Dialog,4 - Fixed ToolWindow,5 - Sizable ToolWindow:Sets the style of the dialog border." & vbCrLf & _
    "BorderStyle:ENUM:0 - None,1 - Fixed Single:Sets the style of the control border." & vbCrLf & _
    "Cancel:BOOL::Specifies if the selected control is the 'cancel control' on the dialog." & vbCrLf & _
    "Caption:TEXT::Sets the text displayed in the control." & vbCrLf & _
    "Default:BOOL::Specifies if the selected control is the 'default control' on the dialog." & vbCrLf & _
    "Enabled:BOOL::Sets whether the control is enabled or not upon loading." & vbCrLf & _
    "Font:FONT::Sets the font used to draw the caption/text of the control." & vbCrLf & _
    "ForeColor:COLOR::Sets the foreground color of the control." & vbCrLf & "Height:INT::Sets the height of the control in twips (15 twips = 1 pixel)" & vbCrLf & _
    "ID:TEXT::The ID/Name of the control to distinguish it from other controls of the same type" & vbCrLf & _
    "Left:INT::Sets the Left coordinate of the control in twips (15 twips = 1 pixel)" & vbCrLf & _
    "Locked:BOOL::Specifies whether the control is read-only or not." & vbCrLf & _
    "MaxLength:INT::Sets the maximum length of text that can be placed into the control." & vbCrLf & _
    "MousePointer:ENUM:0 - Default,1 - Arrow,2 - Cross,3 - I-beam,4 - Icon,5 - Size,6 - Size NE SW,7 - Size N S,8 - Size NW SE,9 - Size W E,10 - Up Arrow,11 - Hourglass,12 - No drop,13 - Arrow & Hourglass,14 - Arrow & Question,15 - Size All,99 - Custom:Sets the mouse pointer look when mouse is over the control" & vbCrLf & _
    "MultiLine:BOOL::Specifies whether or not the text control can have multiple lines." & vbCrLf & _
    "PasswordChar:TEXT::Sets the character used to hide text in a password text control." & vbCrLf & _
    "Picture:PICTURE::Sets the picture to be displayed in the control." & vbCrLf & _
    "ScrollBars:ENUM:0 - None,1 - Horizontal,2 - Vertical,3 - Both:Sets the type of scrollbar to display on the control." & vbCrLf & _
    "Style:ENUM:0 - Standard,1 - Graphical:Specified whether the button is text-only or graphical." & vbCrLf & _
    "Text:TEXT::Sets the text to be displayed in the text control." & vbCrLf & _
    "Top:INT::Sets the Top coordinate of the control in twips (15 twips = 1 pixel)" & vbCrLf & "Visible:BOOL::Specifies whether or not the control is visible upon load" & vbCrLf & "Width:INT::Sets the width of the control in twips (15 twips = 1 pixel)" & "WindowState:ENUM:0 - Normal,1 - Minimized,2 - Maximized:Specifies the state of the window upon load." & vbCrLf & _
    "WordWrap:BOOL::Specifies whether or not text is wrapped when too long for a single line"


Sub CreateTheControl(dialog As Form, objcT As Control, x As Integer, Y As Integer, width As Integer, height As Integer)
    MsgBox x & "~" & Y & "~" & width & "~" & height
    If objcT.Tag = "" Then objcT.Tag = 0
    dialog.objcT(0).Tag = dialog.objcT(0).Tag + 1
    Load dialog.objcT(dialog.objcT(0).Tag)
    'object(object(0).Tag).Move StartX, StartY, EndX - StartX, EndY - StartY
    dialog.objcT(dialog.objcT(0).Tag).Move x, Y, width, height
    dialog.objcT(dialog.objcT(0).Tag).Visible = True
End Sub


Sub DrawTheGrid(frm As Form)
    
    frm.Cls
    If Not ShowGrid Then Exit Sub
    
    Dim x As Integer, Y As Integer, startpoint As Integer
    
    startpoint = 0 'GridSize \ 2
    
    For x = startpoint To frm.ScaleWidth Step GridSize
        For Y = startpoint To frm.ScaleHeight Step GridSize
            SetPixelV frm.hdc, x, Y, gridColor
        Next Y
    Next x
    
    
    
End Sub

Sub FillControlList(selectedObj As Object)
On Error Resume Next
    Dim i As Integer
    i = 1
    With Properties.cmbControls
        .Clear
        .AddItem "Dialog"
        If TypeOf selectedObj Is Form Then .ListIndex = 0
        Dim x As Control
        For Each x In selectedForm
            If x.Name Like "*Template" Then
                If x.Index <> 0 Then .AddItem x.Tag
                'MsgBox "~" & selectedObj.Tag & " = " & x.Tag & "~" & x.Index & "~"
                If selectedObj.Tag = x.Tag Then .ListIndex = .ListCount - 1
                i = i + 1
            End If
        Next x
    End With
End Sub

Sub FillProperties(obj As Object, PropList As Variant)
    Dim strData() As String, i As Integer
    Dim lvCount As Integer, var As ListItem
    Properties.lvProperties.ListItems.Clear
    
    On Error Resume Next
    strData = Split(PropList, ",")
    For i = LBound(strData) To UBound(strData)
       
        lvCount = Properties.lvProperties.ListItems.Count + 1
        Set var = Properties.lvProperties.ListItems.Add(, "", strData(i), 0, 0)
        'MsgBox "~" & RealPropName(LeftOf(strData(i), " ")) & "~"
        var.SubItems(1) = CallByName(obj, RealPropName(strData(i)), VbGet)
        
    Next i
End Sub

Function GetPropertyENum(strPropertyName, ENumIndex As Integer) As String
    Dim i As Integer, j As Integer
    
    For i = LBound(PropertiesA) To UBound(PropertiesA)
        If strPropertyName = PropertiesA(i).Name Then
            For j = 1 To PropertiesA(i).ENums
                If ENumIndex = j Then
                    GetPropertyENum = PropertiesA(i).ENumItems(j)
                    Exit Function
                End If
            Next j
        End If
    Next i
    GetPropertyENum = ""

End Function

Function GetPropertyENums(strPropertyName As String) As Integer
    Dim i As Integer
    
    For i = LBound(PropertiesA) To UBound(PropertiesA)
        If strPropertyName = PropertiesA(i).Name Then
            GetPropertyENums = UBound(PropertiesA(i).ENumItems)
            Exit Function
        End If
    Next i
    GetPropertyENums = 0
End Function

Function GetPropertyTip(strPropertyName As String) As String
    Dim i As Integer
    
    For i = LBound(PropertiesA) To UBound(PropertiesA)
        If strPropertyName = LeftOf(PropertiesA(i).Name, " ") Then
            GetPropertyTip = PropertiesA(i).Help
            Exit Function
        End If
    Next i
    GetPropertyTip = ""
End Function

Function GetPropertyType(strPropertyName As String) As String
    Dim i As Integer
    
    For i = LBound(PropertiesA) To UBound(PropertiesA)
        If strPropertyName = PropertiesA(i).Name Then
            GetPropertyType = PropertiesA(i).Type
            Exit Function
        End If
    Next i
    GetPropertyType = ""
End Function

Function GridX(x As Integer) As Integer
    If useGrid Then
        GridX = Int(Int(x \ GridSize) * GridSize) '+ (GridSize / 2)
    Else
        GridX = x
    End If
End Function


Function GridY(Y As Integer) As String
    If useGrid Then
        GridY = Int(Int(Y \ GridSize) * GridSize) '+ (GridSize / 2)
    Else
        GridY = Y
    End If
End Function


Function LeftOf(strText As String, strLeftOf As String) As String
    If InStr(strText, strLeftOf) Then
        LeftOf = Left(strText, InStr(strText, strLeftOf) - 1)
    Else
        LeftOf = strText
    End If
End Function

Sub LoadProperties()
    Dim strData() As String, strData2() As String, strData3() As String
    Dim i As Integer, j As Integer
    
    'On Error Resume Next
    strData = Split(Props, vbCrLf)
    For i = LBound(strData) To UBound(strData)
        PropCount = PropCount + 1
        ReDim Preserve PropertiesA(1 To PropCount) As Property
        strData2 = Split(strData(i), ":")
        With PropertiesA(PropCount)
            .Name = strData2(0)
            .Type = strData2(1)
            .Help = strData2(3)
            If InStr(strData2(2), ",") Then
                strData3 = Split(strData2(2), ",")
                .ENums = 0
                For j = LBound(strData3) To UBound(strData3)
                    .ENums = .ENums + 1
                    ReDim Preserve .ENumItems(1 To .ENums) As String
                    .ENumItems(.ENums) = strData3(j)
                Next j
            End If
        End With
    Next i
End Sub

Sub Main()
    useGrid = True
    gridColor = vbBlack
    GridSize = 8
    ShowGrid = True
    
    LoadProperties
    
    ToolBox.Show
    Properties.Show
    Properties.lvProperties.Picture = LoadPicture("")
    
    mainMenu.Show
        
End Sub

Function Pixels(Twips As Integer) As Integer
    Pixels = Int(Twips \ Screen.TwipsPerPixelX)
End Function


Public Function RealPropName(strPropName As String) As String
    
    Select Case strPropName
        Case "ID"
            RealPropName = "Tag"
        Case "BorderStyle Form"
            RealPropName = "BorderStyle"
        Case "Alignment Text"
            RealPropName = "Alignment"
        Case Else
            RealPropName = strPropName
    End Select
End Function


