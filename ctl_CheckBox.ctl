VERSION 5.00
Begin VB.UserControl CheckBox 
   AutoRedraw      =   -1  'True
   ClientHeight    =   300
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1455
   ScaleHeight     =   300
   ScaleWidth      =   1455
End
Attribute VB_Name = "CheckBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'
'
'
'10/25/2002
'written by: Eric Madison
'hopefully no errors
'updated: 10/27/2002
'thanks to john priestley for pointing out my oversights on tabstops and spacebar keypress :)
'thanks to jeroen  for pointing out my oversight on parent control as oposed to just a parent form :)
'
'
'
Option Explicit '<---keeps us honest

'api for drawing boxes and marks and focus rectangle
Private Declare Function RoundRect Lib "gdi32" (ByVal hdc As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function InflateRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, lpPoint As POINTAPI) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long

'api for freeing resources
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

'api for caption rect and printing caption
Private Declare Function DrawTextEx Lib "user32" Alias "DrawTextExA" (ByVal hdc As Long, ByVal lpsz As String, ByVal n As Long, lpRect As RECT, ByVal un As Long, ByVal lpDrawTextParams As Any) As Long

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

'declare enumerated properties for end user
'   properties window dropdown selections
Public Enum TheBackStyle
    Independent
    [Ambient Mode]
End Enum

Public Enum TheBoxStyle
    [3D]
    Flat
    Inset
    button
End Enum

Public Enum TheValue
    Unchecked
    Checked
End Enum

'declare end user public events
Public Event Click()
Public Event KeyDown(KeyCode As Integer, shift As Integer)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyUp(KeyCode As Integer, shift As Integer)
Public Event MouseDown(button As Integer, shift As Integer, x As Single, y As Single)
Public Event MouseMove(button As Integer, shift As Integer, x As Single, y As Single)
Public Event MouseUp(button As Integer, shift As Integer, x As Single, y As Single)

'store our font with stdfont to allow end user to
'   select a font at design time or run time
Private WithEvents TheFont As StdFont
Attribute TheFont.VB_VarHelpID = -1

'declare property variables
Private TheBoxStyleX As TheBoxStyle
Private TheBackStyleX As TheBackStyle, TheValueX As TheValue
Private TheForeColor As OLE_COLOR, TheBackColor As OLE_COLOR
Private TheBoxBorderDark As OLE_COLOR, TheBoxBorderLight As OLE_COLOR, TheMarkColor As OLE_COLOR
Private TheFocusColor As OLE_COLOR, TheBoxBackColor As OLE_COLOR
Private TheEnabled As Boolean, HasFocus As Boolean
Private TheBorderWidth As Integer
Private TheCaption As String

'declare general variables
Private OldBcolor As Long

'raise end user events as they occur
Private Sub UserControl_Click()
    If TheValueX = Unchecked Then
        TheValueX = Checked
    ElseIf TheValueX = Checked Then
        TheValueX = Unchecked
    End If
    DrawBox
    RaiseEvent Click
End Sub

Private Sub UserControl_GotFocus()
    HasFocus = True 'gain focus if tabstop
    DrawBox
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, shift As Integer)
    RaiseEvent KeyDown(KeyCode, shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then UserControl_Click 'if spacebar is pressed while control has focus
    'doesnt react exactly as if mouse was used (boxbackcolor doesnt change)...fix later
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, shift As Integer)
    RaiseEvent KeyUp(KeyCode, shift)
End Sub

Private Sub UserControl_LostFocus()
    HasFocus = False 'remove focus rectangle
    DrawBox
End Sub

Private Sub UserControl_MouseDown(button As Integer, shift As Integer, x As Single, y As Single)
    OldBcolor = TheBoxBackColor 'grab the boxbackcolor and hold it
    TheBoxBackColor = RGB(192, 192, 192) 'change boxbackcolor while mouse is down
    HasFocus = True
    DrawBox
    RaiseEvent MouseDown(button, shift, x, y)
End Sub

Private Sub UserControl_MouseMove(button As Integer, shift As Integer, x As Single, y As Single)
    RaiseEvent MouseMove(button, shift, x, y)
End Sub

Private Sub UserControl_MouseUp(button As Integer, shift As Integer, x As Single, y As Single)
    TheBoxBackColor = OldBcolor 'reset boxbackcolor to original color
    RaiseEvent MouseUp(button, shift, x, y)
End Sub

Private Sub UserControl_AmbientChanged(PropertyName As String)
    If TheBackStyleX = [Ambient Mode] Then 'if parent forms backcolor changes
        If PropertyName = "BackColor" Then
            DrawBox
        End If
    End If
End Sub

Private Sub TheFont_FontChanged(ByVal PropertyName As String)
    Set UserControl.Font = TheFont
    UserControl.Refresh
End Sub

Private Sub UserControl_Initialize()
    'create instance of the stdfont object and assign
    '   it to the controls font property
    Set TheFont = New StdFont
    Set UserControl.Font = TheFont
End Sub

Private Sub UserControl_InitProperties()
    BackColor = RGB(192, 192, 192)
    BoxBackColor = RGB(192, 192, 192)
    BoxBorderDark = RGB(128, 128, 128)
    BoxBorderLight = RGB(255, 255, 255)
    BoxStyle = 0
    BorderWidth = 1
    Caption = Extender.Name
    Enabled = True
    FocusColor = RGB(0, 0, 0)
    TheFont.Name = "Arial"
    ForeColor = RGB(0, 0, 0)
    MarkColor = RGB(0, 0, 0)
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        BackColor = .ReadProperty("BackColor", RGB(192, 192, 192))
        BackStyle = .ReadProperty("BackStyle", 0)
        BoxBackColor = .ReadProperty("BoxBackColor", RGB(192, 192, 192))
        BoxBorderDark = .ReadProperty("BoxBorderDark", RGB(128, 128, 128))
        BoxBorderLight = .ReadProperty("BoxBorderLight", RGB(255, 255, 255))
        BoxStyle = .ReadProperty("BoxStyle", 0)
        BorderWidth = .ReadProperty("BorderWidth", 1)
        Caption = .ReadProperty("Caption", Extender.Name)
        Enabled = .ReadProperty("Enabled", True)
        FocusColor = .ReadProperty("FocusColor", RGB(0, 0, 0))
        Set Font = .ReadProperty("Font")
        ForeColor = .ReadProperty("ForeColor", RGB(0, 0, 0))
        MarkColor = .ReadProperty("MarkColor", RGB(0, 0, 0))
        UserControl.MousePointer = .ReadProperty("MousePointer", vbDefault)
        Set UserControl.MouseIcon = .ReadProperty("MouseIcon", Nothing)
        Value = .ReadProperty("Value", 0)
    End With
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty "BackColor", BackColor, RGB(192, 192, 192)
        .WriteProperty "BackStyle", BackStyle, 0
        .WriteProperty "BoxBackColor", BoxBackColor, RGB(192, 192, 192)
        .WriteProperty "BoxBorderDark", BoxBorderDark, RGB(128, 128, 128)
        .WriteProperty "BoxBorderLight", BoxBorderLight, RGB(255, 255, 255)
        .WriteProperty "BoxStyle", BoxStyle, 0
        .WriteProperty "BorderWidth", BorderWidth, 1
        .WriteProperty "Caption", Caption, Extender.Name
        .WriteProperty "Enabled", Enabled, True
        .WriteProperty "Font", Font
        .WriteProperty "FocusColor", FocusColor, RGB(0, 0, 0)
        .WriteProperty "ForeColor", ForeColor, RGB(0, 0, 0)
        .WriteProperty "MarkColor", MarkColor, RGB(0, 0, 0)
        .WriteProperty "MousePointer", MousePointer, vbDefault
        .WriteProperty "MouseIcon", MouseIcon, Nothing
        .WriteProperty "Value", Value, 0
    End With
End Sub

Private Sub UserControl_Resize()
    DrawBox 'make sure to redraw when resizing
End Sub

'set up end user properties
Public Property Get BackColor() As OLE_COLOR
    BackColor = TheBackColor
End Property

Public Property Let BackColor(ByVal NewColor As OLE_COLOR)
    TheBackColor = NewColor
    DrawBox
    PropertyChanged "BackColor"
End Property

Public Property Get BackStyle() As TheBackStyle
    BackStyle = TheBackStyleX
End Property

Public Property Let BackStyle(ByVal NewStyle As TheBackStyle)
    TheBackStyleX = NewStyle
    DrawBox
    PropertyChanged "BackStyle"
End Property

Public Property Get BoxBackColor() As OLE_COLOR
    BoxBackColor = TheBoxBackColor
End Property

Public Property Let BoxBackColor(ByVal NewColor As OLE_COLOR)
    TheBoxBackColor = NewColor
    DrawBox
    PropertyChanged "BoxBackColor"
End Property

Public Property Get BoxBorderDark() As OLE_COLOR
    BoxBorderDark = TheBoxBorderDark
End Property

Public Property Let BoxBorderDark(ByVal NewColor As OLE_COLOR)
    TheBoxBorderDark = NewColor
    DrawBox
    PropertyChanged "BoxBorderDark"
End Property

Public Property Get BoxBorderLight() As OLE_COLOR
    BoxBorderLight = TheBoxBorderLight
End Property

Public Property Let BoxBorderLight(ByVal NewColor As OLE_COLOR)
    TheBoxBorderLight = NewColor
    DrawBox
    PropertyChanged "BoxBorderLight"
End Property

Public Property Get BoxStyle() As TheBoxStyle
    BoxStyle = TheBoxStyleX
End Property

Public Property Let BoxStyle(ByVal NewStyle As TheBoxStyle)
    TheBoxStyleX = NewStyle
    DrawBox
    PropertyChanged "BoxStyle"
End Property

Public Property Get BorderWidth() As Integer
   BorderWidth = TheBorderWidth
End Property

Public Property Let BorderWidth(ByVal NewWidth As Integer)
   If NewWidth > 1 And TheBoxStyleX <> 1 Then
      TheBorderWidth = 1
   Else
      TheBorderWidth = NewWidth
   End If
   DrawBox
   PropertyChanged "BorderWidth"
End Property

Public Property Get Caption() As String
    Caption = TheCaption
End Property

Public Property Let Caption(ByVal NewCaption As String)
    TheCaption = NewCaption
    DrawBox
    PropertyChanged "Caption"
End Property

Public Property Get Enabled() As Boolean
    Enabled = TheEnabled
End Property

Public Property Let Enabled(ByVal NewValue As Boolean)
    TheEnabled = NewValue
    If Enabled = True Then
        UserControl.Enabled = True
    Else
        UserControl.Enabled = False
    End If
    DrawBox
    PropertyChanged "Enabled"
End Property

Public Property Get FocusColor() As OLE_COLOR
    FocusColor = TheFocusColor
End Property

Public Property Let FocusColor(ByVal NewColor As OLE_COLOR)
    TheFocusColor = NewColor
    DrawBox
    PropertyChanged "FocusColor"
End Property

Public Property Get Font() As StdFont
    Set Font = TheFont
End Property

Public Property Set Font(NewFont As StdFont)
    If NewFont Is Nothing Then Exit Property
    With TheFont
        .Bold = NewFont.Bold
        .Charset = NewFont.Charset
        .Italic = NewFont.Italic
        .Name = NewFont.Name
        .Size = NewFont.Size
        .Strikethrough = NewFont.Strikethrough
        .Underline = NewFont.Underline
        .Weight = NewFont.Weight
    End With
    DrawBox
    PropertyChanged "Font"
End Property

Public Property Get ForeColor() As OLE_COLOR
    ForeColor = TheForeColor
End Property

Public Property Let ForeColor(ByVal NewColor As OLE_COLOR)
    TheForeColor = NewColor
    DrawBox
    PropertyChanged "ForeColor"
End Property

Public Property Get MarkColor() As OLE_COLOR
    MarkColor = TheMarkColor
End Property

Public Property Let MarkColor(ByVal NewColor As OLE_COLOR)
    TheMarkColor = NewColor
    DrawBox
    PropertyChanged "MarkColor"
End Property

Public Property Get MouseIcon() As StdPicture
    Set MouseIcon = UserControl.MouseIcon
End Property

Public Property Set MouseIcon(ByVal MouseIcon As StdPicture)
    Set UserControl.MouseIcon = MouseIcon
    PropertyChanged "MouseIcon"
End Property

Public Property Get MousePointer() As MousePointerConstants
    MousePointer = UserControl.MousePointer
End Property

Public Property Let MousePointer(ByVal MousePointer As MousePointerConstants)
    UserControl.MousePointer = MousePointer
    PropertyChanged "MousePointer"
End Property

Public Property Get Value() As TheValue
    Value = TheValueX
End Property

Public Property Let Value(ByVal NewValue As TheValue)
    TheValueX = NewValue
    DrawBox
    PropertyChanged "Value"
End Property

'this subroutine draws the checkbox, caption and marks
Private Sub DrawBox()
    Dim x As Integer, i As Integer
    Dim r As RECT
    Dim pt As POINTAPI
    Dim fBrush As Long, dcolor As Long, lColor As Long
    
    With UserControl
        .Cls 'erase everything before redrawing
        .ScaleMode = 3 'make sure to set to pixels for api
        
        'set backcolor
        If TheBackStyleX = 0 Then 'if control has independent backcolor
            .BackColor = TheBackColor 'will have own backcolor
        Else
            .BackColor = Extender.Container.BackColor 'will change to match parent backcolor
        End If
        
        'set box co-ordinates
        r.Top = .ScaleTop + .ScaleHeight / 2 - 13 / 2
        r.Left = .ScaleLeft
        If TheBoxStyleX = Flat Then
            r.Bottom = .ScaleTop + .ScaleHeight / 2 + 13 / 2
            r.Right = .ScaleLeft + 13
        Else
            r.Bottom = .ScaleTop + .ScaleHeight / 2 + 12 / 2
            r.Right = .ScaleLeft + 12
        End If
        
        'draw the box
        If TheEnabled = False Then 'check if usercontrol is enabled
            fBrush = CreateSolidBrush(RGB(192, 192, 192))
        Else
            fBrush = CreateSolidBrush(TheBoxBackColor)
        End If
        FillRect .hdc, r, fBrush 'fill the box with selected brush
        Select Case TheBoxStyleX
            Case [3D]
                MoveToEx .hdc, r.Right, r.Top, pt
                .ForeColor = TheBoxBorderDark
                LineTo .hdc, r.Left, r.Top
                LineTo .hdc, r.Left, r.Bottom
                .ForeColor = TheBoxBorderLight
                LineTo .hdc, r.Right, r.Bottom
                LineTo .hdc, r.Right, r.Top - 1
                MoveToEx .hdc, r.Right - 1, r.Top + 1, pt
                .ForeColor = RGB(0, 0, 0)
                LineTo .hdc, r.Left + 1, r.Top + 1
                LineTo .hdc, r.Left + 1, r.Bottom - 1
                .ForeColor = RGB(192, 192, 192)
                LineTo .hdc, r.Right - 1, r.Bottom - 1
                LineTo .hdc, r.Right - 1, r.Top
            Case Flat
                .ForeColor = TheBoxBorderDark
                For x = 1 To TheBorderWidth
                    RoundRect .hdc, r.Left, r.Top, r.Right, r.Bottom, 0, 0
                    InflateRect r, -1, -1
                Next
            Case Inset
                MoveToEx .hdc, r.Right, r.Top, pt
                .ForeColor = TheBoxBorderDark
                LineTo .hdc, r.Left, r.Top
                LineTo .hdc, r.Left, r.Bottom
                .ForeColor = TheBoxBorderLight
                LineTo .hdc, r.Right, r.Bottom
                LineTo .hdc, r.Right, r.Top - 1
            Case button
                MoveToEx .hdc, r.Right, r.Top, pt
                If TheValueX = Unchecked Then
                    lColor = TheBoxBorderLight
                    dcolor = TheBoxBorderDark
                Else
                    lColor = TheBoxBorderDark
                    dcolor = TheBoxBorderLight
                End If
                .ForeColor = lColor
                LineTo .hdc, r.Left, r.Top
                LineTo .hdc, r.Left, r.Bottom
                .ForeColor = dcolor
                LineTo .hdc, r.Right, r.Bottom
                LineTo .hdc, r.Right, r.Top - 1
        End Select
        DeleteObject fBrush 'get rid of brush to free up resources

        'draw marks
        .ForeColor = TheMarkColor
        If TheBoxStyleX = Flat Then
            r.Top = r.Top - 1 'offset one pixel since flat box is pixel taller than the rest
        End If
        If TheValueX = Checked Then
            'this could probably be done in a more code efficient manner but.... (without bitmap etc.)
            'draw right angle line for checkmark
            MoveToEx .hdc, 9, r.Top + 5, pt
            LineTo .hdc, 5, r.Top + 9
            MoveToEx .hdc, 9, r.Top + 4, pt
            LineTo .hdc, 4, r.Top + 9
            MoveToEx .hdc, 9, r.Top + 3, pt
            LineTo .hdc, 3, r.Top + 9
            'draw left angle line for checkmark
            MoveToEx .hdc, 3, r.Top + 5, pt
            LineTo .hdc, 5, r.Top + 8
            MoveToEx .hdc, 3, r.Top + 6, pt
            LineTo .hdc, 6, r.Top + 9
            MoveToEx .hdc, 3, r.Top + 7, pt
            LineTo .hdc, 6, r.Top + 10
            MoveToEx .hdc, 3, r.Top + 7, pt
            LineTo .hdc, 7, r.Top + 9
        End If
         
        'set caption co-ordinates to allow for font sizing
        r.Top = .ScaleTop + .ScaleHeight / 2 - .TextHeight(TheCaption) / 2
        r.Bottom = .ScaleTop + .ScaleHeight / 2 + .TextHeight(TheCaption) / 2
        r.Top = r.Top - 1
        'draw caption
        If TheEnabled = False Then 'if checkbox is disabled draw caption inset for disabled appearance
            .ForeColor = RGB(255, 255, 255)
            r.Right = .ScaleLeft + .ScaleWidth
            r.Left = .ScaleLeft + 19
            DrawTextEx .hdc, TheCaption, Len(TheCaption), r, 0&, 0&
            .ForeColor = RGB(128, 128, 128)
            r.Right = .ScaleLeft + .ScaleWidth
            r.Left = .ScaleLeft + 18
            DrawTextEx .hdc, TheCaption, Len(TheCaption), r, 0&, 0&
        Else 'it's enabled
            .ForeColor = TheForeColor
            r.Right = .ScaleLeft + .ScaleWidth
            r.Left = .ScaleLeft + 19
            DrawTextEx .hdc, TheCaption, Len(TheCaption), r, 0&, 0&
        End If
        
        'draw focus rectangle as needed
        If HasFocus = True Then
            r.Left = 19
            r.Right = ScaleLeft + 19 + .TextWidth(TheCaption)
            For i = r.Left - 2 To r.Right Step 2
                SetPixel .hdc, i, r.Top, TheFocusColor
            Next
            For i = r.Left - 2 To r.Right Step 2
                SetPixel .hdc, i, r.Top + .TextHeight(TheCaption), TheFocusColor
            Next
            For i = r.Top To r.Top + .TextHeight(TheCaption) Step 2
                SetPixel .hdc, r.Left - 2, i, TheFocusColor
            Next
            For i = r.Top To r.Top + .TextHeight(TheCaption) Step 2
                SetPixel .hdc, r.Right, i, TheFocusColor
            Next
        End If
    End With
End Sub

