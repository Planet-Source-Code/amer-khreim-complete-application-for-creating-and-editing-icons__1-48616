VERSION 4.00
Begin VB.Form ColorPalette 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "IconWorks Color Palette"
   ClientHeight    =   3090
   ClientLeft      =   1590
   ClientTop       =   1935
   ClientWidth     =   5670
   ClipControls    =   0   'False
   BeginProperty Font 
      name            =   "System"
      charset         =   1
      weight          =   700
      size            =   9.75
      underline       =   0   'False
      italic          =   0   'False
      strikethrough   =   0   'False
   EndProperty
   Height          =   3495
   HelpContextID   =   1906
   Icon            =   "COLORPAL.frx":0000
   Left            =   1530
   LinkMode        =   1  'Source
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   5670
   Top             =   1590
   Width           =   5790
   Begin VB.PictureBox Pic_ColorPalette 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   700
         size            =   8.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1005
      Left            =   60
      ScaleHeight     =   975
      ScaleWidth      =   5520
      TabIndex        =   0
      Top             =   60
      Width           =   5550
   End
   Begin VB.HScrollBar Scrl_RGB 
      Height          =   300
      Index           =   0
      LargeChange     =   10
      Left            =   750
      Max             =   255
      TabIndex        =   4
      Top             =   1260
      Width           =   2550
   End
   Begin VB.TextBox Txt_RGB 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   700
         size            =   8.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Index           =   0
      Left            =   3330
      TabIndex        =   7
      Top             =   1260
      Width           =   480
   End
   Begin VB.PictureBox Pic_RGB 
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   700
         size            =   8.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Index           =   0
      Left            =   3840
      ScaleHeight     =   390
      ScaleWidth      =   555
      TabIndex        =   15
      Top             =   1260
      Width           =   585
   End
   Begin VB.PictureBox Pic_SelectedColor 
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   700
         size            =   8.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   630
      Left            =   4455
      ScaleHeight     =   600
      ScaleWidth      =   1170
      TabIndex        =   10
      Top             =   1440
      Width           =   1200
   End
   Begin VB.PictureBox Pic_RGB 
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   700
         size            =   8.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Index           =   1
      Left            =   3840
      ScaleHeight     =   390
      ScaleWidth      =   555
      TabIndex        =   16
      Top             =   1680
      Width           =   585
   End
   Begin VB.HScrollBar Scrl_RGB 
      Height          =   300
      Index           =   1
      LargeChange     =   10
      Left            =   750
      Max             =   255
      TabIndex        =   5
      Top             =   1740
      Width           =   2550
   End
   Begin VB.TextBox Txt_RGB 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   700
         size            =   8.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Index           =   1
      Left            =   3315
      TabIndex        =   8
      Top             =   1740
      Width           =   480
   End
   Begin VB.PictureBox Pic_RGB 
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   700
         size            =   8.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Index           =   2
      Left            =   3840
      ScaleHeight     =   390
      ScaleWidth      =   555
      TabIndex        =   17
      Top             =   2100
      Width           =   585
   End
   Begin VB.HScrollBar Scrl_RGB 
      Height          =   300
      Index           =   2
      LargeChange     =   10
      Left            =   750
      Max             =   255
      TabIndex        =   6
      Top             =   2220
      Width           =   2550
   End
   Begin VB.TextBox Txt_RGB 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   700
         size            =   8.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Index           =   2
      Left            =   3330
      TabIndex        =   9
      Top             =   2220
      Width           =   480
   End
   Begin VB.PictureBox Pic_NearestSolidColor 
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   700
         size            =   8.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   630
      Left            =   4455
      ScaleHeight     =   600
      ScaleWidth      =   1170
      TabIndex        =   11
      Top             =   2445
      Width           =   1200
   End
   Begin VB.CommandButton Cmd_OK 
      Cancel          =   -1  'True
      Caption         =   "&Done"
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   700
         size            =   8.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   75
      TabIndex        =   12
      Top             =   2640
      Width           =   1305
   End
   Begin VB.CommandButton Cmd_Set 
      Caption         =   "&Set"
      Default         =   -1  'True
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   700
         size            =   8.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1560
      TabIndex        =   13
      Top             =   2640
      Width           =   1305
   End
   Begin VB.CommandButton Cmd_Reset 
      Caption         =   "&Reset"
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   700
         size            =   8.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3045
      TabIndex        =   14
      Top             =   2640
      Width           =   1305
   End
   Begin VB.Label Lbl_RGBValues 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "RGB Values"
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   700
         size            =   8.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   3360
      TabIndex        =   20
      Top             =   1065
      Width           =   1035
   End
   Begin VB.Label Lbl_SelectedColor 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Selected Color"
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   700
         size            =   8.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   4605
      TabIndex        =   18
      Top             =   1065
      Width           =   975
   End
   Begin VB.Label Lbl_Red 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Red"
      ForeColor       =   &H000000FF&
      Height          =   300
      Left            =   120
      TabIndex        =   1
      Top             =   1260
      Width           =   600
   End
   Begin VB.Label Lbl_Green 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Green"
      ForeColor       =   &H00008000&
      Height          =   300
      Left            =   120
      TabIndex        =   2
      Top             =   1740
      Width           =   600
   End
   Begin VB.Label Lbl_NearestSolidColor 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Nearest Solid Color"
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   700
         size            =   8.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   4605
      TabIndex        =   19
      Top             =   2070
      Width           =   975
   End
   Begin VB.Label Lbl_Blue 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Blue"
      ForeColor       =   &H00FF0000&
      Height          =   300
      Left            =   120
      TabIndex        =   3
      Top             =   2220
      Width           =   600
   End
End
Attribute VB_Name = "ColorPalette"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
DefInt A-Z

Private Sub Cmd_OK_Click()
    Unload ColorPalette
End Sub

Private Sub Cmd_Reset_Click()
    Initialize_RGB_Scrollbars
End Sub

' Places new color into the ColorPalette and Refreshes
' the color palettes so the new colors are displayed.
Private Sub Cmd_Set_Click()
    ' Create the Long Integer RGB value from the RGB scrollbar values, and
    ' place into Color array.
    Colors(ColorIndex) = RGB(Scrl_RGB(0).Value, Scrl_RGB(1).Value, Scrl_RGB(2).Value)
    ' Display new ColorPalette
    Display_Color_Palette Pic_ColorPalette
    Display_Color_Palette Editor.Pic_ColorPalette
End Sub

Private Sub Display_New_Color_And_Elements(FirstElement, LastElement)
    Pic_SelectedColor.BackColor = RGB(Scrl_RGB(0).Value, Scrl_RGB(1).Value, Scrl_RGB(2).Value)
    ' Since some of the drawing tools cannot use dithered colors,
    ' the nearest Solid color to the actual color selected is also displayed.
    Pic_NearestSolidColor.BackColor = GetNearestColor(hdc, Pic_SelectedColor.BackColor)
    For I = FirstElement To LastElement
        Txt_RGB(I).Text = Format$(Scrl_RGB(I).Value)
        Pic_RGB(I).BackColor = Scrl_RGB(I).Value * 2 ^ (I * 8)
    Next I
End Sub

Private Sub Form_Load()
    ColorPaletteLoaded = True
    Remove_Items_From_Sysmenu ColorPalette
End Sub

' Extracts the Red, Green, and Blue elements from the
' selected ColorPalette color and assigns these values to the
' corresponding RGB Scrollbars.
Private Sub Initialize_RGB_Scrollbars()
    Scrl_RGB(RED_ELEMENT).Value = Colors(ColorIndex) And &hff&
    Scrl_RGB(GREEN_ELEMENT).Value = (Colors(ColorIndex) \ 2 ^ 8) And &hff&
    Scrl_RGB(BLUE_ELEMENT).Value = (Colors(ColorIndex) \ 2 ^ 16) And &hff&
    ' Display the numerical and visual values for these Elements
    ' along with the selected color and its nearest solid color.
    Display_New_Color_And_Elements RED_ELEMENT, BLUE_ELEMENT
End Sub

Private Sub Pic_ColorPalette_GotFocus()
    ' Pic_ColorPalette has a tabindex of 0, thus it receives the focus
    ' first when the ColorPalette form gains the focus, so Initialization
    ' is done here.
    Initialize_RGB_Scrollbars
End Sub

Private Sub Pic_ColorPalette_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' Check if Mouse Coordinates are within the ColorPalette
    If (X >= 0) And (X <= 16) And (Y >= 0) And (Y <= 3) Then
        ' Set the Editor's current drawing color to selected color.
        Update_Mouse_Colors Button, X, Y
        ' Display selected color and elements of selected color
        Initialize_RGB_Scrollbars
    End If
End Sub

Private Sub Pic_ColorPalette_Paint()
    Display_Color_Palette Pic_ColorPalette
End Sub

Private Sub Scrl_RGB_Change(Index As Integer)
    Display_New_Color_And_Elements Index, Index
End Sub

Private Sub Txt_RGB_Change(Index As Integer)
    If Val(Txt_RGB(Index).Text) > 255 Then
        ' A value outside the value RGB range was entered.  Beep
        ' to signal the user, then reset value to previous value
        Beep
        Txt_RGB(Index).Text = Format$(Scrl_RGB(Index).Value)
    Else
        ' A valid RGB value was entered so reset corresponding RGB Scrollbar
        Scrl_RGB(Index).Value = Val(Txt_RGB(Index).Text)
    End If
    Txt_RGB(Index).SelStart = Len(Txt_RGB(Index).Text)
End Sub

Private Sub Txt_RGB_KeyPress(Index As Integer, KeyAscii As Integer)
    ' Do not allow any characters other than 0123456789 to be entered.
    If ((KeyAscii < 48) Or (KeyAscii > 57)) And (KeyAscii <> 8) Then
        KeyAscii = 0
        Beep
    End If
End Sub

