VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Alpha Mask Selection"
   ClientHeight    =   4140
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6240
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   276
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   416
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog C2 
      Left            =   3600
      Top             =   2400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Save Extracted Picture As ..."
      Filter          =   "Bitmap File (*.bmp)|*.bmp|"
   End
   Begin VB.PictureBox Picture4 
      AutoRedraw      =   -1  'True
      FillColor       =   &H00FFFFFF&
      ForeColor       =   &H00FFFFFF&
      Height          =   2025
      Left            =   5520
      ScaleHeight     =   131
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   160
      TabIndex        =   15
      Top             =   360
      Visible         =   0   'False
      Width           =   2460
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Save"
      Enabled         =   0   'False
      Height          =   615
      Left            =   2760
      TabIndex        =   1
      Top             =   1680
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "OK"
      Height          =   495
      Left            =   4920
      TabIndex        =   4
      Top             =   3480
      Width           =   1095
   End
   Begin MSComDlg.CommonDialog c1 
      Left            =   3000
      Top             =   2400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.Slider Slider1 
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   2880
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   450
      _Version        =   393216
      Enabled         =   0   'False
      Min             =   1
      Max             =   240
      SelStart        =   200
      TickStyle       =   3
      Value           =   200
   End
   Begin VB.PictureBox Picture3 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      ForeColor       =   &H00FFFFFF&
      Height          =   5655
      Left            =   6120
      ScaleHeight     =   373
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   301
      TabIndex        =   8
      Top             =   120
      Visible         =   0   'False
      Width           =   4575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "--->>"
      Enabled         =   0   'False
      Height          =   735
      Left            =   2760
      TabIndex        =   0
      Top             =   480
      Width           =   615
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      FillColor       =   &H00FFFFFF&
      ForeColor       =   &H00FFFFFF&
      Height          =   2025
      Left            =   3600
      ScaleHeight     =   131
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   160
      TabIndex        =   6
      Top             =   360
      Width           =   2460
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      FillColor       =   &H00FFFFFF&
      ForeColor       =   &H00FFFFFF&
      Height          =   2025
      Left            =   120
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   131
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   160
      TabIndex        =   5
      Top             =   360
      Width           =   2460
   End
   Begin VB.Label Label8 
      Caption         =   "PLEASE VOTE FOR ME !!!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1200
      TabIndex        =   14
      Top             =   3720
      Width           =   2535
   End
   Begin VB.Label Label7 
      Caption         =   "tuckwai98@yahoo.com"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   3120
      TabIndex        =   13
      Top             =   3480
      Width           =   1335
   End
   Begin VB.Label Label6 
      Caption         =   "Suggestion && Comments please send to : "
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   3480
      Width           =   3015
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   495
      Left            =   5280
      TabIndex        =   3
      Top             =   2760
      Width           =   495
   End
   Begin VB.Label Label4 
      Caption         =   "Mask Colour :"
      Height          =   255
      Left            =   5040
      TabIndex        =   11
      Top             =   2520
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Mask Transparency Level :"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   2520
      Width           =   2415
   End
   Begin VB.Label Label2 
      Caption         =   "Selected Mask :"
      Height          =   255
      Left            =   3600
      TabIndex        =   9
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Picture Source (Make Selection Here) :"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   3015
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Original Selection Module Written By Mohammed
'Author : Looi Tuck Wai
'Date   : 8/8/2006
Option Explicit
Const AC_SRC_OVER = &H0
' This structure holds the arguments required by Alphablend function to work
Private Type BLENDFUNCTION
  BlendOp As Byte
  BlendFlags As Byte
  SourceConstantAlpha As Byte
  AlphaFormat As Byte
End Type
' This is the main API that is blending the pictures
Private Declare Function AlphaBlend Lib "msimg32.dll" (ByVal hdc As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal hdc As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal BLENDFUNCT As Long) As Long
' This is a commenly used API function(maybe by me only) which is very helpful to Tranfer ALL the values of a 'Structure'(Type) to a Long variable
Private Declare Sub RtlMoveMemory Lib "kernel32.dll" (Destination As Any, Source As Any, ByVal Length As Long)
' Being used by the Timer
Dim Counter As Long
' The BlendFunction 'Structure' is used by the 'AlphaBlend' API function
Dim BF As BLENDFUNCTION
' Actually the AlphaBlend API Function requires a refrence to a "LONG" value containing the values of BlendFunction structure!. This Variale holds the values done in the BlendFunction Structure.
' A Structure (Type) can be converted into a 'Long' value by using the 'RtlMoveMemory' API Function.. See below for its example ;)
Dim lBF As Long
Dim X1 As Integer
Dim Y1 As Integer
Dim Width1 As Integer
Dim Height1 As Integer
Private Sub Command1_Click()
    Picture2.Cls
    Picture2.Picture = Picture1.Image
    With BF
        .BlendOp = AC_SRC_OVER
        .BlendFlags = 0
        .SourceConstantAlpha = Slider1.Value
        .AlphaFormat = 0
    End With
    
    'copy the BLENDFUNCTION-structure to a Long
    RtlMoveMemory lBF, BF, 4
    
    'AlphaBlend the picture from Picture1 over the picture of Picture2
    AlphaBlend Picture2.hdc, 0, 0, Picture2.ScaleWidth, Picture2.ScaleHeight, Picture3.hdc, 0, 0, Picture3.ScaleWidth, Picture3.ScaleHeight, lBF
    Picture2.Refresh
    Picture2.PaintPicture Picture1.Picture, X1, Y1, , , X1, Y1, Width1, Height1
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
On Error GoTo err:
Picture4.Cls
    Picture4.PaintPicture Picture1.Picture, 0, 0, , , X1, Y1, Width1, Height1
    Set Picture4.Picture = Picture4.Image
    C2.ShowSave
    C2.DialogTitle = "Save Extracted Picture As ..."
    SavePicture Picture4.Picture, C2.FileName
    MsgBox "Extracted Picture Saved To " & C2.FileName
err:
Exit Sub
End Sub

Private Sub Form_Load()
    Label3.Caption = "Mask Transparency Level : " & Slider1.Value
End Sub

Private Sub Label5_Click()
c1.ShowColor
Picture3.BackColor = c1.Color
    Picture2.Cls
    Picture2.Picture = Picture1.Image
    With BF
        .BlendOp = AC_SRC_OVER
        .BlendFlags = 0
        .SourceConstantAlpha = Slider1.Value
        .AlphaFormat = 0
    End With
    
    'copy the BLENDFUNCTION-structure to a Long
    RtlMoveMemory lBF, BF, 4
    
    'AlphaBlend the picture from Picture1 over the picture of Picture2
    AlphaBlend Picture2.hdc, 0, 0, Picture2.ScaleWidth, Picture2.ScaleHeight, Picture3.hdc, 0, 0, Picture3.ScaleWidth, Picture3.ScaleHeight, lBF
    Picture2.Refresh
    Picture2.PaintPicture Picture1.Picture, X1, Y1, , , X1, Y1, Width1, Height1


End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo err:
    If Button = vbLeftButton Then
        Picture1.Cls
        X1 = X
        Y1 = Y
Picture4.Width = X1
Picture4.Height = Y1
    End If
    Command1.Enabled = True
    Slider1.Enabled = True
    Label5.Enabled = True
    Command3.Enabled = True
err:
Exit Sub
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo err:
    If Button = vbLeftButton Then
        Picture1.Cls

        Picture1.Line (X1, Y1)-(X, Y), , B
        Picture4.Width = X - X1
        Picture4.Height = Y - Y1
    End If
err:
Exit Sub
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo err:
    If Button = vbLeftButton Then
        If X > X1 Then
            Width1 = X - X1
            Picture4.Width = X - X1
        Else
            Width1 = X1 - X
            X1 = X
          Picture4.Width = X1
        End If
        If Y > Y1 Then
            Height1 = Y - Y1
            Picture4.Height = Y - Y1
        Else
            Height1 = Y1 - Y
            Y1 = Y
            Picture4.Height = Y1
        End If
    End If
err:
Exit Sub
End Sub

Private Sub Slider1_Scroll()
    Picture2.Cls
    Picture2.Picture = Picture1.Image
    With BF
        .BlendOp = AC_SRC_OVER
        .BlendFlags = 0
        .SourceConstantAlpha = Slider1.Value
        .AlphaFormat = 0
    End With
    Label3.Caption = "Mask Transparency Level : " & Slider1.Value
    'copy the BLENDFUNCTION-structure to a Long
    RtlMoveMemory lBF, BF, 4
    
    'AlphaBlend the picture from Picture1 over the picture of Picture2
    AlphaBlend Picture2.hdc, 0, 0, Picture2.ScaleWidth, Picture2.ScaleHeight, Picture3.hdc, 0, 0, Picture3.ScaleWidth, Picture3.ScaleHeight, lBF
    Picture2.Refresh
    Picture2.PaintPicture Picture1.Picture, X1, Y1, , , X1, Y1, Width1, Height1

End Sub
