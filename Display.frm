VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{EE128208-4F73-11D3-83BB-C47C02EE3D01}#1.0#0"; "CONTROLRESIZER.OCX"
Begin VB.Form frmDisplay 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Enter a Value"
   ClientHeight    =   1320
   ClientLeft      =   3750
   ClientTop       =   2745
   ClientWidth     =   2805
   Icon            =   "Display.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1320
   ScaleWidth      =   2805
   Begin VB.CommandButton Command3 
      Cancel          =   -1  'True
      Caption         =   "Clear"
      Height          =   255
      Left            =   2160
      TabIndex        =   4
      Top             =   480
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      Height          =   255
      Left            =   2160
      TabIndex        =   3
      Top             =   960
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "GO"
      Default         =   -1  'True
      Height          =   255
      Left            =   2160
      TabIndex        =   1
      Top             =   120
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Text            =   "300"
      Top             =   120
      Width           =   1935
   End
   Begin ControlResizer.AutoResizer AutoResizer1 
      Height          =   375
      Left            =   1680
      Top             =   1680
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
   End
   Begin VB.Label Label1 
      Caption         =   ","
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1155
      TabIndex        =   2
      Top             =   840
      Width           =   150
   End
   Begin VB.Image LEDA 
      Height          =   600
      Left            =   120
      Picture         =   "Display.frx":030A
      Stretch         =   -1  'True
      Top             =   600
      Width           =   345
   End
   Begin ComctlLib.ImageList LEDs 
      Left            =   720
      Top             =   1560
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   23
      ImageHeight     =   40
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   11
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Display.frx":0E8C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Display.frx":1A1E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Display.frx":25B0
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Display.frx":3142
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Display.frx":3CD4
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Display.frx":4866
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Display.frx":53F8
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Display.frx":5F8A
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Display.frx":6B1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Display.frx":76AE
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Display.frx":8240
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Image LEDE 
      Height          =   600
      Left            =   1650
      Picture         =   "Display.frx":8DD2
      Stretch         =   -1  'True
      Top             =   600
      Width           =   345
   End
   Begin VB.Image LEDD 
      Height          =   600
      Left            =   1305
      Picture         =   "Display.frx":9954
      Stretch         =   -1  'True
      Top             =   600
      Width           =   345
   End
   Begin VB.Image LEDC 
      Height          =   600
      Left            =   810
      Picture         =   "Display.frx":A4D6
      Stretch         =   -1  'True
      Top             =   600
      Width           =   345
   End
   Begin VB.Image LEDB 
      Height          =   600
      Left            =   465
      Picture         =   "Display.frx":B058
      Stretch         =   -1  'True
      Top             =   600
      Width           =   345
   End
End
Attribute VB_Name = "frmDisplay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub DisplayConv(LED100 As Image, LED10 As Image, LED1 As Image, LED01 As Image, LED001 As Image)
Dim DValue As Double
Dim V100 As Boolean
Dim Current As Byte
    
    On Error Resume Next
    
    DValue = Text1.Text
    
    If Int(DValue / 100) >= 10 Then
        LED100.Picture = LEDs.ListImages(9).Picture
        LED10.Picture = LEDs.ListImages(9).Picture
        LED1.Picture = LEDs.ListImages(9).Picture
        LED01.Picture = LEDs.ListImages(9).Picture
        LED001.Picture = LEDs.ListImages(9).Picture
        Exit Sub
    End If
    
    If Int(DValue / 100) >= 1 Then
        LED100.Picture = LEDs.ListImages(Int(DValue / 100) + 1).Picture
        V100 = True
    Else
        LED100.Picture = LEDs.ListImages(11).Picture
        V100 = False
    End If
    
    DValue = DValue - (Int(DValue / 100) * 100)
    If Int(DValue / 10) >= 1 Then
        LED10.Picture = LEDs.ListImages(Int(DValue / 10) + 1).Picture
    Else
        If V100 Then
            LED10.Picture = LEDs.ListImages(1).Picture
        Else
            LED10.Picture = LEDs.ListImages(11).Picture
        End If
    End If

    DValue = DValue - (Int(DValue / 10) * 10)
    If Int(DValue) >= 1 Then
        LED1.Picture = LEDs.ListImages(Int(DValue) + 1).Picture
    Else
        LED1.Picture = LEDs.ListImages(1).Picture
    End If

    DValue = Round((DValue - Int(DValue)) * 100)
    If Int(DValue / 10) >= 1 Then
        Current = Int(DValue / 10) + 1
        LED01.Picture = LEDs.ListImages(Int(DValue / 10) + 1).Picture
    Else
        LED01.Picture = LEDs.ListImages(1).Picture
    End If

    DValue = DValue - (Int(DValue / 10) * 10)
    If Int(DValue) >= 1 Then
        If Int(Round(DValue)) >= 10 Then
            If Current < 9 Then
                LED01.Picture = LEDs.ListImages(Current + 1).Picture
                LED001.Picture = LEDs.ListImages(1).Picture
            Else
                LED01.Picture = LEDs.ListImages(10).Picture
                LED001.Picture = LEDs.ListImages(10).Picture
            End If
        Else
            LED001.Picture = LEDs.ListImages(Int(Round(DValue)) + 1).Picture
        End If
    Else
        LED001.Picture = LEDs.ListImages(1).Picture
    End If

End Sub

Private Sub Command1_Click()
    DisplayConv LEDA, LEDB, LEDC, LEDD, LEDE
End Sub

Private Sub Command2_Click()
    End
End Sub

Private Sub Command3_Click()
    Text1.Text = ""
    LEDA.Picture = LEDs.ListImages(11).Picture
    LEDB.Picture = LEDs.ListImages(11).Picture
    LEDC.Picture = LEDs.ListImages(11).Picture
    LEDD.Picture = LEDs.ListImages(11).Picture
    LEDE.Picture = LEDs.ListImages(11).Picture
    Text1.SetFocus
End Sub

Private Sub Form_Load()
    Text1.SelStart = 0
    Text1.SelLength = Len(Text1.Text)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub Text1_Change()
Dim PointPos As Integer

    PointPos = InStr(Text1.Text, ".")
    If PointPos <> 0 Then
        Beep
        Text1.Text = Left(Text1.Text, PointPos - 1) & Right(Text1.Text, Len(Text1.Text) - PointPos)
        Text1.SelStart = PointPos - 1
    End If

End Sub

