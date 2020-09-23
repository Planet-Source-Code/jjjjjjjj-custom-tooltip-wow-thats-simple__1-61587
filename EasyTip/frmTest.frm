VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmTest 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Easy ToolTip     By Jim Jose"
   ClientHeight    =   6255
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7365
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6255
   ScaleWidth      =   7365
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdNone 
      Caption         =   "None"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   2880
      Width           =   2055
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1080
      Width           =   2055
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1560
      Width           =   2055
   End
   Begin VB.CommandButton cmdInfo 
      Caption         =   "Info"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3360
      Width           =   2055
   End
   Begin VB.CommandButton cmdError 
      Caption         =   "Error"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3840
      Width           =   2055
   End
   Begin VB.CommandButton cmdWarning 
      Caption         =   "Warning"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4320
      Width           =   2055
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5160
      Picture         =   "frmTest.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5400
      Width           =   2055
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   600
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      Height          =   5775
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4695
      Begin VB.CommandButton cmdBack 
         BackColor       =   &H00FFFFFF&
         Caption         =   "..."
         Height          =   375
         Left            =   3960
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   3960
         Width           =   375
      End
      Begin VB.CommandButton cmdText 
         BackColor       =   &H00FF0000&
         Caption         =   "..."
         Height          =   375
         Left            =   2160
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   3960
         Width           =   375
      End
      Begin MSComDlg.CommonDialog cdlg 
         Left            =   4200
         Top             =   120
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.ListBox lstType 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1260
         ItemData        =   "frmTest.frx":058A
         Left            =   1920
         List            =   "frmTest.frx":059A
         TabIndex        =   16
         Top             =   2640
         Width           =   2415
      End
      Begin VB.ListBox lstStyle 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         ItemData        =   "frmTest.frx":05E2
         Left            =   1920
         List            =   "frmTest.frx":05EC
         TabIndex        =   14
         Top             =   1680
         Width           =   2415
      End
      Begin VB.TextBox txtHead 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1920
         TabIndex        =   12
         Text            =   "Easy Tooltip"
         Top             =   1080
         Width           =   2415
      End
      Begin VB.TextBox txtTool 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1920
         TabIndex        =   11
         Text            =   "This is a test tooltip specialy designed for ur UNSUBCLASSED app"
         Top             =   600
         Width           =   2415
      End
      Begin VB.CommandButton cmdTest 
         Caption         =   "Test ToolTip"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   360
         TabIndex        =   8
         Top             =   4560
         Width           =   3975
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ToolTip_TextColor"
         Height          =   240
         Left            =   360
         TabIndex        =   18
         Top             =   4080
         Width           =   1695
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "BackColor"
         Height          =   240
         Left            =   2760
         TabIndex        =   17
         Top             =   4080
         Width           =   945
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ToolTip_Type"
         Height          =   240
         Left            =   360
         TabIndex        =   15
         Top             =   3120
         Width           =   1305
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ToolTip_Style"
         Height          =   240
         Left            =   360
         TabIndex        =   13
         Top             =   2040
         Width           =   1275
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ToolTip_Head"
         Height          =   240
         Left            =   360
         TabIndex        =   10
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ToolTip_Text"
         Height          =   240
         Left            =   360
         TabIndex        =   9
         Top             =   720
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAdd_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ShowToolTip cmdAdd.hwnd, "Click here to Add new entry!!!", "Add", Tip_Balloon, Tip_Info
End Sub

Private Sub cmdBack_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ShowToolTip cmdBack.hwnd, "Click here to change the tooltip backcolor!!!", "BackColor", Tip_Balloon, Tip_Info
End Sub

Private Sub cmdClose_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ShowToolTip cmdClose.hwnd, "Click here to close this window!!!", "Close", Tip_Balloon, Tip_Error
End Sub

Private Sub cmdDelete_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ShowToolTip cmdDelete.hwnd, "Click here to Delete current entry!!!", "Delete", Tip_Balloon, Tip_Info
End Sub

Private Sub cmdError_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ShowToolTip cmdError.hwnd, "This is an Error type tooltip", "Info", Tip_Balloon, Tip_Error
End Sub

Private Sub cmdInfo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ShowToolTip cmdInfo.hwnd, "This is an Info type tooltip", "Info", Tip_Balloon, Tip_Info
End Sub

Private Sub cmdNone_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ShowToolTip cmdNone.hwnd, "This is the None type tooltip", "None", Tip_Balloon, Tip_None
End Sub

Private Sub cmdSave_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ShowToolTip cmdSave.hwnd, "Click here to Save the entry!!!", "Save", Tip_Balloon, Tip_Info
End Sub

Private Sub cmdTest_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ShowToolTip cmdTest.hwnd, txtTool, txtHead, lstStyle.ListIndex, lstType.ListIndex, TranslateColor(cmdBack.BackColor), TranslateColor(cmdText.BackColor)
End Sub

Private Sub cmdText_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ShowToolTip cmdText.hwnd, "Click here to change the tooltip textcolor!!!", "TextColor", Tip_Balloon, Tip_Info
End Sub

Private Sub cmdWarning_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ShowToolTip cmdWarning.hwnd, "This is an Warning type tooltip", "Info", Tip_Balloon, Tip_Warning
End Sub

Private Sub cmdText_Click()
    cdlg.ShowColor
    cmdText.BackColor = cdlg.Color
End Sub

Private Sub cmdBack_Click()
    cdlg.ShowColor
    cmdBack.BackColor = cdlg.Color
End Sub

Private Sub Form_Load()
    lstStyle.ListIndex = 1
    lstType.ListIndex = 2
End Sub
