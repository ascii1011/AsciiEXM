VERSION 5.00
Begin VB.Form frmRequirements 
   Caption         =   "Requirement Management"
   ClientHeight    =   5430
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10020
   LinkTopic       =   "Form1"
   ScaleHeight     =   5430
   ScaleWidth      =   10020
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command14 
      Caption         =   "Test All"
      Height          =   315
      Left            =   3240
      TabIndex        =   31
      Top             =   3540
      Width           =   915
   End
   Begin VB.Frame Frame5 
      Caption         =   "Test"
      Height          =   3075
      Left            =   6000
      TabIndex        =   26
      Top             =   360
      Width           =   1335
      Begin VB.CommandButton Command13 
         Caption         =   "..."
         Height          =   255
         Left            =   480
         TabIndex        =   30
         Top             =   2520
         Width           =   315
      End
      Begin VB.CommandButton Command12 
         Caption         =   "..."
         Height          =   255
         Left            =   480
         TabIndex        =   29
         Top             =   1920
         Width           =   315
      End
      Begin VB.CommandButton Command11 
         Caption         =   "..."
         Height          =   255
         Left            =   480
         TabIndex        =   28
         Top             =   1380
         Width           =   315
      End
      Begin VB.CommandButton Command10 
         Caption         =   "..."
         Height          =   255
         Left            =   480
         TabIndex        =   27
         Top             =   840
         Width           =   315
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Install"
      Height          =   3075
      Left            =   4260
      TabIndex        =   25
      Top             =   360
      Width           =   1695
   End
   Begin VB.Frame Frame3 
      Caption         =   "Download"
      Height          =   3075
      Left            =   3180
      TabIndex        =   20
      Top             =   360
      Width           =   1035
      Begin VB.CommandButton Command9 
         Caption         =   "..."
         Height          =   255
         Left            =   360
         TabIndex        =   24
         Top             =   840
         Width           =   315
      End
      Begin VB.CommandButton Command8 
         Caption         =   "..."
         Height          =   255
         Left            =   360
         TabIndex        =   23
         Top             =   1380
         Width           =   315
      End
      Begin VB.CommandButton Command7 
         Caption         =   "..."
         Height          =   255
         Left            =   360
         TabIndex        =   22
         Top             =   1920
         Width           =   315
      End
      Begin VB.CommandButton Command6 
         Caption         =   "..."
         Height          =   255
         Left            =   360
         TabIndex        =   21
         Top             =   2520
         Width           =   315
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Redirect"
      Height          =   3075
      Left            =   2160
      TabIndex        =   15
      Top             =   360
      Width           =   975
      Begin VB.CommandButton Command5 
         Caption         =   "..."
         Height          =   255
         Left            =   300
         TabIndex        =   19
         Top             =   2520
         Width           =   315
      End
      Begin VB.CommandButton Command4 
         Caption         =   "..."
         Height          =   255
         Left            =   300
         TabIndex        =   18
         Top             =   1920
         Width           =   315
      End
      Begin VB.CommandButton Command3 
         Caption         =   "..."
         Height          =   255
         Left            =   300
         TabIndex        =   17
         Top             =   1380
         Width           =   315
      End
      Begin VB.CommandButton Command2 
         Caption         =   "..."
         Height          =   255
         Left            =   300
         TabIndex        =   16
         Top             =   840
         Width           =   315
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Check For the Following"
      Height          =   3075
      Left            =   120
      TabIndex        =   4
      Top             =   360
      Width           =   1995
      Begin VB.Label Label12 
         BackColor       =   &H8000000D&
         Height          =   195
         Left            =   240
         TabIndex        =   14
         Top             =   2580
         Width           =   195
      End
      Begin VB.Label Label11 
         BackColor       =   &H8000000D&
         Height          =   195
         Left            =   240
         TabIndex        =   13
         Top             =   1980
         Width           =   195
      End
      Begin VB.Label Label10 
         BackColor       =   &H8000000D&
         Height          =   195
         Left            =   240
         TabIndex        =   12
         Top             =   1440
         Width           =   195
      End
      Begin VB.Label Label9 
         BackColor       =   &H8000000D&
         Height          =   195
         Left            =   240
         TabIndex        =   11
         Top             =   900
         Width           =   195
      End
      Begin VB.Label Label8 
         BackColor       =   &H8000000D&
         Height          =   195
         Left            =   240
         TabIndex        =   10
         Top             =   360
         Width           =   195
      End
      Begin VB.Label Label6 
         Caption         =   "ExMerge"
         Height          =   195
         Left            =   540
         TabIndex        =   9
         Top             =   2580
         Width           =   1275
      End
      Begin VB.Label Label5 
         Caption         =   "Exchange Tools"
         Height          =   195
         Left            =   540
         TabIndex        =   8
         Top             =   1980
         Width           =   1275
      End
      Begin VB.Label Label4 
         Caption         =   "System Tools"
         Height          =   195
         Left            =   540
         TabIndex        =   7
         Top             =   1440
         Width           =   1275
      End
      Begin VB.Label Label3 
         Caption         =   "IIS"
         Height          =   195
         Left            =   540
         TabIndex        =   6
         Top             =   900
         Width           =   1275
      End
      Begin VB.Label Label1 
         Caption         =   "OS Version:"
         Height          =   195
         Left            =   540
         TabIndex        =   5
         Top             =   360
         Width           =   1275
      End
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Exchange 2007"
      Height          =   195
      Left            =   3420
      TabIndex        =   3
      Top             =   60
      Value           =   -1  'True
      Width           =   1515
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Exchange 2003"
      Enabled         =   0   'False
      Height          =   195
      Left            =   1680
      TabIndex        =   1
      Top             =   60
      Width           =   1515
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Test All"
      Height          =   315
      Left            =   6300
      TabIndex        =   0
      Top             =   3540
      Width           =   915
   End
   Begin VB.Label Label2 
      Caption         =   "Exchange Version"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   60
      Width           =   1395
   End
End
Attribute VB_Name = "frmRequirements"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command2_Click()
    
    'comDialog "All Files .* |*.*"
End Function

Sub Init()
    SysInfo
    
    Req.OS.CurrentAccount = sUser
    Req.OS.ComputerName = sComputer
    Req.OS.OS = sOS
    Req.OS.Build = sBuild
    Req.OS.Version_Major = sVMajor
    Req.OS.Version_Minor = sVMinor
    Req.OS.Version = sVersion
    Req.OS.RootDir = sGRootDir

End Sub

Private Sub Form_Load()
    Init
End Sub
