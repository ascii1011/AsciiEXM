VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmBackupGui 
   Caption         =   "Form1"
   ClientHeight    =   10200
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13395
   LinkTopic       =   "Form1"
   ScaleHeight     =   10200
   ScaleWidth      =   13395
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame_Intro 
      Caption         =   "Introduction"
      Height          =   5295
      Left            =   120
      TabIndex        =   123
      Top             =   0
      Width           =   6615
      Begin VB.ListBox List4 
         Height          =   1230
         ItemData        =   "frmBackupGui.frx":0000
         Left            =   1440
         List            =   "frmBackupGui.frx":0016
         TabIndex        =   126
         Top             =   1740
         Width           =   4935
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Next -->"
         Height          =   435
         Left            =   5220
         TabIndex        =   125
         Top             =   4680
         Width           =   1155
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Load Test"
         Height          =   435
         Left            =   3840
         TabIndex        =   124
         Top             =   4680
         Width           =   1155
      End
      Begin VB.Label Label4 
         Caption         =   "Options include:"
         Height          =   255
         Left            =   180
         TabIndex        =   130
         Top             =   1740
         Width           =   1215
      End
      Begin VB.Label Label9 
         Caption         =   "AsciiEXM is a flexible application that allows you backup many aspects of Exchange 2003 via the MS ExMerge tool."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   240
         TabIndex        =   129
         Top             =   360
         Width           =   6315
      End
      Begin VB.Label Label10 
         Caption         =   "Choose an option and click 'Next' to continue."
         Height          =   255
         Left            =   240
         TabIndex        =   128
         Top             =   4200
         Width           =   3435
      End
      Begin VB.Label Label11 
         Caption         =   $"frmBackupGui.frx":00BC
         Height          =   435
         Left            =   240
         TabIndex        =   127
         Top             =   3300
         Width           =   6135
      End
   End
   Begin VB.Frame Frame_INI 
      Caption         =   "AdminEx.ini Settings"
      Height          =   5295
      Left            =   120
      TabIndex        =   91
      Top             =   240
      Width           =   6615
      Begin VB.CommandButton Command3 
         Caption         =   "..."
         Height          =   315
         Left            =   4920
         TabIndex        =   113
         Top             =   2400
         Width           =   375
      End
      Begin VB.TextBox txtExMergePath 
         Height          =   315
         Left            =   1500
         TabIndex        =   112
         Top             =   1980
         Width           =   3375
      End
      Begin VB.TextBox txtEmailServer 
         Height          =   315
         Left            =   1500
         TabIndex        =   111
         Top             =   1560
         Width           =   3375
      End
      Begin VB.TextBox txtPSTPath 
         Height          =   315
         Left            =   1500
         TabIndex        =   110
         Top             =   2400
         Width           =   3375
      End
      Begin VB.CommandButton Command2 
         Caption         =   "..."
         Height          =   315
         Left            =   4920
         TabIndex        =   109
         Top             =   1980
         Width           =   375
      End
      Begin VB.CommandButton Command12 
         Caption         =   "..."
         Height          =   315
         Left            =   4920
         TabIndex        =   108
         Top             =   2820
         Width           =   375
      End
      Begin VB.TextBox txtLogPath 
         Height          =   315
         Left            =   1500
         TabIndex        =   107
         Top             =   2820
         Width           =   3375
      End
      Begin VB.CommandButton Command14 
         Caption         =   "..."
         Height          =   315
         Left            =   4320
         TabIndex        =   106
         Top             =   900
         Width           =   375
      End
      Begin VB.TextBox txtINIPath 
         Height          =   315
         Left            =   1500
         TabIndex        =   105
         Top             =   900
         Width           =   2775
      End
      Begin VB.TextBox txtINIFileName 
         Height          =   315
         Left            =   4980
         TabIndex        =   104
         Text            =   "ExMerge.ini"
         Top             =   900
         Width           =   1275
      End
      Begin VB.CheckBox chkDumpster 
         Caption         =   "Dumpster"
         Height          =   195
         Left            =   2040
         TabIndex        =   103
         Top             =   4140
         Width           =   1215
      End
      Begin VB.CheckBox chkFolderRules 
         Caption         =   "Folder Rules"
         Height          =   195
         Left            =   3600
         TabIndex        =   102
         Top             =   3840
         Width           =   1275
      End
      Begin VB.TextBox txtRootPath 
         Height          =   315
         Left            =   1500
         TabIndex        =   101
         Top             =   360
         Width           =   2775
      End
      Begin VB.CommandButton Command17 
         Caption         =   "..."
         Height          =   315
         Left            =   4320
         TabIndex        =   100
         Top             =   360
         Width           =   375
      End
      Begin VB.CheckBox chkUseRoot 
         Caption         =   "Use"
         Height          =   195
         Left            =   4860
         TabIndex        =   99
         Top             =   420
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.CheckBox chkFolderData 
         Caption         =   "Folder Data"
         Height          =   195
         Left            =   3600
         TabIndex        =   98
         Top             =   4200
         Width           =   1275
      End
      Begin VB.CheckBox chkUserData 
         Caption         =   "User Data"
         Height          =   195
         Left            =   2040
         TabIndex        =   97
         Top             =   3840
         Width           =   1215
      End
      Begin VB.ComboBox cmbLogLevel 
         Height          =   315
         ItemData        =   "frmBackupGui.frx":015E
         Left            =   1500
         List            =   "frmBackupGui.frx":016E
         TabIndex        =   96
         Text            =   "Medium"
         Top             =   3240
         Width           =   1575
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Clear Form"
         Height          =   435
         Left            =   3360
         TabIndex        =   95
         Top             =   4680
         Width           =   1155
      End
      Begin VB.CommandButton Command15 
         Caption         =   "Next -->"
         Height          =   435
         Left            =   5220
         TabIndex        =   94
         Top             =   4680
         Width           =   1155
      End
      Begin VB.CommandButton Command16 
         Caption         =   "Use Default"
         Height          =   435
         Left            =   1980
         TabIndex        =   93
         Top             =   4680
         Width           =   1155
      End
      Begin VB.CommandButton Command26 
         Caption         =   "<-- Back"
         Height          =   435
         Left            =   240
         TabIndex        =   92
         Top             =   4680
         Width           =   1155
      End
      Begin VB.Label Label13 
         Caption         =   "Exmerge Path:"
         Height          =   195
         Left            =   300
         TabIndex        =   122
         Top             =   2040
         Width           =   1155
      End
      Begin VB.Label Label12 
         Caption         =   "Email Server:"
         Height          =   195
         Left            =   300
         TabIndex        =   121
         Top             =   1620
         Width           =   1155
      End
      Begin VB.Label Label5 
         Caption         =   "PST Path:"
         Height          =   195
         Left            =   300
         TabIndex        =   120
         Top             =   2460
         Width           =   1155
      End
      Begin VB.Label Label15 
         Caption         =   "Log Path:"
         Height          =   195
         Left            =   300
         TabIndex        =   119
         Top             =   2880
         Width           =   1155
      End
      Begin VB.Label Label17 
         Caption         =   "INI Path:"
         Height          =   195
         Left            =   300
         TabIndex        =   118
         Top             =   960
         Width           =   1155
      End
      Begin VB.Label Label18 
         Caption         =   "\"
         Height          =   255
         Left            =   4800
         TabIndex        =   117
         Top             =   960
         Width           =   135
      End
      Begin VB.Label Label16 
         Caption         =   "Root Path:"
         Height          =   195
         Left            =   300
         TabIndex        =   116
         Top             =   420
         Width           =   1155
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   6360
         Y1              =   780
         Y2              =   780
      End
      Begin VB.Line Line2 
         X1              =   120
         X2              =   6360
         Y1              =   3720
         Y2              =   3720
      End
      Begin VB.Label Label19 
         Caption         =   "Export the following:"
         Height          =   195
         Left            =   300
         TabIndex        =   115
         Top             =   3840
         Width           =   1515
      End
      Begin VB.Label Label20 
         Caption         =   "Log Level:"
         Height          =   195
         Left            =   300
         TabIndex        =   114
         Top             =   3300
         Width           =   1155
      End
   End
   Begin VB.Frame Frame_MailBoxes 
      Caption         =   "MailBox.txt Template Settings"
      Height          =   5295
      Left            =   0
      TabIndex        =   72
      Top             =   600
      Width           =   6615
      Begin VB.TextBox txtOrganization 
         Height          =   315
         Left            =   1740
         TabIndex        =   85
         Top             =   900
         Width           =   3375
      End
      Begin VB.TextBox txtCN1 
         Height          =   315
         Left            =   1740
         TabIndex        =   84
         Top             =   1860
         Width           =   3375
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Clear Form"
         Height          =   435
         Left            =   3540
         TabIndex        =   83
         Top             =   4680
         Width           =   1155
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Next -->"
         Height          =   435
         Left            =   5220
         TabIndex        =   82
         Top             =   4680
         Width           =   1155
      End
      Begin VB.CommandButton Command8 
         Caption         =   "?"
         Height          =   315
         Left            =   5280
         TabIndex        =   81
         Top             =   900
         Width           =   375
      End
      Begin VB.CommandButton Command9 
         Caption         =   "?"
         Height          =   315
         Left            =   5280
         TabIndex        =   80
         Top             =   1380
         Width           =   375
      End
      Begin VB.CommandButton Command10 
         Caption         =   "?"
         Height          =   315
         Left            =   5280
         TabIndex        =   79
         Top             =   1860
         Width           =   375
      End
      Begin VB.CommandButton Command11 
         Caption         =   "..."
         Height          =   315
         Left            =   4560
         TabIndex        =   78
         Top             =   420
         Width           =   375
      End
      Begin VB.TextBox txtMailBoxPath 
         Height          =   315
         Left            =   1740
         TabIndex        =   77
         Top             =   420
         Width           =   2775
      End
      Begin VB.CommandButton Command13 
         Caption         =   "Use Default"
         Height          =   435
         Left            =   1980
         TabIndex        =   76
         Top             =   4680
         Width           =   1155
      End
      Begin VB.CommandButton Command25 
         Caption         =   "<-- Back"
         Height          =   435
         Left            =   240
         TabIndex        =   75
         Top             =   4680
         Width           =   1155
      End
      Begin VB.TextBox txtMailBoxFileName 
         Height          =   315
         Left            =   5220
         TabIndex        =   74
         Text            =   "mailboxes.txt"
         Top             =   420
         Width           =   1275
      End
      Begin VB.TextBox txtGroup 
         Height          =   315
         Left            =   1740
         TabIndex        =   73
         Top             =   1380
         Width           =   3375
      End
      Begin VB.Label Label6 
         Caption         =   "Organization (/O=):"
         Height          =   195
         Left            =   240
         TabIndex        =   90
         Top             =   960
         Width           =   1395
      End
      Begin VB.Label Label7 
         Caption         =   "CN1 (/CN=):"
         Height          =   195
         Left            =   240
         TabIndex        =   89
         Top             =   1920
         Width           =   1395
      End
      Begin VB.Label Label8 
         Caption         =   "Group (/OU):"
         Height          =   195
         Left            =   240
         TabIndex        =   88
         Top             =   1440
         Width           =   1395
      End
      Begin VB.Label Label14 
         Caption         =   "MailBox Path:"
         Height          =   195
         Left            =   240
         TabIndex        =   87
         Top             =   480
         Width           =   1275
      End
      Begin VB.Label Label25 
         Caption         =   "\"
         Height          =   255
         Left            =   5040
         TabIndex        =   86
         Top             =   480
         Width           =   135
      End
   End
   Begin VB.Frame Frame_AdvancedOptions 
      Caption         =   "Advanced Options"
      Height          =   5295
      Left            =   4140
      TabIndex        =   46
      Top             =   2520
      Width           =   6615
      Begin VB.CommandButton Command19 
         Caption         =   "Clear Form"
         Height          =   435
         Left            =   3300
         TabIndex        =   63
         Top             =   4680
         Width           =   1155
      End
      Begin VB.CommandButton Command20 
         Caption         =   "Next -->"
         Height          =   435
         Left            =   5160
         TabIndex        =   62
         Top             =   4680
         Width           =   1155
      End
      Begin VB.CommandButton Command21 
         Caption         =   "Use Default"
         Height          =   435
         Left            =   1920
         TabIndex        =   61
         Top             =   4680
         Width           =   1155
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Group Process"
         Height          =   195
         Left            =   1800
         TabIndex        =   60
         Top             =   2460
         Width           =   1395
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Individual Process"
         Height          =   195
         Left            =   3480
         TabIndex        =   59
         Top             =   2460
         Width           =   1635
      End
      Begin VB.ComboBox cmbCompressionOption 
         Height          =   315
         ItemData        =   "frmBackupGui.frx":0192
         Left            =   1320
         List            =   "frmBackupGui.frx":01A2
         TabIndex        =   58
         Text            =   "None"
         Top             =   3300
         Width           =   2655
      End
      Begin VB.ComboBox cmbEncryptionOption 
         Height          =   315
         ItemData        =   "frmBackupGui.frx":01CC
         Left            =   1320
         List            =   "frmBackupGui.frx":01D6
         TabIndex        =   57
         Text            =   "None"
         Top             =   2880
         Width           =   2655
      End
      Begin VB.ComboBox cmbMoveToOption 
         Height          =   315
         ItemData        =   "frmBackupGui.frx":01EA
         Left            =   1380
         List            =   "frmBackupGui.frx":01FA
         TabIndex        =   56
         Text            =   "None"
         Top             =   1200
         Width           =   1815
      End
      Begin VB.TextBox txtMoveToFolder 
         Enabled         =   0   'False
         Height          =   315
         Left            =   3600
         TabIndex        =   55
         Top             =   1200
         Width           =   1815
      End
      Begin VB.CommandButton Command22 
         Caption         =   "..."
         Height          =   315
         Left            =   5520
         TabIndex        =   54
         Top             =   1200
         Width           =   375
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "frmBackupGui.frx":022E
         Left            =   1320
         List            =   "frmBackupGui.frx":023E
         TabIndex        =   53
         Text            =   "None"
         Top             =   4200
         Width           =   1815
      End
      Begin VB.CommandButton Command27 
         Caption         =   "<-- Back"
         Height          =   435
         Left            =   180
         TabIndex        =   52
         Top             =   4680
         Width           =   1155
      End
      Begin VB.CheckBox chkUseWinZip 
         Caption         =   "Use Winzip on each pst file"
         Height          =   195
         Left            =   300
         TabIndex        =   51
         Top             =   360
         Width           =   2235
      End
      Begin VB.TextBox txtDotZipFileName 
         Height          =   315
         Left            =   3780
         TabIndex        =   50
         Text            =   "MailBoxes.zip"
         Top             =   720
         Width           =   1275
      End
      Begin VB.TextBox txtDotZipFilePath 
         Height          =   315
         Left            =   300
         TabIndex        =   49
         Top             =   720
         Width           =   2775
      End
      Begin VB.CommandButton Command32 
         Caption         =   "..."
         Height          =   315
         Left            =   3120
         TabIndex        =   48
         Top             =   720
         Width           =   375
      End
      Begin VB.CheckBox chkDeleteAfterZip 
         Caption         =   "Delete PST after zipped"
         Height          =   195
         Left            =   2940
         TabIndex        =   47
         Top             =   360
         Width           =   2295
      End
      Begin VB.Label Label21 
         Caption         =   "Compression:"
         Height          =   195
         Left            =   240
         TabIndex        =   71
         Top             =   3360
         Width           =   1035
      End
      Begin VB.Label Label22 
         Caption         =   "Encryption:"
         Height          =   195
         Left            =   240
         TabIndex        =   70
         Top             =   2940
         Width           =   1035
      End
      Begin VB.Label Label23 
         Caption         =   "Move Files:"
         Height          =   195
         Left            =   300
         TabIndex        =   69
         Top             =   1260
         Width           =   1035
      End
      Begin VB.Label Label24 
         Caption         =   "MailBox Archiving:"
         Height          =   195
         Left            =   240
         TabIndex        =   68
         Top             =   2460
         Width           =   1335
      End
      Begin VB.Label Label26 
         Caption         =   "To"
         Height          =   195
         Left            =   3300
         TabIndex        =   67
         Top             =   1260
         Width           =   255
      End
      Begin VB.Label Label27 
         Caption         =   "Move:"
         Height          =   195
         Left            =   240
         TabIndex        =   66
         Top             =   4260
         Width           =   1035
      End
      Begin VB.Label Label28 
         Caption         =   "2 GB PST Limit Options:"
         Height          =   195
         Left            =   240
         TabIndex        =   65
         Top             =   3840
         Width           =   1815
      End
      Begin VB.Line Line3 
         X1              =   180
         X2              =   6420
         Y1              =   3720
         Y2              =   3720
      End
      Begin VB.Label Label36 
         Caption         =   "\"
         Height          =   255
         Left            =   3600
         TabIndex        =   64
         Top             =   780
         Width           =   135
      End
   End
   Begin VB.Frame Frame_ExmergeOptions 
      Caption         =   "ExMerge Options"
      Height          =   5235
      Left            =   9240
      TabIndex        =   39
      Top             =   6720
      Width           =   6615
      Begin VB.ComboBox cboThreadCount 
         Height          =   315
         ItemData        =   "frmBackupGui.frx":0272
         Left            =   1020
         List            =   "frmBackupGui.frx":0285
         TabIndex        =   44
         Top             =   540
         Width           =   2835
      End
      Begin VB.CheckBox chkDefaultThreads 
         Caption         =   "Use Default"
         Height          =   195
         Left            =   4020
         TabIndex        =   43
         Top             =   600
         Width           =   1215
      End
      Begin VB.CommandButton Command30 
         Caption         =   "<-- Back"
         Height          =   435
         Left            =   240
         TabIndex        =   42
         Top             =   4620
         Width           =   1155
      End
      Begin VB.CommandButton Command31 
         Caption         =   "Next -->"
         Height          =   435
         Left            =   5220
         TabIndex        =   41
         Top             =   4620
         Width           =   1155
      End
      Begin VB.CheckBox chkShowInterface 
         Caption         =   "Show ExMerge interface while processing"
         Height          =   195
         Left            =   240
         TabIndex        =   40
         Top             =   1140
         Width           =   3315
      End
      Begin VB.Label Label35 
         Caption         =   "Threads:"
         Height          =   195
         Left            =   240
         TabIndex        =   45
         Top             =   600
         Width           =   675
      End
   End
   Begin VB.Frame Frame_Accounts 
      Caption         =   "Add Users to be ExMerged"
      Height          =   5295
      Left            =   9480
      TabIndex        =   28
      Top             =   6960
      Width           =   6615
      Begin VB.TextBox txtEmailAccount 
         Height          =   315
         Left            =   1440
         TabIndex        =   36
         Top             =   360
         Width           =   1755
      End
      Begin VB.ListBox List1 
         Height          =   2205
         Left            =   300
         TabIndex        =   35
         Top             =   780
         Width           =   2895
      End
      Begin VB.CommandButton cmdAddUser 
         Caption         =   "Add"
         Height          =   315
         Left            =   3300
         TabIndex        =   34
         Top             =   360
         Width           =   915
      End
      Begin VB.CommandButton cmdClearUsers 
         Caption         =   "Clear List"
         Height          =   315
         Left            =   3300
         TabIndex        =   33
         Top             =   2700
         Width           =   915
      End
      Begin VB.CommandButton cmdRemoveUser 
         Caption         =   "Remove"
         Height          =   315
         Left            =   3300
         TabIndex        =   32
         Top             =   1560
         Width           =   915
      End
      Begin VB.CommandButton Command18 
         Caption         =   "Next -->"
         Height          =   435
         Left            =   5220
         TabIndex        =   31
         Top             =   4680
         Width           =   1155
      End
      Begin VB.CommandButton Command28 
         Caption         =   "<-- Back"
         Height          =   435
         Left            =   3780
         TabIndex        =   30
         Top             =   4680
         Width           =   1155
      End
      Begin VB.ListBox List3 
         Height          =   1815
         Left            =   180
         TabIndex        =   29
         Top             =   3360
         Width           =   3195
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   5760
         Top             =   1680
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label Label1 
         Caption         =   "Email Account:"
         Height          =   195
         Left            =   300
         TabIndex        =   38
         Top             =   420
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Drives:"
         Height          =   195
         Left            =   180
         TabIndex        =   37
         Top             =   3120
         Width           =   615
      End
   End
   Begin VB.Frame Frame_Process 
      Caption         =   "Process"
      Height          =   5295
      Left            =   9960
      TabIndex        =   12
      Top             =   7620
      Width           =   6615
      Begin VB.ListBox List2 
         Height          =   4155
         Left            =   2220
         TabIndex        =   25
         Top             =   420
         Width           =   3975
      End
      Begin VB.CommandButton cmdStartProcess 
         Caption         =   "Process"
         Height          =   435
         Left            =   3300
         TabIndex        =   24
         Top             =   4680
         Width           =   915
      End
      Begin VB.CommandButton Command23 
         Caption         =   "Halt"
         Height          =   435
         Left            =   4380
         TabIndex        =   23
         Top             =   4680
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.Frame Frame1 
         Enabled         =   0   'False
         Height          =   1815
         Left            =   120
         TabIndex        =   16
         Top             =   180
         Width           =   1995
         Begin VB.CheckBox Check15 
            Caption         =   "Processing Options"
            ForeColor       =   &H000000C0&
            Height          =   195
            Left            =   120
            TabIndex        =   20
            Top             =   840
            Width           =   1755
         End
         Begin VB.CheckBox Check13 
            Caption         =   "Compiling Details."
            ForeColor       =   &H000000C0&
            Height          =   195
            Left            =   120
            TabIndex        =   19
            Top             =   1140
            Width           =   1755
         End
         Begin VB.CheckBox Check12 
            Caption         =   "Backup Completed."
            ForeColor       =   &H000000C0&
            Height          =   195
            Left            =   120
            TabIndex        =   18
            Top             =   540
            Width           =   1755
         End
         Begin VB.CheckBox Check11 
            Caption         =   "Starting Backup."
            ForeColor       =   &H000000C0&
            Height          =   195
            Left            =   120
            TabIndex        =   17
            Top             =   240
            Width           =   1755
         End
         Begin VB.Label Label30 
            Caption         =   "Errors:"
            ForeColor       =   &H000000C0&
            Height          =   195
            Left            =   120
            TabIndex        =   22
            Top             =   1440
            Width           =   495
         End
         Begin VB.Label Label31 
            ForeColor       =   &H000000C0&
            Height          =   195
            Left            =   660
            TabIndex        =   21
            Top             =   1440
            Width           =   975
         End
      End
      Begin VB.ListBox List5 
         Height          =   2790
         Left            =   120
         TabIndex        =   15
         Top             =   2340
         Width           =   1995
      End
      Begin VB.CommandButton Command24 
         Caption         =   "Next -->"
         Height          =   435
         Left            =   5400
         TabIndex        =   14
         Top             =   4680
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton Command29 
         Caption         =   "<-- Back"
         Height          =   435
         Left            =   2220
         TabIndex        =   13
         Top             =   4680
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Output:"
         Height          =   195
         Left            =   2280
         TabIndex        =   27
         Top             =   180
         Width           =   615
      End
      Begin VB.Label Label29 
         Caption         =   "Accounts:"
         Height          =   195
         Left            =   120
         TabIndex        =   26
         Top             =   2100
         Width           =   855
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1995
      Left            =   180
      TabIndex        =   0
      Top             =   7500
      Visible         =   0   'False
      Width           =   4335
      Begin VB.CheckBox Check3 
         Caption         =   "For this session only backup certain date range."
         Height          =   195
         Left            =   180
         TabIndex        =   9
         Top             =   480
         Width           =   3855
      End
      Begin VB.TextBox Text12 
         Enabled         =   0   'False
         Height          =   315
         Left            =   2520
         TabIndex        =   8
         ToolTipText     =   "Ex. '4:00 PM'"
         Top             =   1380
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   360
         Left            =   3660
         Picture         =   "frmBackupGui.frx":0330
         ScaleHeight     =   330
         ScaleWidth      =   360
         TabIndex        =   7
         Top             =   1380
         Visible         =   0   'False
         Width           =   390
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   2040
         Picture         =   "frmBackupGui.frx":04BA
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   6
         Top             =   1380
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.TextBox Text11 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1020
         TabIndex        =   5
         Top             =   1380
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox Text10 
         Enabled         =   0   'False
         Height          =   315
         Left            =   2520
         TabIndex        =   4
         ToolTipText     =   "Ex. '4:00 PM'"
         Top             =   960
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.PictureBox Picture6 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   360
         Left            =   3660
         Picture         =   "frmBackupGui.frx":08FC
         ScaleHeight     =   330
         ScaleWidth      =   360
         TabIndex        =   3
         Top             =   960
         Visible         =   0   'False
         Width           =   390
      End
      Begin VB.PictureBox Picture4 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   2040
         Picture         =   "frmBackupGui.frx":0A86
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   2
         Top             =   960
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.TextBox Text9 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1020
         TabIndex        =   1
         Top             =   960
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label33 
         Caption         =   "End Date:"
         Enabled         =   0   'False
         Height          =   255
         Left            =   180
         TabIndex        =   11
         Top             =   1440
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.Label Label34 
         Caption         =   "Start Date:"
         Enabled         =   0   'False
         Height          =   255
         Left            =   180
         TabIndex        =   10
         Top             =   1020
         Visible         =   0   'False
         Width           =   795
      End
   End
   Begin VB.Label Label32 
      Caption         =   "winzip, compression, encrypt, break up of the pst files, move to"
      Height          =   375
      Left            =   180
      TabIndex        =   131
      Top             =   9600
      Width           =   4455
   End
End
Attribute VB_Name = "frmBackupGui"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
