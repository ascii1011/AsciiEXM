VERSION 5.00
Begin VB.Form frmExMergeV1 
   Caption         =   "Form1"
   ClientHeight    =   10965
   ClientLeft      =   420
   ClientTop       =   435
   ClientWidth     =   17340
   LinkTopic       =   "Form1"
   ScaleHeight     =   10965
   ScaleWidth      =   17340
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame_Process 
      Caption         =   "Process"
      Height          =   5295
      Left            =   12660
      TabIndex        =   14
      Top             =   6360
      Width           =   6615
      Begin VB.CommandButton Command29 
         Caption         =   "<-- Back"
         Height          =   435
         Left            =   2220
         TabIndex        =   101
         Top             =   4680
         Width           =   975
      End
      Begin VB.CommandButton Command24 
         Caption         =   "Next -->"
         Height          =   435
         Left            =   5400
         TabIndex        =   94
         Top             =   4680
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.ListBox List5 
         Height          =   2790
         Left            =   120
         TabIndex        =   92
         Top             =   2340
         Width           =   1995
      End
      Begin VB.Frame Frame1 
         Enabled         =   0   'False
         Height          =   1815
         Left            =   120
         TabIndex        =   87
         Top             =   180
         Width           =   1995
         Begin VB.CheckBox Check11 
            Caption         =   "Starting Backup."
            ForeColor       =   &H000000C0&
            Height          =   195
            Left            =   120
            TabIndex        =   91
            Top             =   240
            Width           =   1755
         End
         Begin VB.CheckBox Check12 
            Caption         =   "Backup Completed."
            ForeColor       =   &H000000C0&
            Height          =   195
            Left            =   120
            TabIndex        =   90
            Top             =   540
            Width           =   1755
         End
         Begin VB.CheckBox Check13 
            Caption         =   "Compiling Details."
            ForeColor       =   &H000000C0&
            Height          =   195
            Left            =   120
            TabIndex        =   89
            Top             =   1140
            Width           =   1755
         End
         Begin VB.CheckBox Check15 
            Caption         =   "Processing Options"
            ForeColor       =   &H000000C0&
            Height          =   195
            Left            =   120
            TabIndex        =   88
            Top             =   840
            Width           =   1755
         End
         Begin VB.Label Label31 
            ForeColor       =   &H000000C0&
            Height          =   195
            Left            =   660
            TabIndex        =   96
            Top             =   1440
            Width           =   975
         End
         Begin VB.Label Label30 
            Caption         =   "Errors:"
            ForeColor       =   &H000000C0&
            Height          =   195
            Left            =   120
            TabIndex        =   95
            Top             =   1440
            Width           =   495
         End
      End
      Begin VB.CommandButton Command23 
         Caption         =   "Halt"
         Height          =   435
         Left            =   4380
         TabIndex        =   85
         Top             =   4680
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.CommandButton cmdStartProcess 
         Caption         =   "Process"
         Height          =   435
         Left            =   3300
         TabIndex        =   6
         Top             =   4680
         Width           =   915
      End
      Begin VB.ListBox List2 
         Height          =   4155
         Left            =   2220
         TabIndex        =   7
         Top             =   420
         Width           =   3975
      End
      Begin VB.Label Label29 
         Caption         =   "Accounts:"
         Height          =   195
         Left            =   120
         TabIndex        =   93
         Top             =   2100
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Output:"
         Height          =   195
         Left            =   2280
         TabIndex        =   15
         Top             =   180
         Width           =   615
      End
   End
   Begin VB.Frame Frame_Accounts 
      Caption         =   "Add Users to be ExMerged"
      Height          =   5295
      Left            =   6420
      TabIndex        =   12
      Top             =   4620
      Width           =   6615
      Begin VB.ComboBox cboMailBoxStores 
         Height          =   315
         Left            =   2880
         TabIndex        =   125
         Top             =   720
         Width           =   2535
      End
      Begin VB.ComboBox cboStorageGroups 
         Height          =   315
         Left            =   120
         TabIndex        =   124
         Top             =   720
         Width           =   2415
      End
      Begin VB.CommandButton Command33 
         Caption         =   ">"
         Height          =   315
         Left            =   3720
         TabIndex        =   121
         Top             =   2280
         Width           =   435
      End
      Begin VB.ListBox List7 
         Height          =   2205
         Left            =   120
         TabIndex        =   119
         Top             =   1740
         Width           =   3555
      End
      Begin VB.CommandButton Command28 
         Caption         =   "<-- Back"
         Height          =   435
         Left            =   3780
         TabIndex        =   100
         Top             =   4680
         Width           =   1155
      End
      Begin VB.CommandButton Command18 
         Caption         =   "Next -->"
         Height          =   435
         Left            =   5220
         TabIndex        =   62
         Top             =   4680
         Width           =   1155
      End
      Begin VB.CommandButton cmdRemoveUser 
         Caption         =   "Remove"
         Height          =   315
         Left            =   5520
         TabIndex        =   51
         Top             =   2520
         Width           =   915
      End
      Begin VB.CommandButton cmdClearUsers 
         Caption         =   "Clear List"
         Height          =   315
         Left            =   5520
         TabIndex        =   5
         Top             =   3660
         Width           =   915
      End
      Begin VB.CommandButton cmdAddUser 
         Caption         =   "Add"
         Height          =   315
         Left            =   5520
         TabIndex        =   3
         Top             =   1320
         Width           =   915
      End
      Begin VB.ListBox List1 
         Height          =   2205
         Left            =   4200
         TabIndex        =   4
         Top             =   1740
         Width           =   1215
      End
      Begin VB.TextBox txtEmailAccount 
         Height          =   315
         Left            =   4020
         TabIndex        =   2
         Top             =   1320
         Width           =   1395
      End
      Begin VB.Label Label37 
         Caption         =   "Storage Groups:"
         Height          =   195
         Left            =   120
         TabIndex        =   127
         Top             =   480
         Width           =   1635
      End
      Begin VB.Label Label34 
         Caption         =   "MailBox Databases:"
         Height          =   195
         Left            =   2880
         TabIndex        =   126
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label Label33 
         Caption         =   "Available Emails:"
         Height          =   195
         Left            =   120
         TabIndex        =   120
         Top             =   1500
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Email Account:"
         Height          =   195
         Left            =   2880
         TabIndex        =   13
         Top             =   1380
         Width           =   1095
      End
   End
   Begin VB.Frame Frame_ExmergeOptions 
      Caption         =   "ExMerge Options"
      Height          =   5235
      Left            =   2520
      TabIndex        =   103
      Top             =   4980
      Width           =   6615
      Begin VB.ListBox List3 
         Height          =   2400
         Left            =   180
         TabIndex        =   122
         Top             =   1860
         Width           =   5835
      End
      Begin VB.CheckBox chkShowInterface 
         Caption         =   "Show ExMerge interface while processing"
         Height          =   195
         Left            =   240
         TabIndex        =   109
         Top             =   1140
         Width           =   3315
      End
      Begin VB.CommandButton Command31 
         Caption         =   "Next -->"
         Height          =   435
         Left            =   5220
         TabIndex        =   108
         Top             =   4620
         Width           =   1155
      End
      Begin VB.CommandButton Command30 
         Caption         =   "<-- Back"
         Height          =   435
         Left            =   240
         TabIndex        =   107
         Top             =   4620
         Width           =   1155
      End
      Begin VB.CheckBox chkDefaultThreads 
         Caption         =   "Use Default"
         Height          =   195
         Left            =   4020
         TabIndex        =   106
         Top             =   600
         Width           =   1215
      End
      Begin VB.ComboBox cboThreadCount 
         Height          =   315
         ItemData        =   "frmExMergeV1.frx":0000
         Left            =   1020
         List            =   "frmExMergeV1.frx":0013
         TabIndex        =   105
         Top             =   540
         Width           =   2835
      End
      Begin VB.Label Label3 
         Caption         =   "Drives:"
         Height          =   195
         Left            =   180
         TabIndex        =   123
         Top             =   1620
         Width           =   615
      End
      Begin VB.Label Label35 
         Caption         =   "Threads:"
         Height          =   195
         Left            =   240
         TabIndex        =   104
         Top             =   600
         Width           =   675
      End
   End
   Begin VB.Frame Frame_AdvancedOptions 
      Caption         =   "Advanced Options"
      Height          =   5295
      Left            =   420
      TabIndex        =   63
      Top             =   4560
      Width           =   6615
      Begin VB.CheckBox chkDeleteAfterZip 
         Caption         =   "Delete PST after zipped"
         Height          =   195
         Left            =   2940
         TabIndex        =   115
         Top             =   360
         Width           =   2295
      End
      Begin VB.CommandButton Command32 
         Caption         =   "..."
         Height          =   315
         Left            =   3120
         TabIndex        =   113
         Top             =   720
         Width           =   375
      End
      Begin VB.TextBox txtDotZipFilePath 
         Height          =   315
         Left            =   300
         TabIndex        =   112
         Top             =   720
         Width           =   2775
      End
      Begin VB.TextBox txtDotZipFileName 
         Height          =   315
         Left            =   3780
         TabIndex        =   111
         Text            =   "MailBoxes.zip"
         Top             =   720
         Width           =   1275
      End
      Begin VB.CheckBox chkUseWinZip 
         Caption         =   "Use Winzip on each pst file"
         Height          =   195
         Left            =   300
         TabIndex        =   110
         Top             =   360
         Width           =   2235
      End
      Begin VB.CommandButton Command27 
         Caption         =   "<-- Back"
         Height          =   435
         Left            =   180
         TabIndex        =   99
         Top             =   4680
         Width           =   1155
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "frmExMergeV1.frx":00BE
         Left            =   1320
         List            =   "frmExMergeV1.frx":00CE
         TabIndex        =   83
         Text            =   "None"
         Top             =   4200
         Width           =   1815
      End
      Begin VB.CommandButton Command22 
         Caption         =   "..."
         Height          =   315
         Left            =   5520
         TabIndex        =   80
         Top             =   1200
         Width           =   375
      End
      Begin VB.TextBox txtMoveToFolder 
         Enabled         =   0   'False
         Height          =   315
         Left            =   3600
         TabIndex        =   79
         Top             =   1200
         Width           =   1815
      End
      Begin VB.ComboBox cmbMoveToOption 
         Height          =   315
         ItemData        =   "frmExMergeV1.frx":0102
         Left            =   1380
         List            =   "frmExMergeV1.frx":0112
         TabIndex        =   77
         Text            =   "None"
         Top             =   1200
         Width           =   1815
      End
      Begin VB.ComboBox cmbEncryptionOption 
         Height          =   315
         ItemData        =   "frmExMergeV1.frx":0146
         Left            =   1320
         List            =   "frmExMergeV1.frx":0150
         TabIndex        =   75
         Text            =   "None"
         Top             =   2880
         Width           =   2655
      End
      Begin VB.ComboBox cmbCompressionOption 
         Height          =   315
         ItemData        =   "frmExMergeV1.frx":0164
         Left            =   1320
         List            =   "frmExMergeV1.frx":0174
         TabIndex        =   73
         Text            =   "None"
         Top             =   3300
         Width           =   2655
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Individual Process"
         Height          =   195
         Left            =   3480
         TabIndex        =   71
         Top             =   2460
         Width           =   1635
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Group Process"
         Height          =   195
         Left            =   1800
         TabIndex        =   70
         Top             =   2460
         Width           =   1395
      End
      Begin VB.CommandButton Command21 
         Caption         =   "Use Default"
         Height          =   435
         Left            =   1920
         TabIndex        =   69
         Top             =   4680
         Width           =   1155
      End
      Begin VB.CommandButton Command20 
         Caption         =   "Next -->"
         Height          =   435
         Left            =   5160
         TabIndex        =   68
         Top             =   4680
         Width           =   1155
      End
      Begin VB.CommandButton Command19 
         Caption         =   "Clear Form"
         Height          =   435
         Left            =   3300
         TabIndex        =   67
         Top             =   4680
         Width           =   1155
      End
      Begin VB.Label Label36 
         Caption         =   "\"
         Height          =   255
         Left            =   3600
         TabIndex        =   114
         Top             =   780
         Width           =   135
      End
      Begin VB.Line Line3 
         X1              =   180
         X2              =   6420
         Y1              =   3720
         Y2              =   3720
      End
      Begin VB.Label Label28 
         Caption         =   "2 GB PST Limit Options:"
         Height          =   195
         Left            =   240
         TabIndex        =   84
         Top             =   3840
         Width           =   1815
      End
      Begin VB.Label Label27 
         Caption         =   "Move:"
         Height          =   195
         Left            =   240
         TabIndex        =   82
         Top             =   4260
         Width           =   1035
      End
      Begin VB.Label Label26 
         Caption         =   "To"
         Height          =   195
         Left            =   3300
         TabIndex        =   81
         Top             =   1260
         Width           =   255
      End
      Begin VB.Label Label24 
         Caption         =   "MailBox Archiving:"
         Height          =   195
         Left            =   240
         TabIndex        =   78
         Top             =   2460
         Width           =   1335
      End
      Begin VB.Label Label23 
         Caption         =   "Move Files:"
         Height          =   195
         Left            =   300
         TabIndex        =   76
         Top             =   1260
         Width           =   1035
      End
      Begin VB.Label Label22 
         Caption         =   "Encryption:"
         Height          =   195
         Left            =   240
         TabIndex        =   74
         Top             =   2940
         Width           =   1035
      End
      Begin VB.Label Label21 
         Caption         =   "Compression:"
         Height          =   195
         Left            =   240
         TabIndex        =   72
         Top             =   3360
         Width           =   1035
      End
   End
   Begin VB.Frame Frame_MailBoxes 
      Caption         =   "MailBox.txt Template Settings"
      Height          =   5295
      Left            =   7740
      TabIndex        =   8
      Top             =   2100
      Width           =   6615
      Begin VB.ListBox List6 
         Height          =   2205
         Left            =   240
         TabIndex        =   116
         Top             =   2280
         Width           =   6075
      End
      Begin VB.TextBox txtGroup 
         Height          =   315
         Left            =   1740
         TabIndex        =   102
         Top             =   1380
         Width           =   3375
      End
      Begin VB.TextBox txtMailBoxFileName 
         Height          =   315
         Left            =   5220
         TabIndex        =   24
         Text            =   "mailboxes.txt"
         Top             =   420
         Width           =   1275
      End
      Begin VB.CommandButton Command25 
         Caption         =   "<-- Back"
         Height          =   435
         Left            =   240
         TabIndex        =   97
         Top             =   4680
         Width           =   1155
      End
      Begin VB.CommandButton Command13 
         Caption         =   "Use Default"
         Height          =   435
         Left            =   1980
         TabIndex        =   52
         Top             =   4680
         Width           =   1155
      End
      Begin VB.TextBox txtMailBoxPath 
         Height          =   315
         Left            =   1740
         TabIndex        =   22
         Top             =   420
         Width           =   2775
      End
      Begin VB.CommandButton Command11 
         Caption         =   "..."
         Height          =   315
         Left            =   4560
         TabIndex        =   23
         Top             =   420
         Width           =   375
      End
      Begin VB.CommandButton Command10 
         Caption         =   "?"
         Height          =   315
         Left            =   5280
         TabIndex        =   30
         Top             =   1860
         Width           =   375
      End
      Begin VB.CommandButton Command9 
         Caption         =   "?"
         Height          =   315
         Left            =   5280
         TabIndex        =   29
         Top             =   1380
         Width           =   375
      End
      Begin VB.CommandButton Command8 
         Caption         =   "?"
         Height          =   315
         Left            =   5280
         TabIndex        =   28
         Top             =   900
         Width           =   375
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Next -->"
         Height          =   435
         Left            =   5220
         TabIndex        =   27
         Top             =   4680
         Width           =   1155
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Clear Form"
         Height          =   435
         Left            =   3540
         TabIndex        =   1
         Top             =   4680
         Width           =   1155
      End
      Begin VB.TextBox txtCN1 
         Height          =   315
         Left            =   1740
         TabIndex        =   0
         Top             =   1860
         Width           =   3375
      End
      Begin VB.TextBox txtOrganization 
         Height          =   315
         Left            =   1740
         TabIndex        =   26
         Top             =   900
         Width           =   3375
      End
      Begin VB.Label Label25 
         Caption         =   "\"
         Height          =   255
         Left            =   5040
         TabIndex        =   43
         Top             =   480
         Width           =   135
      End
      Begin VB.Label Label14 
         Caption         =   "MailBox Path:"
         Height          =   195
         Left            =   240
         TabIndex        =   42
         Top             =   480
         Width           =   1275
      End
      Begin VB.Label Label8 
         Caption         =   "Group (/OU):"
         Height          =   195
         Left            =   240
         TabIndex        =   11
         Top             =   1440
         Width           =   1395
      End
      Begin VB.Label Label7 
         Caption         =   "CN1 (/CN=):"
         Height          =   195
         Left            =   240
         TabIndex        =   10
         Top             =   1920
         Width           =   1395
      End
      Begin VB.Label Label6 
         Caption         =   "Organization (/O=):"
         Height          =   195
         Left            =   240
         TabIndex        =   9
         Top             =   960
         Width           =   1395
      End
   End
   Begin VB.Frame Frame_INI 
      Caption         =   "AdminEx.ini Settings"
      Height          =   5295
      Left            =   180
      TabIndex        =   31
      Top             =   540
      Width           =   6615
      Begin VB.ComboBox cboEmailServers 
         Height          =   315
         Left            =   1500
         TabIndex        =   117
         Top             =   1560
         Width           =   3375
      End
      Begin VB.CommandButton Command26 
         Caption         =   "<-- Back"
         Height          =   435
         Left            =   240
         TabIndex        =   98
         Top             =   4680
         Width           =   1155
      End
      Begin VB.CommandButton Command16 
         Caption         =   "Use Default"
         Height          =   435
         Left            =   1980
         TabIndex        =   66
         Top             =   4680
         Width           =   1155
      End
      Begin VB.CommandButton Command15 
         Caption         =   "Next -->"
         Height          =   435
         Left            =   5220
         TabIndex        =   65
         Top             =   4680
         Width           =   1155
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Clear Form"
         Height          =   435
         Left            =   3360
         TabIndex        =   64
         Top             =   4680
         Width           =   1155
      End
      Begin VB.ComboBox cmbLogLevel 
         Height          =   315
         ItemData        =   "frmExMergeV1.frx":019E
         Left            =   1500
         List            =   "frmExMergeV1.frx":01AE
         TabIndex        =   61
         Text            =   "Medium"
         Top             =   3240
         Width           =   1575
      End
      Begin VB.CheckBox chkUserData 
         Caption         =   "User Data"
         Height          =   195
         Left            =   2040
         TabIndex        =   58
         Top             =   3840
         Width           =   1215
      End
      Begin VB.CheckBox chkFolderData 
         Caption         =   "Folder Data"
         Height          =   195
         Left            =   3600
         TabIndex        =   57
         Top             =   4200
         Width           =   1275
      End
      Begin VB.CheckBox chkUseRoot 
         Caption         =   "Use"
         Height          =   195
         Left            =   4860
         TabIndex        =   56
         Top             =   420
         Width           =   675
      End
      Begin VB.CommandButton Command17 
         Caption         =   "..."
         Height          =   315
         Left            =   4320
         TabIndex        =   55
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox txtRootPath 
         Height          =   315
         Left            =   1500
         TabIndex        =   53
         Top             =   360
         Width           =   2775
      End
      Begin VB.CheckBox chkFolderRules 
         Caption         =   "Folder Rules"
         Height          =   195
         Left            =   3600
         TabIndex        =   50
         Top             =   3840
         Width           =   1275
      End
      Begin VB.CheckBox chkDumpster 
         Caption         =   "Dumpster"
         Height          =   195
         Left            =   2040
         TabIndex        =   49
         Top             =   4200
         Width           =   1215
      End
      Begin VB.TextBox txtINIFileName 
         Height          =   315
         Left            =   4980
         TabIndex        =   47
         Text            =   "ExMerge.ini"
         Top             =   900
         Width           =   1275
      End
      Begin VB.TextBox txtINIPath 
         Height          =   315
         Left            =   1500
         TabIndex        =   45
         Top             =   900
         Width           =   2775
      End
      Begin VB.CommandButton Command14 
         Caption         =   "..."
         Height          =   315
         Left            =   4320
         TabIndex        =   44
         Top             =   900
         Width           =   375
      End
      Begin VB.TextBox txtLogPath 
         Height          =   315
         Left            =   1500
         TabIndex        =   40
         Top             =   2820
         Width           =   3375
      End
      Begin VB.CommandButton Command12 
         Caption         =   "..."
         Height          =   315
         Left            =   4920
         TabIndex        =   39
         Top             =   2820
         Width           =   375
      End
      Begin VB.CommandButton Command2 
         Caption         =   "..."
         Height          =   315
         Left            =   4920
         TabIndex        =   38
         Top             =   1980
         Width           =   375
      End
      Begin VB.TextBox txtPSTPath 
         Height          =   315
         Left            =   1500
         TabIndex        =   34
         Top             =   2400
         Width           =   3375
      End
      Begin VB.TextBox txtExMergePath 
         Height          =   315
         Left            =   1500
         TabIndex        =   33
         Top             =   1980
         Width           =   3375
      End
      Begin VB.CommandButton Command3 
         Caption         =   "..."
         Height          =   315
         Left            =   4920
         TabIndex        =   32
         Top             =   2400
         Width           =   375
      End
      Begin VB.Label Label20 
         Caption         =   "Log Level:"
         Height          =   195
         Left            =   300
         TabIndex        =   60
         Top             =   3300
         Width           =   1155
      End
      Begin VB.Label Label19 
         Caption         =   "Export the following:"
         Height          =   195
         Left            =   300
         TabIndex        =   59
         Top             =   3840
         Width           =   1515
      End
      Begin VB.Line Line2 
         X1              =   120
         X2              =   6360
         Y1              =   3720
         Y2              =   3720
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   6360
         Y1              =   780
         Y2              =   780
      End
      Begin VB.Label Label16 
         Caption         =   "Root Path:"
         Height          =   195
         Left            =   300
         TabIndex        =   54
         Top             =   420
         Width           =   1155
      End
      Begin VB.Label Label18 
         Caption         =   "\"
         Height          =   255
         Left            =   4800
         TabIndex        =   48
         Top             =   960
         Width           =   135
      End
      Begin VB.Label Label17 
         Caption         =   "INI Path:"
         Height          =   195
         Left            =   300
         TabIndex        =   46
         Top             =   960
         Width           =   1155
      End
      Begin VB.Label Label15 
         Caption         =   "Log Path:"
         Height          =   195
         Left            =   300
         TabIndex        =   41
         Top             =   2880
         Width           =   1155
      End
      Begin VB.Label Label5 
         Caption         =   "PST Path:"
         Height          =   195
         Left            =   300
         TabIndex        =   37
         Top             =   2460
         Width           =   1155
      End
      Begin VB.Label Label12 
         Caption         =   "Email Server:"
         Height          =   195
         Left            =   300
         TabIndex        =   36
         Top             =   1620
         Width           =   1155
      End
      Begin VB.Label Label13 
         Caption         =   "Exmerge Path:"
         Height          =   195
         Left            =   300
         TabIndex        =   35
         Top             =   2040
         Width           =   1155
      End
   End
   Begin VB.Frame Frame_Intro 
      Caption         =   "Introduction"
      Height          =   5295
      Left            =   60
      TabIndex        =   16
      Top             =   60
      Width           =   6615
      Begin VB.CommandButton Command7 
         Caption         =   "Load Test"
         Height          =   435
         Left            =   3840
         TabIndex        =   86
         Top             =   4680
         Width           =   1155
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Next -->"
         Height          =   435
         Left            =   5220
         TabIndex        =   18
         Top             =   4680
         Width           =   1155
      End
      Begin VB.ListBox List4 
         Height          =   1230
         ItemData        =   "frmExMergeV1.frx":01D2
         Left            =   1440
         List            =   "frmExMergeV1.frx":01E8
         TabIndex        =   17
         Top             =   1740
         Width           =   4935
      End
      Begin VB.Label Label32 
         Height          =   255
         Left            =   240
         TabIndex        =   118
         Top             =   4800
         Width           =   3255
      End
      Begin VB.Label Label11 
         Caption         =   $"frmExMergeV1.frx":028E
         Height          =   435
         Left            =   240
         TabIndex        =   25
         Top             =   3300
         Width           =   6135
      End
      Begin VB.Label Label10 
         Caption         =   "Choose an option and click 'Next' to continue."
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   4200
         Width           =   3435
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
         TabIndex        =   20
         Top             =   360
         Width           =   6315
      End
      Begin VB.Label Label4 
         Caption         =   "Options include:"
         Height          =   255
         Left            =   180
         TabIndex        =   19
         Top             =   1740
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmExMergeV1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboEmailServers_Change()
'    GetStorageGroups Trim(cboEmailServers.Text), cboStorageGroups
    ''''''Main.Current.Server = Trim(cboEmailServers.Text)
    Show_StorageGroups Trim(cboEmailServers.Text)
End Sub

Private Sub cboEmailServers_Click()
'    GetStorageGroups Trim(cboEmailServers.Text), cboStorageGroups
    '''''Main.Current.Server = Trim(cboEmailServers.Text)
    Show_StorageGroups Trim(cboEmailServers.Text)
End Sub

Private Sub cboMailBoxStores_Change()
    'GetMailBoxes Trim(cboEmailServers.Text), Trim(cboStorageGroups.Text), Trim(cboMailBoxStores.Text), cboMailBoxStores
    Show_MailBoxes Trim(cboMailBoxStores.Text)

End Sub

Private Sub cboMailBoxStores_Click()
    'GetMailBoxes Trim(cboEmailServers.Text), Trim(cboStorageGroups.Text), Trim(cboMailBoxStores.Text), cboMailBoxStores
    Show_MailBoxes Trim(cboMailBoxStores.Text)
End Sub

Private Sub cboStorageGroups_Change()
    'GetMailBoxStores Trim(cboEmailServers.Text), Trim(cboStorageGroups.Text), cboMailBoxStores
    Show_MailBoxStoreDBs Trim(cboStorageGroups.Text)
End Sub

Private Sub cboStorageGroups_Click()
    'GetMailBoxStores Trim(cboEmailServers.Text), Trim(cboStorageGroups.Text), cboMailBoxStores
    Show_MailBoxStoreDBs Trim(cboStorageGroups.Text)
End Sub

Private Sub chkDefaultThreads_Click()
    ThreadCount
End Sub

Private Sub chkDefaultThreads_KeyUp(KeyCode As Integer, Shift As Integer)
    ThreadCount
End Sub

'search for exchange servers -> databases -> mailboxes/groups
'possibly use CSVDE -f list.csv to find the servers -> then exchange and Domain Controllers.
'if not apart of the domain search for computers with ldap traffic.
'if DC found then run "LDAP://Server/RootDSE" to start a dig for info.

'verify all tools, components, apps, and exmerge exist

'create fields ini/txt/log/etc files

'create variable structure

'add zip feature with/verifying that winzip32 is installed - zipping available

'drive discovery with size - functionality available

'calculations for drive vs accounts being backed up plus zip file

'a move to function, if the drive will get too full

'mailbox prediscovery - size

'2gb file limit solution
    'get total size of mailbox
    'get create date of mailbox
        'option 1:
            'take the total amount of time and divide by the size of mailbox in GBs
            'then do a back up of each segment of time seperately
            'this will attempt to average a 1 GB file per segment
        'option 2:
            'start with start date and do (for example) 3 month backups to the same pst
            'file and check size after each pass, when becomes greater than 1GB then
            'create then next file to continue the backup process.  keep repeating until mailbox


Private Sub cmdClearUsers_Click()
    List1.Clear
    'ValidateForm
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdStartProcess_Click()
    StartProcess
End Sub

Sub StartProcess()
    If funCreate_Mailboxes_txt(List1) Then
        
        If funCreateExMergeINI Then
            ProcessBackup
            
            If chkUseWinZip.Value = 1 Then
                ZipAllFiles
            End If
        Else
            List2.AddItem "An error has occured while creating the ExMergeX.ini file."
        End If
        
    Else
        List2.AddItem "An error has occured while creating the MailBoxes.txt file."
    End If
End Sub

Sub ZipAllFiles()
    Dim i As Integer
    Dim src As String, dest As String
    
    src = Trim(txtDotZipFilePath.Text) & "\" & Trim(txtDotZipFileName.Text)
    
    For i = 0 To List1.ListCount
        dest = Trim(txtPSTPath.Text) & "\" & Trim(List1.List(i)) & ".pst"
        ZipAFile src, dest
        If chkDeleteAfterZip.Value = 1 Then 'delete pst file
    Next i
End Sub

Private Sub cmdAddUser_Click()
    AddUserToList
    ThreadCount
End Sub

Sub ThreadCount()
    If chkDefaultThreads.Value = 1 Then
        If List1.ListCount < 5 Then
            cboThreadCount.ListIndex = 0
        ElseIf List1.ListCount < 25 Then
            cboThreadCount.ListIndex = 1
        ElseIf List1.ListCount < 50 Then
            cboThreadCount.ListIndex = 2
        ElseIf List1.ListCount < 100 Then
            cboThreadCount.ListIndex = 3
        ElseIf List1.ListCount >= 100 Then
            cboThreadCount.ListIndex = 4
        End If
    End If
End Sub

Private Sub cmdRemoveUser_Click()
    List1.RemoveItem List1.Selected
    'ValidateForm
End Sub

Private Sub Command1_Click()
    MailBox_ClearForm
End Sub

'Private Sub cmdResults_Click()
'    Call Shell("notepad " & Trim(frmExMerge.txtPSTPath.Text) & "\AdminEx.Log", vbNormalFocus)
'End Sub


Private Sub Command11_Click()
    txtMailBoxPath.Text = SHFolder
End Sub

Private Sub Command12_Click()
    txtLogPath.Text = SHFolder
End Sub

Private Sub Command13_Click()
    MailBox_Default
End Sub

Private Sub Command14_Click()
    txtINIPath.Text = SHFolder
End Sub





Private Sub Command15_Click()
    ShowFrame 2
End Sub

Private Sub Command16_Click()
    INI_Default
End Sub

Private Sub Command17_Click()
    txtRootPath.Text = SHFolder
End Sub

Private Sub Command18_Click()
    ShowFrame 6
End Sub

Private Sub Command2_Click()
    txtExMergePath.Text = SHFolder
End Sub

Private Sub Command20_Click()
    ShowFrame 4
End Sub

Private Sub Command22_Click()
    txtMoveToFolder.Text = SHFolder
End Sub

Private Sub Command25_Click()
    ShowFrame 1
End Sub

Private Sub Command26_Click()
    ShowFrame 0
End Sub

Private Sub Command27_Click()
   ShowFrame 2
End Sub

Private Sub Command28_Click()
   ShowFrame 4
End Sub

Private Sub Command29_Click()
   ShowFrame 5
End Sub

Private Sub Command3_Click()
    txtPSTPath.Text = SHFolder
End Sub



Private Sub Command30_Click()
    ShowFrame 3
End Sub

Private Sub Command31_Click()
    ShowFrame 5
End Sub

Private Sub Command32_Click()
    txtDotZipFilePath.Text = SHFolder
End Sub

Private Sub Command4_Click()
    INI_ClearForm
End Sub

Private Sub Command5_Click()
    ShowFrame 1
End Sub

Private Sub Command6_Click()
    ShowFrame 4
End Sub

Private Sub Command7_Click()
    LoadDefault
    'ValidateForm
End Sub

Private Sub Form_Load()
    Me.Width = 6930
    Me.Height = 5970
    Init
    ShowFrame 0
    Me.Show
    
    
    LoadAll
    
    
End Sub

Sub LoadAll()

    'do these later
    '''Check Resources (cdo, systools, wmi, winzip, etc)
    '''CheckResources
    
    'normally used
    'LoadDefaultSettings
    'LoadResources
    
    'Grab All Exchange Information
    GetAllExchangeInfo
    
    'Fake_ExchangeInfo
    
    Show_Servers
    
End Sub

Sub Fake_ExchangeInfo()
    ReDim Main.Exch.Svrs(3)
    Main.Exch.Svrs(0).name = "SUSAMAIL"
    Main.Exch.Svrs(1).name = "NASDATA"
    Main.Exch.Svrs(2).name = "SNYEDGE"
        
        ReDim Main.Exch.Svrs(0).SG(2)
        Main.Exch.Svrs(0).SG(0).name = "First"
            ReDim Main.Exch.Svrs(0).SG(0).MBSDB(3)
                Main.Exch.Svrs(0).SG(0).MBSDB(0).name = "acc"
                    ReDim Main.Exch.Svrs(0).SG(0).MBSDB(0).MBX(2)
                        Main.Exch.Svrs(0).SG(0).MBSDB(0).MBX(0).Alias = "alon"
                        Main.Exch.Svrs(0).SG(0).MBSDB(0).MBX(1).Alias = "bvon"
                Main.Exch.Svrs(0).SG(0).MBSDB(1).name = "exec"
                Main.Exch.Svrs(0).SG(0).MBSDB(2).name = "MBX"
            
        Main.Exch.Svrs(0).SG(1).name = "Second"
            ReDim Main.Exch.Svrs(0).SG(1).MBSDB(3)
                Main.Exch.Svrs(0).SG(1).MBSDB(0).name = "meat"
                Main.Exch.Svrs(0).SG(1).MBSDB(1).name = "meat2"
                Main.Exch.Svrs(0).SG(1).MBSDB(2).name = "meat3"
                    ReDim Main.Exch.Svrs(0).SG(1).MBSDB(1).MBX(2)
                        Main.Exch.Svrs(0).SG(1).MBSDB(1).MBX(0).Alias = "2alon"
                        Main.Exch.Svrs(0).SG(1).MBSDB(1).MBX(1).Alias = "2bvon"
            
            
        ReDim Main.Exch.Svrs(1).SG(3)
        Main.Exch.Svrs(1).SG(0).name = "nas-SG1"
        Main.Exch.Svrs(1).SG(1).name = "nas-SG2"
        Main.Exch.Svrs(1).SG(2).name = "nas-SG3"
            
            
        ReDim Main.Exch.Svrs(2).SG(3)
        Main.Exch.Svrs(2).SG(0).name = "edge-SG1"
        Main.Exch.Svrs(2).SG(1).name = "edge-SG1"
        Main.Exch.Svrs(2).SG(2).name = "edge-SG1"
End Sub

Sub Show_Servers()
    Dim iObj As Integer
    
    cboEmailServers.Clear
    cboStorageGroups.Clear
    cboMailBoxStores.Clear
    List7.Clear
    For iObj = 0 To UBound(Main.Exch.Svrs) - 1
    
        If iObj = 0 Then
            Main.Current.Server = Main.Exch.Svrs(iObj).name
            cboEmailServers.Text = Main.Exch.Svrs(iObj).name
        End If
            
        cboEmailServers.AddItem Main.Exch.Svrs(iObj).name
    
    Next iObj
End Sub

Function GetServerIndex(sTarget As String) As Integer
    Dim i As Integer
    
    GetServerIndex = -i
    
    For i = 0 To UBound(Main.Exch.Svrs) - 1
        If Main.Exch.Svrs(i).name = sTarget Then
            GetServerIndex = i
            Exit Function
        End If
    Next i
End Function

Sub Show_StorageGroups(sTarget As String)
    Dim iObj As Integer, iSvr As Integer
    
    Main.Current.Server = GetServerIndex(sTarget)
    
    cboStorageGroups.Clear
    cboMailBoxStores.Clear
    List7.Clear
    For iObj = 0 To UBound(Main.Exch.Svrs(Main.Current.Server).SG) - 1
                
        If iObj = 0 Then
            cboStorageGroups.Text = Main.Exch.Svrs(Main.Current.Server).SG(iObj).name
            'Main.Current.StorageGroup = Main.Exch.Svrs(Main.Current.Server).SG(iObj).name
        End If
            
        cboStorageGroups.AddItem Main.Exch.Svrs(Main.Current.Server).SG(iObj).name
    
    Next iObj
End Sub

Function GetStorageGroupIndex(sTarget As String) As Integer
    Dim i As Integer
    
    GetStorageGroupIndex = -i
    
    For i = 0 To UBound(Main.Exch.Svrs(Main.Current.Server).SG) - 1
        If Main.Exch.Svrs(Main.Current.Server).SG(i).name = sTarget Then
            GetStorageGroupIndex = i
            Exit Function
        End If
    Next i
End Function

Sub Show_MailBoxStoreDBs(sTarget As String)
    Dim iObj As Integer
    
    Main.Current.StorageGroup = GetStorageGroupIndex(sTarget)
    
On Error GoTo Err:
    
    With Main.Exch.Svrs(Main.Current.Server).SG(Main.Current.StorageGroup)
    
        cboMailBoxStores.Clear
        List7.Clear
        
        For iObj = 0 To UBound(.MBSDB) - 1
        
            If iObj = 0 Then
                cboMailBoxStores.Text = .MBSDB(iObj).name
                'Main.Current.MailBoxStoreDBs = .MBSDB(iObj).name
            End If
                
            cboMailBoxStores.AddItem .MBSDB(iObj).name
        
        Next iObj
    End With
    
    Exit Sub
Err:
    Exit Sub
End Sub

Function GetMailBoxStoreDBIndex(sTarget As String) As Integer
    Dim i As Integer
    
    GetMailBoxStoreDBIndex = -i
    
    For i = 0 To UBound(Main.Exch.Svrs(Main.Current.Server).SG(Main.Current.StorageGroup).MBSDB) - 1
        If Main.Exch.Svrs(Main.Current.Server).SG(Main.Current.StorageGroup).MBSDB(i).name = sTarget Then
            GetMailBoxStoreDBIndex = i
            Exit Function
        End If
    Next i
End Function

Sub Show_MailBoxes(sTarget As String)
    Dim iObj As Integer
    
    Main.Current.MailBoxStoreDBs = GetMailBoxStoreDBIndex(sTarget)
    
On Error GoTo Err:
    
    With Main.Exch.Svrs(Main.Current.Server).SG(Main.Current.StorageGroup).MBSDB(Main.Current.MailBoxStoreDBs)
    
        List7.Clear
        For iObj = 0 To UBound(.MBX) - 1
                    
            List7.AddItem .MBX(iObj).Alias
        
        Next iObj
    End With
    
    Exit Sub
Err:
    Exit Sub
End Sub

Sub LoadResources()
    ''' get Exchange Server list
    'GetExchangeServers
    
    Label32.Caption = "loading Drives..."
    Pause 1
    mod_Drives.Drives2Struct
    
    Label32.Caption = "loading Servers..."
    Pause 1
    GetExchangeServers2 cboEmailServers
    
    Label32.Caption = ""
End Sub

Sub CheckResources()
    Label32.Caption = "Checking Winzip..."
    Check_Winzip
End Sub
Sub Check_Winzip()
    chkUseWinZip.Enabled = False
    chkUseWinZip.Value = 0
    chkDeleteAfterZip.Enabled = False
    chkDeleteAfterZip.Value = 0
    txtDotZipFilePath.Enabled = False
    Command32.Enabled = False
    txtDotZipFilePath.Enabled = False
    If Exists_fso("") Then
        chkUseWinZip.Enabled = True
        chkDeleteAfterZip.Enabled = True
        txtDotZipFilePath.Enabled = True
        Command32.Enabled = True
        txtDotZipFilePath.Enabled = True
    End If
End Sub

Sub LoadDefaultSettings()
    MailBox_Default
    ExMergeOptions_Default
    Account_Default
    INI_Default
End Sub

Sub Init()

    Me.Frame_Intro.Top = 60
    Me.Frame_Intro.Left = 100
    Me.Frame_INI.Top = 50
    Me.Frame_INI.Left = 100
    Me.Frame_MailBoxes.Top = 50
    Me.Frame_MailBoxes.Left = 100
    Me.Frame_AdvancedOptions.Top = 50
    Me.Frame_AdvancedOptions.Left = 100
    Me.Frame_ExmergeOptions.Top = 50
    Me.Frame_ExmergeOptions.Left = 100
    Me.Frame_Accounts.Top = 50
    Me.Frame_Accounts.Left = 100
    Me.Frame_Process.Top = 50
    Me.Frame_Process.Left = 100
    
End Sub

Sub ShowFrame(iFrame)
    
    If iFrame = 0 Then
        Frame_Intro.ZOrder vbBringToFront
    ElseIf iFrame = 1 Then
        Frame_INI.ZOrder vbBringToFront
    ElseIf iFrame = 2 Then
        Frame_MailBoxes.ZOrder vbBringToFront
    ElseIf iFrame = 3 Then
        Frame_AdvancedOptions.ZOrder vbBringToFront
    ElseIf iFrame = 4 Then
        Frame_ExmergeOptions.ZOrder vbBringToFront
    ElseIf iFrame = 5 Then
        Frame_Accounts.ZOrder vbBringToFront
    ElseIf iFrame = 6 Then
        Frame_Process.ZOrder vbBringToFront
    End If
    
End Sub


'example: sFilter = "All Files .* |*.*"
Function funFileChoice(cdg As CommonDialog, sFilter As String) As String
    cdg.Filter = sFilter
    cdg.ShowOpen
    
    If cdg.filename = "" Then Exit Function
    
    funFileChoice = cdg.filename
End Function


'example: sFilter = "All Files .* |*.*"
Function funFolderChoice(cdg As CommonDialog, sFilter As String) As String
    cdg.Filter = sFilter
    cdg.ShowOpen
    
    If cdg.filename = "" Then Exit Function
    
    funFileChoice = cdg.filename
End Function


Sub LoadTestVariables()
    With Me
        '.cboEmailServers.Text = "susamail"
        .txtExMergePath.Text = "c:\program files\exchsrvr\bin"
        .txtPSTPath.Text = "c:\chris"
        .txtOrganization.Text = "SIBONEYUSA CORP"
        .txtGroup.Text = "EXCHANGE ADMINISTRATIVE GROUP (FYDIBOHF23SPDLT)"
        .txtCN1.Text = "RECIPIENTS"
        .List1.AddItem "NGUPTA"
    End With
End Sub

Sub Account_Default()
    With Me
        .List1.AddItem "NGUPTA"
        .List1.AddItem "CCANTO"
        .List1.AddItem "BFORD"
        .List1.AddItem "AMINOGUE"
    End With
End Sub

Sub Account_ClearForm()
    With Me
        .List1.Clear
        .txtEmailAccount.Text = ""
    End With
End Sub


Sub INI_ClearForm()
    With Me
        .txtRootPath.Text = ""
        .chkUseRoot.Value = 0
        .txtINIPath.Text = ""
        .txtINIFileName.Text = ""
        '.cboEmailServers.Text = ""
        .txtExMergePath.Text = ""
        .txtPSTPath.Text = ""
        .txtLogPath.Text = ""
        .cmbLogLevel.ListIndex = 0
        .chkUserData.Value = 0
        .chkDumpster.Value = 0
        .chkFolderRules.Value = 0
        .chkFolderData.Value = 0
    End With
End Sub

Sub ExMergeOptions_Default()
    With Me
        .cboThreadCount.ListIndex = 0
        .chkShowInterface.Value = 1
    End With
End Sub

Sub INI_Default()
    With Me
        .txtINIPath.Text = "C:\backup\exchange"
        .txtINIFileName.Text = "ExMerge.ini"
        '.cboEmailServers.Text = "ServerName"
        .txtExMergePath.Text = "c:\program files\exchsrvr\bin"
        .txtPSTPath.Text = "C:\backup\exchange\PST"
        .txtLogPath.Text = "C:\backup\exchange\Log"
        .cmbLogLevel.ListIndex = 2
        .chkUserData.Value = 0
        .chkDumpster.Value = 1
        .chkFolderRules.Value = 0
        .chkFolderData.Value = 0
        .chkUseRoot.Value = 1
        .txtRootPath.Text = "C:\chris"
    End With
End Sub
Sub MailBox_ClearForm()
    With Me
        .txtMailBoxPath.Text = ""
        .txtMailBoxFileName.Text = ""
        .txtOrganization.Text = ""
        .txtGroup.Text = ""
        .txtCN1.Text = ""
    End With
End Sub
Sub MailBox_Default()
    With Me
        .txtMailBoxPath.Text = "C:\backup\exchange"
        .txtMailBoxFileName.Text = "mailboxes.txt"
        .txtOrganization.Text = Get_OrganizationName '"SIBONEYUSA CORP"
        .txtGroup.Text = "EXCHANGE ADMINISTRATIVE GROUP (FYDIBOHF23SPDLT)"
        .txtCN1.Text = "RECIPIENTS"
    End With
End Sub



Function UsersAvailable() As Boolean
    UsersAvailable = False
    
    If List1.ListCount > 0 Then
        UsersAvailable = True
        Me.cmdRemoveUser.Enabled = True
    Else
        Me.cmdRemoveUser.Enabled = False
    End If
        
End Function


Sub ValidateForm()
    Dim checksum As Integer
    Dim FieldColor_Empty, FieldColor_Filled
    
    FieldColor_Empty = &H80000000
    FieldColor_Filled = vbWhite
    
    
    ' reset all variables
    Me.cmdStartProcess.Enabled = False
    
    Me.cboEmailServers.BackColor = FieldColor_Filled
    Me.txtExMergePath.BackColor = FieldColor_Filled
    Me.txtPSTPath.BackColor = FieldColor_Filled
    Me.txtOrganization.BackColor = FieldColor_Filled
    Me.txtGroup.BackColor = FieldColor_Filled
    Me.txtCN1.BackColor = FieldColor_Filled
    Me.txtEmailAccount.BackColor = FieldColor_Filled
        
    ' start checking if they are ok or not
    If UsersAvailable = True Then
        checksum = checksum + 1
    Else
        Me.txtEmailAccount.BackColor = FieldColor_Empty
    End If
    ' cboEmailServers
    If Trim(Me.cboEmailServers.Text) <> "" Then
        checksum = checksum + 1
    Else
        Me.cboEmailServers.BackColor = FieldColor_Empty
    End If
    ' txtExMergePath
    If Trim(Me.txtExMergePath.Text) <> "" Then
        checksum = checksum + 1
    Else
        Me.txtExMergePath.BackColor = FieldColor_Empty
    End If
    ' txtPSTPath
    If Trim(Me.txtPSTPath.Text) <> "" Then
        checksum = checksum + 1
    Else
        Me.txtPSTPath.BackColor = FieldColor_Empty
    End If
    ' txtOrganization
    If Trim(Me.txtOrganization.Text) <> "" Then
        checksum = checksum + 1
    Else
        Me.txtOrganization.BackColor = FieldColor_Empty
    End If
    ' txtGroup
    If Trim(Me.txtGroup.Text) <> "" Then
        checksum = checksum + 1
    Else
        Me.txtGroup.BackColor = FieldColor_Empty
    End If
    ' txtCN1
    If Trim(Me.txtCN1.Text) <> "" Then
        checksum = checksum + 1
    Else
        Me.txtCN1.BackColor = FieldColor_Empty
    End If
    
    ' if file exists then enable button
    ' Me.cmdResults.Enabled = False
    ' If FileExists(Trim(frmExMerge.txtPSTPath.Text) & "\AdminEx.Log") Then Me.cmdResults.Enabled = True
        
    ' if process ok enable the start button
    If checksum = 7 Then Me.cmdStartProcess.Enabled = True
End Sub









Private Sub List1_Click()
    If List1.ListCount > 0 Then cmdRemoveUser.Enabled = True
End Sub

Private Sub List1_LostFocus()
    cmdRemoveUser.Enabled = False
End Sub

Private Sub txtCN1_Change()
    'ValidateForm
End Sub

Private Sub txtEmailAccount_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        AddUserToList
    End If
End Sub

Sub AddUserToList()
    If Trim(Me.txtEmailAccount.Text) <> "" Then
        List1.AddItem Trim(Me.txtEmailAccount.Text)
    End If
    txtEmailAccount.Text = ""
End Sub


Private Sub txtExMergePath_Change()
    'ValidateForm
End Sub

Private Sub txtGroup_Change()
    'ValidateForm
End Sub

Private Sub txtPSTPath_Change()
    'ValidateForm
End Sub






Public Function funCreate_Mailboxes_txt(lst As ListBox) As Boolean
    funCreate_Mailboxes_txt = funCreatefile(CreateMailBoxTxTPath, SetupMailBoxTxt(lst))
    
End Function

Function CreateMailBoxTxTPath() As String
    CreateMailBoxTxTPath = Trim(txtMailBoxPath.Text) & "\" & Trim(txtMailBoxFileName.Text)
    
End Function

Function SetupMailBoxTxt(lst As ListBox) As String
    Dim i As Integer
    Dim sTemplate As String, sContent As String
    
    sTemplate = "/o=" & Trim(txtOrganization.Text) & _
                "/ou=" & Trim(txtGroup.Text) & _
                "/cn=" & Trim(txtCN1.Text) & _
                "/cn="
    
    sContent = "##~This file was generated by AsciiEXM for use with ExMerge.exe" & vbNewLine
    
    For i = 0 To lst.ListCount - 1
        sContent = sContent & sTemplate & lst.List(i) & vbNewLine
    Next i
    
    SetupMailBoxTxt = sContent
End Function

Function funCreateExMergeINI() As Boolean
    funCreateExMergeINI = funCreatefile(CreateExMergeINIPath, SetupExMergeInI)
    
End Function

Function CreateExMergeINIPath() As String
    CreateExMergeINIPath = Trim(txtINIPath.Text) & "\" & Trim(txtINIFileName.Text)

End Function

Function SetupExMergeInI() As String
    Dim sBody As String
    
    sBody = "; EXEMERGE.INI; This file is for use with the EXMERGE.EXE program, version 4.00 or higher." & vbNewLine
    sBody = sBody & "[EXMERGE]" & vbNewLine
    sBody = sBody & "MergeAction=0" & vbNewLine
    sBody = sBody & "SourceServerName=" & Trim(cboEmailServers.Text) & vbNewLine
    sBody = sBody & "DataDirectoryName=" & Trim(txtPSTPath.Text) & vbNewLine
    sBody = sBody & "FileContainingListOfMailboxes=" & Trim(txtMailBoxPath.Text) & "\mailboxes.txt" & vbNewLine
    'sBody = sBody & "FileContainingListOfDatabases=C:\Program Files\Exchsrvr\bin\PRIVMDBS.TXT" & vbNewLine
    sBody = sBody & "LogFileName=" & Trim(txtLogPath.Text) & "\ExMergeX.Log" & vbNewLine
    sBody = sBody & "LoggingLevel=" & cmbLogLevel.ListIndex & vbNewLine
    If chkDumpster.Value = 1 Then sBody = sBody & "CopyDeletedItemsFromDumpster = 1" & vbNewLine
    If chkFolderRules.Value = 1 Then sBody = sBody & "CopyFolderPermissions = 1" & vbNewLine
    If chkFolderData.Value = 1 Then sBody = sBody & "CopyAssociatedFolderData = 1" & vbNewLine
    If chkUserData.Value = 1 Then sBody = sBody & "CopyUserData = 1" & vbNewLine
    
    SetupExMergeInI = sBody

End Function

Sub INIParameters()
'   MergeAction
'       0 = specify SourceServerName
'       1 = specify DestServerName
'       2 = specify both
'
'   SourceServername =
'
'   DestServerName =
'
'   DomainControllerForSourceServer =
'
'   SrcServerLDAP-Port = 389
'
'   DomainControllerForDestServer =
'
'   DestServerLDAP-Port =
'
'   LogFileName = file.log
'
'   LoggingLevel
'       0 = None
'       1 = minimum
'       2 = medium
'       3 = maximum
'
'   DataDirectoryName
'       default = c:\EXMERGEDATA
'       where the .pst files are written or exported to.
'
'   FileContainingListOfMailboxes
'       any file, default = mailboxes.txt
'
'   FileContainingListOfDatabases
'       if listed will override FileContainingListOfMailboxes
'
'   SelectMessageStartDate = 1/1/98 00:00:00
'
'   SelectMessageEndDate = 12/31/01 23:59:59
'
'       when using SelectMessageStartDate and SelectMessageEndDate
'       the PR_DELIVERY_TIME MAPI attribute is used.
'
'       to use the modified time (aka PR_MODIFIED_TIME MAPI attribute)
'       use the DateAttribute setting in the ini file to 1
'
'   DataImportMethod = 3
'       0 = copy all messages from the source store to the target store
'       1 = merge messages into the target store.  copy only those messages
'           that do not exist in the target store.
'       2 = replace existing message in the target store. (if a message
'           in the source store exists in the target store, delete that
'           message in the target store and then copy the message from
'           the target store.)
'       3 = archive existing messages from the source store into the
'           target store.  if this option is selected, the program copies data from the source store to the target store and then deletes the data from the source store.  this option is valid only if the mergeaction is extract.
'       use this setting to archive data from the source
'       WARNING - this will delete all messages/items that are exported to the destination
'
'   ReplaceDataOnlyIfSourceItemIsMoreRecent
'       0 = replace all in the target store
'       1 = replace anly items in the target store, if the source store has more recent version
'       this is applicable only if DataImportMethod = 3(replace data)
'       This setting uses PR_LAST_MODIFICATION_TIME when checking datetime
'
'   CopyFolderPermissions = 1
'       using this will copy all folder permissions from each mailbox
'
'   CopyAssociatedFolderData = 1
'
'   CopyUserData
'       0 = do not copy user data (msg, folders, cal, contacts)
'       1 = copy all
'
'   CopyDeletedItemsFromDumpster
'       0 = do not copy items from the dumpster
'       1 = copy items from the dumpster
'       default = 0
'       all dumpster data will be retrieved regardless of date selections
'       all dumpster data recovered will put placed in the Deleted Items folder
'
'
'   -NUMTHREADS #
'       max 10 threads
'
'       boxes < 5 = 1 thread
'       boxes < 25 = 2 threads
'       boxes < 50 = 3 threads
'       boxes < 100 = 4 threads
'       boxes >= 100 = 5 threads
'
'       use of threads is cautioned and is urged to use on a dedicated workstation
'       resources can increase exponentially as the thread count increases.
'
'
'

End Sub

Function ProcessBackup() As Boolean
    Dim sPath As String
    ProcessBackup = False
    
    ChDrive "c:"
    ChDir (Trim(txtExMergePath.Text))
    
    sPath = CreateExMergeString
    ProcessBackup = ShellAndWait(sPath, vbNormalFocus)
End Function

Function CreateExMergeString() As String
    Dim sPath As String
    
    'Create InI Full Path
    sPath = Trim(txtINIPath.Text) & "\" & Trim(txtINIFileName.Text)
    
    'Create Exmerge Shell cmd
    sPath = "exmerge -F " & sPath & " -B "
    
    'Show ExMerge interface
    If chkShowInterface.Value = 1 Then
        sPath = sPath & "-D "
    End If
    
    'Add custom number of threads
    If chkDefaultThreads.Value = 0 Then
        sPath = sPath & "-NUMTHREADS " & cboThreadCount.ListIndex + 1
    End If
    
    Debug.Print "----"
    Debug.Print sPath
    MsgBox sPath
    CreateExMergeString = sPath
    
End Function

Private Sub txtRootPath_Change()
    If chkUseRoot.Value = 1 Then prcUpdateAllForms
End Sub

Sub prcUpdateAllForms()
    With Me
        .txtMailBoxPath.Text = Trim(txtRootPath.Text)
        .txtINIPath.Text = Trim(txtRootPath.Text)
        .txtPSTPath.Text = Trim(txtRootPath.Text) & "\PST"
        .txtLogPath.Text = Trim(txtRootPath.Text) & "\LOG"
    End With
End Sub
