VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Using the API is easy"
   ClientHeight    =   6390
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8430
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6390
   ScaleWidth      =   8430
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   6375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   11245
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   "VB Controls"
      TabPicture(0)   =   "Form1.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame5"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame4"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "Paths"
      TabPicture(1)   =   "Form1.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame6"
      Tab(1).Control(1)=   "txtPath"
      Tab(1).Control(2)=   "OptLongName"
      Tab(1).Control(3)=   "OptShortName"
      Tab(1).Control(4)=   "Frame7"
      Tab(1).Control(5)=   "Frame8"
      Tab(1).ControlCount=   6
      TabCaption(2)   =   "More"
      TabPicture(2)   =   "Form1.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame9"
      Tab(2).Control(1)=   "Frame10"
      Tab(2).Control(2)=   "Frame11"
      Tab(2).Control(3)=   "Frame12"
      Tab(2).Control(4)=   "cmdKeyCodes"
      Tab(2).ControlCount=   5
      Begin VB.CommandButton cmdKeyCodes 
         Caption         =   "KeyCode Giver"
         Height          =   375
         Left            =   -70320
         TabIndex        =   91
         Top             =   2640
         Width           =   3135
      End
      Begin VB.Frame Frame12 
         Caption         =   "Dialogs"
         Height          =   2295
         Left            =   -74760
         TabIndex        =   77
         Top             =   3600
         Width           =   7815
         Begin VB.CommandButton cmdDialogs 
            Caption         =   "About"
            Height          =   375
            Index           =   11
            Left            =   240
            TabIndex        =   90
            Top             =   1320
            Width           =   1215
         End
         Begin VB.CommandButton cmdDialogs 
            Caption         =   "Expot Favorites"
            Height          =   375
            Index           =   10
            Left            =   6000
            TabIndex        =   89
            Top             =   840
            Width           =   1575
         End
         Begin VB.CommandButton cmdDialogs 
            Caption         =   "Import Favorites"
            Height          =   375
            Index           =   9
            Left            =   6000
            TabIndex        =   88
            Top             =   360
            Width           =   1575
         End
         Begin VB.CommandButton cmdDialogs 
            Caption         =   "Add to Favorites"
            Height          =   375
            Index           =   8
            Left            =   4320
            TabIndex        =   87
            Top             =   360
            Width           =   1575
         End
         Begin VB.CommandButton cmdDialogs 
            Caption         =   "Organize Favorites"
            Height          =   375
            Index           =   7
            Left            =   4320
            TabIndex        =   86
            Top             =   840
            Width           =   1575
         End
         Begin VB.CommandButton cmdDialogs 
            Caption         =   "Browse for Folder"
            Height          =   375
            Index           =   6
            Left            =   2520
            TabIndex        =   85
            Top             =   1320
            Width           =   1575
         End
         Begin VB.TextBox txtDialogResult 
            Height          =   285
            Left            =   240
            Locked          =   -1  'True
            TabIndex        =   84
            Top             =   1800
            Width           =   7215
         End
         Begin VB.CommandButton cmdDialogs 
            Caption         =   "PageSetup"
            Height          =   375
            Index           =   5
            Left            =   2880
            TabIndex        =   83
            Top             =   840
            Width           =   1215
         End
         Begin VB.CommandButton cmdDialogs 
            Caption         =   "Print"
            Height          =   375
            Index           =   4
            Left            =   1560
            TabIndex        =   82
            Top             =   840
            Width           =   1215
         End
         Begin VB.CommandButton cmdDialogs 
            Caption         =   "Font"
            Height          =   375
            Index           =   3
            Left            =   240
            TabIndex        =   81
            Top             =   840
            Width           =   1215
         End
         Begin VB.CommandButton cmdDialogs 
            Caption         =   "Color"
            Height          =   375
            Index           =   2
            Left            =   2880
            TabIndex        =   80
            Top             =   360
            Width           =   1215
         End
         Begin VB.CommandButton cmdDialogs 
            Caption         =   "Save"
            Height          =   375
            Index           =   1
            Left            =   1560
            TabIndex        =   79
            Top             =   360
            Width           =   1215
         End
         Begin VB.CommandButton cmdDialogs 
            Caption         =   "Open"
            Height          =   375
            Index           =   0
            Left            =   240
            TabIndex        =   78
            Top             =   360
            Width           =   1215
         End
      End
      Begin VB.Frame Frame11 
         Caption         =   "Get Unique Temp FileName"
         Height          =   1575
         Left            =   -70440
         TabIndex        =   72
         Top             =   600
         Width           =   3375
         Begin VB.TextBox txtPrefix 
            Height          =   285
            Left            =   720
            MaxLength       =   2
            TabIndex        =   75
            Text            =   "bb"
            Top             =   960
            Width           =   375
         End
         Begin VB.TextBox txtTempFileName 
            Height          =   285
            Left            =   1320
            Locked          =   -1  'True
            TabIndex        =   74
            Top             =   960
            Width           =   1815
         End
         Begin VB.CommandButton cmdGetTempFile 
            Caption         =   "Generate"
            Height          =   375
            Left            =   240
            TabIndex        =   73
            Top             =   360
            Width           =   2895
         End
         Begin VB.Label Label9 
            Caption         =   "Prefix"
            Height          =   255
            Left            =   240
            TabIndex        =   76
            Top             =   1020
            Width           =   495
         End
      End
      Begin VB.Frame Frame10 
         Caption         =   "GUID Generator"
         Height          =   1575
         Left            =   -74760
         TabIndex        =   69
         Top             =   600
         Width           =   4095
         Begin VB.TextBox txtGUID 
            Height          =   285
            Left            =   240
            Locked          =   -1  'True
            TabIndex        =   71
            Top             =   960
            Width           =   3615
         End
         Begin VB.CommandButton cmdGUID 
            Caption         =   "Generate"
            Height          =   375
            Left            =   240
            TabIndex        =   70
            Top             =   360
            Width           =   3615
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "Internet Explorer"
         Height          =   975
         Left            =   -74760
         TabIndex        =   66
         Top             =   2280
         Width           =   4095
         Begin VB.CommandButton cmdIE4orGreater 
            Caption         =   "IE 4 or greater ?"
            Height          =   375
            Left            =   2040
            TabIndex        =   68
            Top             =   360
            Width           =   1815
         End
         Begin VB.CommandButton cmdIEVersion 
            Caption         =   "Get IE Version"
            Height          =   375
            Left            =   240
            TabIndex        =   67
            Top             =   360
            Width           =   1815
         End
      End
      Begin VB.Frame Frame8 
         Height          =   1215
         Left            =   -69960
         TabIndex        =   63
         Top             =   3420
         Width           =   3015
         Begin VB.Label lblUserName 
            Height          =   255
            Left            =   1320
            TabIndex        =   65
            Top             =   360
            Width           =   1575
         End
         Begin VB.Label Label4 
            Caption         =   "UserName:"
            Height          =   255
            Left            =   240
            TabIndex        =   64
            Top             =   360
            Width           =   855
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Hard Disks"
         Height          =   2775
         Left            =   -69960
         TabIndex        =   50
         Top             =   540
         Width           =   3015
         Begin VB.ComboBox Combo1 
            Height          =   315
            Left            =   1560
            Style           =   2  'Dropdown List
            TabIndex        =   51
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label lblVname 
            Height          =   255
            Left            =   1560
            TabIndex        =   62
            Top             =   840
            Width           =   1215
         End
         Begin VB.Label lblSerial 
            Height          =   255
            Left            =   1560
            TabIndex        =   61
            Top             =   1200
            Width           =   1215
         End
         Begin VB.Label lblFS 
            Height          =   255
            Left            =   1560
            TabIndex        =   60
            Top             =   1560
            Width           =   1215
         End
         Begin VB.Label Label5 
            Caption         =   "Volume Name:"
            Height          =   255
            Left            =   360
            TabIndex        =   59
            Top             =   840
            Width           =   1095
         End
         Begin VB.Label Label6 
            Caption         =   "Serial Number:"
            Height          =   255
            Left            =   360
            TabIndex        =   58
            Top             =   1200
            Width           =   1095
         End
         Begin VB.Label Label7 
            Caption         =   "File System:"
            Height          =   255
            Left            =   360
            TabIndex        =   57
            Top             =   1560
            Width           =   855
         End
         Begin VB.Label Label8 
            Caption         =   "Drive Letter:"
            Height          =   255
            Left            =   360
            TabIndex        =   56
            Top             =   420
            Width           =   1335
         End
         Begin VB.Label Label2 
            Caption         =   "Total Space:"
            Height          =   255
            Left            =   360
            TabIndex        =   55
            Top             =   1920
            Width           =   1095
         End
         Begin VB.Label Label3 
            Caption         =   "Free Space:"
            Height          =   255
            Left            =   360
            TabIndex        =   54
            Top             =   2280
            Width           =   1095
         End
         Begin VB.Label lblTotSp 
            Height          =   255
            Left            =   1560
            TabIndex        =   53
            Top             =   1920
            Width           =   1215
         End
         Begin VB.Label lblFree 
            Height          =   255
            Left            =   1560
            TabIndex        =   52
            Top             =   2280
            Width           =   1215
         End
      End
      Begin VB.OptionButton OptShortName 
         Caption         =   "Short Filename"
         Height          =   255
         Left            =   -72240
         TabIndex        =   49
         Top             =   4980
         Width           =   1935
      End
      Begin VB.OptionButton OptLongName 
         Caption         =   "Long Filename"
         Height          =   255
         Left            =   -74640
         TabIndex        =   48
         Top             =   4980
         Value           =   -1  'True
         Width           =   1935
      End
      Begin VB.TextBox txtPath 
         Height          =   375
         Left            =   -74760
         Locked          =   -1  'True
         TabIndex        =   28
         Top             =   5340
         Width           =   7815
      End
      Begin VB.Frame Frame6 
         Caption         =   "Common Paths"
         Height          =   4095
         Left            =   -74760
         TabIndex        =   27
         Top             =   540
         Width           =   4695
         Begin VB.OptionButton OptPath 
            Caption         =   "Temp"
            Height          =   255
            Index           =   18
            Left            =   2400
            TabIndex        =   47
            Top             =   3360
            Width           =   1455
         End
         Begin VB.OptionButton OptPath 
            Caption         =   "System"
            Height          =   255
            Index           =   17
            Left            =   2400
            TabIndex        =   46
            Top             =   3000
            Width           =   1455
         End
         Begin VB.OptionButton OptPath 
            Caption         =   "Windows"
            Height          =   255
            Index           =   16
            Left            =   2400
            TabIndex        =   45
            Top             =   2640
            Width           =   1455
         End
         Begin VB.OptionButton OptPath 
            Caption         =   "History"
            Height          =   255
            Index           =   15
            Left            =   2400
            TabIndex        =   44
            Top             =   2280
            Width           =   1455
         End
         Begin VB.OptionButton OptPath 
            Caption         =   "Cookies"
            Height          =   255
            Index           =   14
            Left            =   2400
            TabIndex        =   43
            Top             =   1920
            Width           =   1455
         End
         Begin VB.OptionButton OptPath 
            Caption         =   "Temporary Internet Files"
            Height          =   255
            Index           =   13
            Left            =   2400
            TabIndex        =   42
            Top             =   1560
            Width           =   2055
         End
         Begin VB.OptionButton OptPath 
            Caption         =   "PrintHood"
            Height          =   255
            Index           =   12
            Left            =   2400
            TabIndex        =   41
            Top             =   1200
            Width           =   1455
         End
         Begin VB.OptionButton OptPath 
            Caption         =   "Application Data"
            Height          =   255
            Index           =   11
            Left            =   2400
            TabIndex        =   40
            Top             =   840
            Width           =   1815
         End
         Begin VB.OptionButton OptPath 
            Caption         =   "All users\desktop"
            Height          =   255
            Index           =   10
            Left            =   2400
            TabIndex        =   39
            Top             =   480
            Width           =   1575
         End
         Begin VB.OptionButton OptPath 
            Caption         =   "Fonts"
            Height          =   255
            Index           =   9
            Left            =   240
            TabIndex        =   38
            Top             =   3720
            Width           =   1455
         End
         Begin VB.OptionButton OptPath 
            Caption         =   "Nethood"
            Height          =   255
            Index           =   8
            Left            =   240
            TabIndex        =   37
            Top             =   3360
            Width           =   1455
         End
         Begin VB.OptionButton OptPath 
            Caption         =   "StartMenu"
            Height          =   255
            Index           =   7
            Left            =   240
            TabIndex        =   36
            Top             =   3000
            Width           =   1455
         End
         Begin VB.OptionButton OptPath 
            Caption         =   "SendTo"
            Height          =   255
            Index           =   6
            Left            =   240
            TabIndex        =   35
            Top             =   2640
            Width           =   1455
         End
         Begin VB.OptionButton OptPath 
            Caption         =   "Recent"
            Height          =   255
            Index           =   5
            Left            =   240
            TabIndex        =   34
            Top             =   2280
            Width           =   1455
         End
         Begin VB.OptionButton OptPath 
            Caption         =   "Startup"
            Height          =   255
            Index           =   4
            Left            =   240
            TabIndex        =   33
            Top             =   1920
            Width           =   1455
         End
         Begin VB.OptionButton OptPath 
            Caption         =   "Favorites"
            Height          =   255
            Index           =   3
            Left            =   240
            TabIndex        =   32
            Top             =   1560
            Width           =   1455
         End
         Begin VB.OptionButton OptPath 
            Caption         =   "My Documents"
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   31
            Top             =   1200
            Width           =   1455
         End
         Begin VB.OptionButton OptPath 
            Caption         =   "StartMenu\Programs"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   30
            Top             =   840
            Width           =   1935
         End
         Begin VB.OptionButton OptPath 
            Caption         =   "Desktop"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   29
            Top             =   480
            Width           =   1455
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "TextBox"
         Height          =   2655
         Left            =   240
         TabIndex        =   21
         Top             =   540
         Width           =   2535
         Begin VB.OptionButton Option4 
            Caption         =   "UppercaseOnly"
            Height          =   255
            Left            =   360
            TabIndex        =   26
            Top             =   1440
            Width           =   1695
         End
         Begin VB.OptionButton Option3 
            Caption         =   "LowercaseOnly"
            Height          =   255
            Left            =   360
            TabIndex        =   25
            Top             =   1080
            Width           =   1695
         End
         Begin VB.OptionButton Option2 
            Caption         =   "NumberOnly"
            Height          =   255
            Left            =   360
            TabIndex        =   24
            Top             =   720
            Width           =   1695
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Standard style"
            Height          =   255
            Left            =   360
            TabIndex        =   23
            Top             =   360
            Value           =   -1  'True
            Width           =   1695
         End
         Begin VB.TextBox Text1 
            Height          =   330
            Left            =   360
            TabIndex        =   22
            Top             =   2040
            Width           =   1815
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "ListView"
         Height          =   2775
         Left            =   240
         TabIndex        =   17
         Top             =   3300
         Width           =   5175
         Begin VB.OptionButton Option5 
            Caption         =   "Standard style"
            Height          =   255
            Left            =   360
            TabIndex        =   19
            Top             =   360
            Value           =   -1  'True
            Width           =   1455
         End
         Begin VB.OptionButton Option6 
            Caption         =   "Flat Column Headers"
            Height          =   255
            Left            =   2520
            TabIndex        =   18
            Top             =   360
            Width           =   1815
         End
         Begin MSComctlLib.ListView ListView1 
            Height          =   1815
            Left            =   240
            TabIndex        =   20
            Top             =   720
            Width           =   4695
            _ExtentX        =   8281
            _ExtentY        =   3201
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   3
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Column 1"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Column 2"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Column 3"
               Object.Width           =   2540
            EndProperty
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "TreeView"
         Height          =   2775
         Left            =   5520
         TabIndex        =   12
         Top             =   3300
         Width           =   2655
         Begin VB.OptionButton Option7 
            Caption         =   "Show Tooltips"
            Height          =   255
            Left            =   360
            TabIndex        =   14
            Top             =   360
            Value           =   -1  'True
            Width           =   1575
         End
         Begin VB.OptionButton Option8 
            Caption         =   "No Tooltips"
            Height          =   255
            Left            =   360
            TabIndex        =   13
            Top             =   720
            Width           =   1695
         End
         Begin MSComctlLib.ImageList ImageList1 
            Left            =   1560
            Top             =   1440
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   16
            ImageHeight     =   16
            MaskColor       =   12632256
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   2
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Form1.frx":0054
                  Key             =   ""
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Form1.frx":01AE
                  Key             =   ""
               EndProperty
            EndProperty
         End
         Begin MSComctlLib.TreeView TreeView1 
            Height          =   1095
            Left            =   360
            TabIndex        =   15
            Top             =   1080
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   1931
            _Version        =   393217
            Indentation     =   441
            Style           =   7
            ImageList       =   "ImageList1"
            Appearance      =   1
         End
         Begin VB.Label Label1 
            Caption         =   "Move mouse over child nodes to see effect"
            Height          =   375
            Left            =   360
            TabIndex        =   16
            Top             =   2280
            Width           =   2055
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "CommandButton"
         Height          =   2655
         Left            =   2880
         TabIndex        =   6
         Top             =   540
         Width           =   2535
         Begin VB.CommandButton Command1 
            Caption         =   "Command1"
            Height          =   495
            Left            =   360
            TabIndex        =   11
            Top             =   1920
            Width           =   1815
         End
         Begin VB.OptionButton Option9 
            Caption         =   "Standard style"
            Height          =   255
            Left            =   360
            TabIndex        =   10
            Top             =   360
            Value           =   -1  'True
            Width           =   1575
         End
         Begin VB.OptionButton Option10 
            Caption         =   "Flat"
            Height          =   255
            Left            =   360
            TabIndex        =   9
            Top             =   720
            Width           =   1575
         End
         Begin VB.OptionButton Option11 
            Caption         =   "Hover"
            Height          =   255
            Left            =   360
            TabIndex        =   8
            Top             =   1080
            Width           =   1575
         End
         Begin VB.OptionButton Option12 
            Caption         =   "Thick Edge"
            Height          =   255
            Left            =   360
            TabIndex        =   7
            Top             =   1440
            Width           =   1455
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "ProgressBar"
         Height          =   2655
         Left            =   5520
         TabIndex        =   1
         Top             =   540
         Width           =   2655
         Begin VB.CommandButton Command2 
            Caption         =   "Run"
            Height          =   375
            Left            =   360
            TabIndex        =   4
            Top             =   1320
            Width           =   1935
         End
         Begin VB.Timer Timer1 
            Enabled         =   0   'False
            Interval        =   50
            Left            =   2040
            Top             =   360
         End
         Begin VB.OptionButton Option13 
            Caption         =   "Standard style"
            Height          =   255
            Left            =   360
            TabIndex        =   3
            Top             =   360
            Value           =   -1  'True
            Width           =   1695
         End
         Begin VB.OptionButton Option14 
            Caption         =   "Change Colors"
            Height          =   255
            Left            =   360
            TabIndex        =   2
            Top             =   720
            Width           =   1455
         End
         Begin MSComctlLib.ProgressBar ProgressBar1 
            Height          =   375
            Left            =   240
            TabIndex        =   5
            Top             =   1920
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   661
            _Version        =   393216
            Appearance      =   1
         End
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'see the module for comments on actual subs
Dim PBcount As Long
Private Sub cmdDialogs_Click(Index As Integer)
    Dim sfile As String, col As Long
    txtDialogResult.Backcolor = vbWhite
    txtDialogResult.Forecolor = vbBlack
    Select Case Index
        Case 0 'open
            cmnDlg.Filename = ""
            cmnDlg.Filter = "Rich text (*.rtf)|*.rtf|Plain text (*.txt)|*.txt|Word Document (*.doc)|*.doc|All Files (*.*)|*.*"
            cmnDlg.Flags = 5 'to remove the 'Readonly' checkbox
            ShowOpen
            If Len(cmnDlg.Filename) = 0 Then
                txtDialogResult.Text = "Cancel pressed"
            Else
                txtDialogResult.Text = cmnDlg.Filename
            End If
        Case 1 'save
            cmnDlg.Filename = ""
            cmnDlg.Filter = "Rich text (*.rtf)|*.rtf|Plain text (*.txt)|*.txt|Word Document (*.doc)|*.doc|All Files (*.*)|*.*"
            cmnDlg.Flags = 5
            ShowSave
            If Len(cmnDlg.Filename) = 0 Then
                txtDialogResult.Text = "Cancel pressed"
                Exit Sub
            Else
                sfile = cmnDlg.Filename
            End If
            'add extensions as necessary if the user forgot to
            Select Case cmnDlg.Filterindex
                Case 1
                    If InStr(1, sfile, ".") = 0 Then
                        sfile = sfile + ".rtf"
                    Else
                        sfile = ChangeExt(sfile, "rtf")
                    End If
                Case 2
                    If InStr(1, sfile, ".") = 0 Then
                        sfile = sfile + ".txt"
                    Else
                        sfile = ChangeExt(sfile, "txt")
                    End If
                Case 3
                    If InStr(1, sfile, ".") = 0 Then
                        sfile = sfile + ".doc"
                    Else
                        sfile = ChangeExt(sfile, "doc")
                    End If
                Case 4
                    If InStr(1, sfile, ".") = 0 Then sfile = sfile + ".txt"
            End Select
            txtDialogResult.Text = sfile
        Case 2 'color
            col = ShowColor
            If col = -1 Then
                txtDialogResult.Text = "Cancel pressed"
            Else
                txtDialogResult.Backcolor = col
            End If
        Case 3 'font
            txtDialogResult.Text = "Selected Font"
            With SelectFont
                .FontBold = txtGUID.FontBold
                .FontColor = vbBlack
                .FontName = txtGUID.FontName
                .Fontsize = txtGUID.Fontsize
                .FontItalic = txtGUID.FontItalic
                .FontStrikethru = txtGUID.FontStrikethru
                .FontUnderline = txtGUID.FontUnderline
                ShowFont
                txtDialogResult.FontName = .FontName
                txtDialogResult.Fontsize = .Fontsize
                txtDialogResult.FontBold = .FontBold
                txtDialogResult.Forecolor = .FontColor
                txtDialogResult.FontItalic = .FontItalic
                txtDialogResult.FontStrikethru = .FontStrikethru
                txtDialogResult.FontUnderline = .FontUnderline
            End With
        Case 4 'printer
            ShowPrinter
        Case 5 'PageSetup
            ShowPageSetupDlg
        Case 6 'BrowseForFolder
            sfile = BrowseForFolder("c:\", Me.hwnd, "Browse for Folder")
            If sfile = "" Then
                txtDialogResult.Text = "Cancel pressed"
            Else
                txtDialogResult.Text = sfile
            End If
        Case 7 'Organize Favorites
            DoOrganizeFavDlg Me.hwnd, SpecialFolder(6)
        Case 8 'Add Favorite
            BrowDlg.AddFavorite "http://www.planetsourcecode.com", "PSC"
        Case 9 'Import Favorite
            BrowDlg.ImportExportFavorites True, ""
        Case 10 'Export Favorite
            BrowDlg.ImportExportFavorites False, ""
        Case 11 'About
            ShellAbout Me.hwnd, App.title, "©PSST Software 2002", ByVal 0&

    End Select
    txtDialogResult.SetFocus
End Sub

Private Sub cmdGetTempFile_Click()
    Dim temp As String
    temp = GetTempFile(txtPrefix.Text)
    txtTempFileName.Text = FileOnly(temp) 'just show the name
    If FileExists(temp) Then Kill temp 'remove it - this is just a demo
    txtTempFileName.SetFocus
End Sub

Private Sub cmdGUID_Click()
    txtGUID.Text = GUIDGen
End Sub

Private Sub cmdIE4orGreater_Click()
    MsgBox IsIE4orGreater
End Sub

Private Sub cmdIEVersion_Click()
    MsgBox Get_ExplorerVersion
End Sub

Private Sub cmdKeyCodes_Click()
    frmKeyCodes.Show vbModal
End Sub

Private Sub Combo1_Click()
    Dim mSerial As Long, mname As String, fSys As String, mTotal As String, mFree As String
    GetHD Combo1.Text, mSerial, mname, fSys, mTotal, mFree
    lblVname = mname
    lblSerial = mSerial
    lblFS = fSys
    lblTotSp = mTotal
    lblFree = mFree
End Sub

Private Sub Command1_Click()
    Text1.SetFocus 'stops silly selection rectangle
End Sub

Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If mHover Then SetInitialBTStyle Command1
End Sub

Private Sub Command2_Click()
    PBcount = 0
    Timer1.Enabled = True
End Sub

Private Sub Form_Load()
    Dim z As Long, mDStr() As String
    InitCmnDlg Me.hwnd
    'get initial styles so we can reset them as necessary
    GetInitialTBStyle Text1
    GetInitialLVStyle ListView1
    GetInitialTVStyle TreeView1
    GetInitialBTStyle Command1
    'fill the treeview with some sample nodes
    TreeView1.Nodes.add , , "mNode", "Node 1", 1, 2
    TreeView1.Nodes.add "mNode", tvwChild, , "This is a very long name", 1, 2
    TreeView1.Nodes.add "mNode", tvwChild, , "This is another very long name", 1, 2
    TreeView1.Nodes.add , , , "Node 2", 1, 2
    TreeView1.Nodes(1).Expanded = True
    OptPath(0).Value = True 'get desktop directory
    mDStr = GetFixedDisks
    For z = 0 To UBound(mDStr)
        Combo1.AddItem mDStr(z)
    Next
    If Combo1.ListCount > 0 Then Combo1.ListIndex = 0
    lblUserName = UserName
End Sub

Private Sub Frame4_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If mHover Then BTFlat Command1
End Sub

Private Sub Option1_Click()
    SetInitialTBStyle Text1
    Text1.SetFocus
End Sub

Private Sub Option10_Click()
    BTFlat Command1
    mHover = False
End Sub

Private Sub Option11_Click()
    mHover = True
    BTFlat Command1
End Sub

Private Sub Option12_Click()
    BTThick Command1
    mHover = False

End Sub

Private Sub Option13_Click()
    PBcolor ProgressBar1, 12632256, 8388608
End Sub

Private Sub Option14_Click()
    PBcolor ProgressBar1, vbWhite, vbRed
End Sub

Private Sub Option2_Click()
    NumberOnly Text1
    Text1.SetFocus
End Sub
Private Sub Option3_Click()
    LowercaseOnly Text1
    Text1.SetFocus
End Sub
Private Sub Option4_Click()
    UppercaseOnly Text1
    Text1.SetFocus
End Sub

Private Sub Option5_Click()
    SetInitialLVStyle ListView1
End Sub

Private Sub Option6_Click()
    LVFlatColumnHeaders ListView1
End Sub

Private Sub Option7_Click()
    SetInitialTVStyle TreeView1
End Sub

Private Sub Option8_Click()
    TVNoTooltips TreeView1
End Sub

Private Sub Option9_Click()
    SetInitialBTStyle Command1
    mHover = False
End Sub

Private Sub OptLongName_Click()
    If txtPath.Text <> "" Then txtPath.Text = GetLongFilename(txtPath.Text)
End Sub

Private Sub OptPath_Click(Index As Integer)
    On Error Resume Next
    Select Case Index
        Case 0
            txtPath.Text = IIf(OptLongName.Value, SpecialFolder(0), GetDosPath(SpecialFolder(0))) 'Desktop
        Case 1
            txtPath.Text = IIf(OptLongName.Value, SpecialFolder(2), GetDosPath(SpecialFolder(2))) 'StartMenu\Programs
        Case 2
            txtPath.Text = IIf(OptLongName.Value, SpecialFolder(5), GetDosPath(SpecialFolder(5))) ' My Documents
        Case 3
            txtPath.Text = IIf(OptLongName.Value, SpecialFolder(6), GetDosPath(SpecialFolder(6))) ' Favorites
        Case 4
            txtPath.Text = IIf(OptLongName.Value, SpecialFolder(7), GetDosPath(SpecialFolder(7))) ' Startup
        Case 5
            txtPath.Text = IIf(OptLongName.Value, SpecialFolder(8), GetDosPath(SpecialFolder(8))) ' Recent
        Case 6
            txtPath.Text = IIf(OptLongName.Value, SpecialFolder(9), GetDosPath(SpecialFolder(9))) ' SendTo
        Case 7
            txtPath.Text = IIf(OptLongName.Value, SpecialFolder(11), GetDosPath(SpecialFolder(11))) ' StartMenu
        Case 8
            txtPath.Text = IIf(OptLongName.Value, SpecialFolder(19), GetDosPath(SpecialFolder(19))) ' Nethood
        Case 9
            txtPath.Text = IIf(OptLongName.Value, SpecialFolder(20), GetDosPath(SpecialFolder(20))) ' Fonts
        Case 10
            txtPath.Text = IIf(OptLongName.Value, SpecialFolder(25), GetDosPath(SpecialFolder(25))) ' All users\desktop
        Case 11
            txtPath.Text = IIf(OptLongName.Value, SpecialFolder(26), GetDosPath(SpecialFolder(26))) ' Application Data
        Case 12
            txtPath.Text = IIf(OptLongName.Value, SpecialFolder(27), GetDosPath(SpecialFolder(27))) ' PrintHood
        Case 13
            txtPath.Text = IIf(OptLongName.Value, SpecialFolder(32), GetDosPath(SpecialFolder(32))) ' Temporary Internet Files
        Case 14
            txtPath.Text = IIf(OptLongName.Value, SpecialFolder(33), GetDosPath(SpecialFolder(33))) ' Cookies
        Case 15
            txtPath.Text = IIf(OptLongName.Value, SpecialFolder(34), GetDosPath(SpecialFolder(34))) ' History
        Case 16
            txtPath.Text = IIf(OptLongName.Value, Winfolder, GetDosPath(Winfolder)) ' Windows
        Case 17
            txtPath.Text = IIf(OptLongName.Value, Sysfolder, GetDosPath(Sysfolder)) ' System
        Case 18
            txtPath.Text = IIf(OptLongName.Value, GetLongFilename(GetTempPathName), GetTempPathName) 'Temp
    End Select
    txtPath.SetFocus
End Sub

Private Sub OptShortName_Click()
    If txtPath.Text <> "" Then txtPath.Text = GetDosPath(txtPath.Text)
End Sub

Private Sub Timer1_Timer()
    'run the progressbar so we can see the effect
    PBcount = PBcount + 1
    If PBcount > ProgressBar1.Max Then
        Timer1.Enabled = False
        Exit Sub
    End If
    ProgressBar1.Value = PBcount
End Sub
