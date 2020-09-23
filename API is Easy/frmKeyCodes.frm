VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmKeyCodes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "KeyCodes"
   ClientHeight    =   4635
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5175
   Icon            =   "frmKeyCodes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4635
   ScaleWidth      =   5175
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   330
      Left            =   3960
      TabIndex        =   13
      Top             =   4200
      Width           =   1095
   End
   Begin MSComctlLib.ListView LV 
      Height          =   3015
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   5318
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Dec"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Hex"
         Object.Width           =   1323
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Char"
         Object.Width           =   1323
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "Code"
         Object.Width           =   1323
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   3075
      Left            =   3480
      TabIndex        =   0
      Top             =   60
      Width           =   1575
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   1080
         MaxLength       =   3
         TabIndex        =   16
         Top             =   2520
         Width           =   375
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1080
         MaxLength       =   1
         TabIndex        =   2
         Top             =   2040
         Width           =   375
      End
      Begin VB.Label Label7 
         Caption         =   "chr() number"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   2580
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "Character"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   2100
         Width           =   735
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   480
         Picture         =   "frmKeyCodes.frx":08CA
         Top             =   360
         Width           =   480
      End
      Begin VB.Label Label1 
         Caption         =   "Enter a key in a textbox to retrieve its code or select an item from the list"
         Height          =   1095
         Left            =   120
         TabIndex        =   3
         Top             =   840
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   120
      TabIndex        =   4
      Top             =   3120
      Width           =   4935
      Begin VB.TextBox txtCopy 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   3
         Left            =   4200
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   360
         Width           =   495
      End
      Begin VB.TextBox txtCopy 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   360
         Width           =   495
      End
      Begin VB.TextBox txtCopy 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   360
         Width           =   495
      End
      Begin VB.TextBox txtCopy 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   720
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Code"
         Height          =   195
         Left            =   3720
         TabIndex        =   12
         Top             =   420
         Width           =   375
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "ASCII"
         Height          =   195
         Left            =   2520
         TabIndex        =   11
         Top             =   420
         Width           =   405
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Hex"
         Height          =   195
         Left            =   1440
         TabIndex        =   10
         Top             =   420
         Width           =   285
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Chr(?)"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   420
         Width           =   420
      End
   End
End
Attribute VB_Name = "frmKeyCodes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mShift As Integer
Dim mAscii As Integer
Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim z As Long, temp As String, tmpLine As String, lItem As ListItem
    temp = OneGulp(App.Path + "\AsciiTable.txt") 'load key data
    For z = 0 To 255
        tmpLine = Split(temp, vbCrLf)(z) 'put it in the listview
        Set lItem = LV.ListItems.add(, , Trim(Split(tmpLine, Chr(9))(0)))
        lItem.SubItems(1) = Trim(Split(tmpLine, Chr(9))(1))
        lItem.SubItems(2) = Trim(Split(tmpLine, Chr(9))(2))
        If UBound(Split(tmpLine, Chr(9))) > 2 Then lItem.SubItems(3) = Trim(Split(tmpLine, Chr(9))(3))
    Next
    LV.ListItems(1).Selected = True
    txtCopy(0).Text = LV.ListItems(1).Text
    txtCopy(1).Text = LV.ListItems(1).SubItems(1)
    txtCopy(2).Text = LV.ListItems(1).SubItems(2)
    txtCopy(3).Text = LV.ListItems(1).SubItems(3)
End Sub

Private Sub LV_ItemClick(ByVal Item As MSComctlLib.ListItem)
    txtCopy(0).Text = Item.Text
    txtCopy(1).Text = Item.SubItems(1)
    txtCopy(2).Text = Item.SubItems(2)
    txtCopy(3).Text = Item.SubItems(3)
    Text1.Text = Chr(Val(Item.Text))
    Item.Selected = True
    Item.EnsureVisible
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
    mAscii = KeyAscii
End Sub

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
    LV_ItemClick LV.ListItems(mAscii + 1)
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii < 48 Or KeyAscii > 57 Then 'Numbers only - a non-API method
    If KeyAscii <> 8 Then KeyAscii = 0
End If

End Sub

Private Sub Text2_KeyUp(KeyCode As Integer, Shift As Integer)
    If Val(Text2.Text) > 255 Then Text2.Text = "255"
    LV.ListItems(Val(Text2.Text) + 1).Selected = True
    LV.SelectedItem.EnsureVisible
End Sub
