VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmConvert 
   Caption         =   "Any-to-Any Database converter"
   ClientHeight    =   3915
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5910
   Icon            =   "frmConvert.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3915
   ScaleWidth      =   5910
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame5 
      Height          =   1215
      Left            =   4080
      TabIndex        =   27
      Top             =   6300
      Width           =   7755
      Begin VB.CommandButton cmdClose 
         Caption         =   "Close"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6360
         TabIndex        =   29
         Top             =   720
         Width           =   1215
      End
      Begin VB.CommandButton cmdCreate 
         Caption         =   "Create"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6360
         TabIndex        =   28
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame4 
      Height          =   1215
      Left            =   120
      TabIndex        =   23
      Top             =   6300
      Width           =   3915
      Begin VB.CheckBox Check2 
         Caption         =   "Save Job details automatically"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   195
         Left            =   180
         TabIndex        =   26
         Top             =   840
         Width           =   2955
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Verify Database after creation"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   240
         Left            =   180
         TabIndex        =   25
         Top             =   540
         Width           =   2895
      End
      Begin VB.CheckBox chkBypass 
         Caption         =   "Bypass the rules"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   240
         Left            =   180
         TabIndex        =   24
         Top             =   240
         Width           =   1755
      End
   End
   Begin VB.Frame Frame3 
      Height          =   735
      Left            =   120
      TabIndex        =   17
      Top             =   5520
      Width           =   11715
      Begin VB.TextBox txtDestDatabaseFile 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   7140
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   240
         Width           =   3795
      End
      Begin VB.CommandButton cmdDestBrowse 
         Caption         =   "..."
         Height          =   360
         Left            =   11040
         TabIndex        =   20
         Top             =   240
         Width           =   435
      End
      Begin VB.ComboBox cmbDestDatabase 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2520
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   240
         Width           =   2115
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Dest. Database File"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   4920
         TabIndex        =   22
         Top             =   300
         Width           =   2085
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Destination Database"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   180
         TabIndex        =   18
         Top             =   300
         Width           =   2265
      End
   End
   Begin VB.Frame Frame2 
      Height          =   2055
      Left            =   120
      TabIndex        =   12
      Top             =   3420
      Width           =   3915
      Begin VB.CommandButton cmdSortFieldsRevAlpha 
         Caption         =   "R"
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
         Left            =   3480
         TabIndex        =   36
         ToolTipText     =   "Sort Reverse Alphabet order"
         Top             =   1140
         Width           =   255
      End
      Begin VB.CommandButton cmdSortFieldsAlpha 
         Caption         =   "A"
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
         Left            =   3480
         TabIndex        =   35
         ToolTipText     =   "Sort Alphabetically"
         Top             =   900
         Width           =   255
      End
      Begin VB.CommandButton cmdOrgOrder 
         Caption         =   "O"
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
         Left            =   3480
         TabIndex        =   34
         ToolTipText     =   "Original Order"
         Top             =   660
         Width           =   255
      End
      Begin VB.CommandButton cmdFieldsDelete 
         Caption         =   "X"
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
         Left            =   3480
         TabIndex        =   16
         ToolTipText     =   "Delete Selected"
         Top             =   420
         Width           =   255
      End
      Begin VB.ListBox lstSelectedFields 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1500
         ItemData        =   "frmConvert.frx":014A
         Left            =   120
         List            =   "frmConvert.frx":014C
         TabIndex        =   13
         Top             =   420
         Width           =   3315
      End
      Begin MSComCtl2.UpDown udFieldsUpDown 
         Height          =   495
         Left            =   3480
         TabIndex        =   15
         ToolTipText     =   "Re-order the fields"
         Top             =   1440
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   873
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Selected Field(s)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   180
         Width           =   1440
      End
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   2760
      Top             =   1740
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame fraFields 
      Height          =   2535
      Left            =   4080
      TabIndex        =   7
      Top             =   840
      Width           =   7755
      Begin VB.CommandButton cmdSelectNone 
         Caption         =   "-"
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
         Left            =   7440
         TabIndex        =   33
         ToolTipText     =   "Select None"
         Top             =   780
         Width           =   255
      End
      Begin VB.CommandButton cmdSelectAll 
         Caption         =   "+"
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
         Left            =   7440
         TabIndex        =   32
         ToolTipText     =   "Select All"
         Top             =   480
         Width           =   255
      End
      Begin MSComctlLib.ListView lvFields 
         Height          =   2010
         Left            =   120
         TabIndex        =   9
         Top             =   420
         Width           =   7275
         _ExtentX        =   12832
         _ExtentY        =   3545
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin VB.Label lblTotal 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   5280
         TabIndex        =   31
         Top             =   180
         Width           =   75
      End
      Begin VB.Label lblTotalRecords 
         AutoSize        =   -1  'True
         Caption         =   "Number of Records found :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   2940
         TabIndex        =   30
         Top             =   180
         Width           =   2310
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Field(s) Available"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   180
         TabIndex        =   11
         Top             =   180
         Width           =   1470
      End
   End
   Begin VB.Frame fraTables 
      Height          =   2535
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   3915
      Begin VB.ListBox lstTables 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1980
         ItemData        =   "frmConvert.frx":014E
         Left            =   120
         List            =   "frmConvert.frx":0150
         TabIndex        =   8
         Top             =   420
         Width           =   3675
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Table(s) Available"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   180
         Width           =   1545
      End
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   11715
      Begin VB.CommandButton cmdSourceBrowse 
         Caption         =   "..."
         Height          =   360
         Left            =   11040
         TabIndex        =   5
         Top             =   240
         Width           =   435
      End
      Begin VB.TextBox txtSourceDatabaseFile 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   7140
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   240
         Width           =   3795
      End
      Begin VB.ComboBox cmbSourceDatabase 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2100
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   240
         Width           =   2115
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Source Database File"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   4740
         TabIndex        =   3
         Top             =   300
         Width           =   2280
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Source Database"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   180
         TabIndex        =   1
         Top             =   300
         Width           =   1830
      End
   End
End
Attribute VB_Name = "frmConvert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit     ''no voodoo variables

Dim dp As DatabaseProperties
Dim rs As ADODB.Recordset
Dim flds As ADODB.Field

Dim li As ListItem
Dim lsi As ListSubItem

Dim strFields() As String

Dim obj As Object

Private Sub cmbDestDatabase_Click()

''Select dest database type
With cmbDestDatabase

    ''IMPORTANT : use item data property
    dp.DestDatabaseType = .ItemData(.ListIndex)
    
    ''if it is file based database let the user browse
    If .ItemData(.ListIndex) < 100 Then
        ''file based databases
        cmdDestBrowse.Enabled = True
    Else
        ''oracle, sqlserver
        cmdDestBrowse.Enabled = False
    End If
    
End With

End Sub

Private Sub cmbSourceDatabase_Click()

''enable/disable cmdbrowse button
With cmbSourceDatabase
    ''assign source datbase type
    
    ''IMPORTANT
    '''''''''''''''''''''''''
    ''Use itemdata property''
    '''''''''''''''''''''''''
    dp.SourceDatabaseType = .ItemData(.ListIndex)
    
    ''if it is file based database let the user browse
    If .ItemData(.ListIndex) < 100 Then
        ''file based databases
        cmdSourceBrowse.Enabled = True
    Else
        ''oracle, sqlserver
        cmdSourceBrowse.Enabled = False
    End If
End With

End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdCreate_Click()

Dim i As Integer
Dim bResult As Boolean  ''process is success or not

''insufficient input details don't try to proceed
If _
    dp.SourceDatabaseFile = "" Or _
    dp.SourceDatabaseFilePath = "" Or _
    dp.DestDatabaseFile = "" Then
            MsgBox "Please verify source/destination database files", vbInformation, "Error processing"
        GoTo PROC_EXIT
End If

''destination file already exists
If Dir(dp.DestDatabaseFilePath & dp.DestDatabaseFile) <> "" Then
    MsgBox "File already exists, Please give a different name", vbInformation, "DitETS"
    GoTo PROC_EXIT
End If

''no fields selected to create a destination database
If lstSelectedFields.ListCount = 0 Then
    MsgBox "Please select fileds", vbInformation, "Invalid selection"
    GoTo PROC_EXIT
End If
    
''Prepare the selected fields from the list box
ReDim strFields(0 To lstSelectedFields.ListCount - 1) As String

For i = 0 To lstSelectedFields.ListCount - 1
    strFields(i) = lstSelectedFields.List(i)
Next
    
''call the convert method
bResult = obj.clsSuper_Convert(strFields, _
                            , CInt(dp.SourceDatabaseType), _
                            dp.SourceDatabaseFile, _
                            dp.SourceDatabaseFilePath, _
                            dp.SourceDatabaseTableName, _
                            , , , , CInt(dp.DestDatabaseType), _
                            dp.DestDatabaseFile, _
                            dp.DestDatabaseFilePath)

If bResult Then
    ''process sucess
    MsgBox "Data converted successfully", vbInformation, "DitETL"
Else
    ''process failed
    MsgBox "Unable to convert the data", vbCritical, "DitETL"
End If

''prepare a new job
dp.DestDatabaseFile = ""
dp.DestDatabaseFilePath = ""
dp.DestDatabaseName = ""
dp.DestDatabasePassword = ""
dp.DestDatabaseServerName = ""
dp.DestDatabaseServerName = ""
dp.DestDatabaseUserName = ""

txtDestDatabaseFile.Text = ""

PROC_EXIT:
    Exit Sub

End Sub

Private Sub cmdDestBrowse_Click()

On Error GoTo ehCmdDestBrowse

With cd
    .DialogTitle = "Select the destination database file"
    ''on cancel don't process anymore
    .CancelError = True
    .FilterIndex = 1
End With

Select Case dp.DestDatabaseType

    Case Is = 51 'Text
            cd.Filter = "Text files (*.txt;*.csv;*.asc)|*.txt;*.csv;*.asc"
        
    Case Is = 1 'Access
            cd.Filter = "Access Database (*.mdb)|*.mdb"
        
    Case Is = 11, 12, 13  'dBASE III, dBASE IV, dBASE 5.0
            cd.Filter = "dBASE files (*.dbf)|*.dbf"
        
    Case Is = 21, 22, 23, 24, 25    'Excel 3,4,5,97,8
            cd.Filter = "Excel files (*.xls)|*.xls"
    
    Case Is = 31    'HTML
            cd.Filter = "HTML files (*.htm;*html)|*.htm;*html"
        
    Case Is = 41, 42, 43    'Paradox 3,4,5
            cd.Filter = "Paradox files (*.db)|*.db"
    
End Select

cd.ShowSave
With dp
    .DestDatabaseFile = GetFileName(cd.FileName)
    .DestDatabaseFilePath = GetFilePath(cd.FileName, True)
    txtDestDatabaseFile.Text = cd.FileName
End With

PROC_EXIT:
    Exit Sub
    
ehCmdDestBrowse:
    Resume Next
End Sub

Private Sub cmdOrgOrder_Click()
''add fields in the org order

Dim i As Integer
lstSelectedFields.Clear

With lvFields
For i = 1 To .ListItems.Count
    If .ListItems(i).Checked = True Then
        lstSelectedFields.AddItem .ListItems(i).Text
    End If
Next
End With

End Sub

Private Sub cmdSelectAll_Click()

Dim i As Integer

lstSelectedFields.Clear

With lvFields

For i = 1 To .ListItems.Count
    .ListItems(i).Checked = False
    .ListItems(i).Checked = True
    lstSelectedFields.AddItem .ListItems(i).Text
Next
.Refresh
End With

End Sub

Private Sub cmdSelectNone_Click()
lstSelectedFields.Clear
ClearFieldSelection
End Sub

Private Sub cmdSortFieldsAlpha_Click()
'''
'Dim i As Integer
'Dim j As Integer
'Dim strTemp As String
'
'''sort here ( a simple sort will do)
'With lstSelectedFields
'    For i = 0 To .ListCount - 1
'        For j = 0 To .ListCount - 1
'            If StrComp(.List(i), .List(j)) < 0 Then
'                strTemp = .List(j)
'                .List(j) = .List(i)
'                .List(i) = strTemp
'            End If
'        Next
'    Next
'End With
SortFields True

End Sub

Private Sub cmdSortFieldsRevAlpha_Click()
SortFields False
End Sub

Private Sub cmdSourceBrowse_Click()

''identify the source database type
''this will help to apply filter in common dialog

On Error GoTo ehCmdSourceBrowse

Dim i As Integer
Dim strTables() As String

With cd
    .DialogTitle = "Select the source database file"
    ''on cancel don't process anymore
    .CancelError = True
    .FilterIndex = 1
    ''hide read only check box
    .Flags = cdlOFNHideReadOnly
End With

Select Case dp.SourceDatabaseType

''=====================================================================
    Case Is = 51 ''Text File
        ''''Steps
        ''''=====
        ''''1. Show File Open Dialog
        ''''2. Set SourceDatabse = text
        With dp
            cd.Filter = "Text files (*.txt;*.csv;*.asc)|*.txt;*.csv;*.asc"
            cd.ShowOpen
            
            .SourceDatabaseFile = GetFileName(cd.FileName)
            .SourceDatabaseFilePath = GetFilePath(cd.FileName, True)
            txtSourceDatabaseFile.Text = cd.FileName
        End With
                  
        ''clear the fields list box
        lstSelectedFields.Clear
        
        ''clear the list view (fields details)
        For i = 1 To lvFields.ListItems.Count
            lvFields.ListItems.Remove 1
        Next

        ''clear the tables list box
        lstTables.Clear

        If Not (obj Is Nothing) Then Set obj = Nothing
        
        Set obj = New clsText

        ''receive the fileds detail
        If rs.State <> 0 Then rs.Close
        
        Set rs = obj.clsSuper_GetFields(, _
                        CInt(dp.SourceDatabaseType), _
                        dp.SourceDatabaseFile, dp.SourceDatabaseFilePath)
        
        If ErrorCode <> 0 Then GoTo PROC_EXIT

        ''display number of records found in the database
        lblTotal.Caption = obj.clsSuper_GetTotalRecords
        
        ''fill list-view with fields name,type,size
        LoadFields
        
        ''close the recordset
        rs.Close
               
''=====================================================================
    Case Is = 1 ''Access Database
        
        With dp
            cd.Filter = "Access Database (*.mdb)|*.mdb"
            cd.ShowOpen
            .SourceDatabaseFile = GetFileName(cd.FileName)
            .SourceDatabaseFilePath = GetFilePath(cd.FileName, True)
            txtSourceDatabaseFile.Text = cd.FileName
        End With
                   
        ''clear the fields list box
        lstSelectedFields.Clear
        
        ''clear the list view (fields details)
        For i = 1 To lvFields.ListItems.Count
            lvFields.ListItems.Remove 1
        Next

        ''clear the tables list box
        lstTables.Clear
        
        If Not (obj Is Nothing) Then Set obj = Nothing
        
        Set obj = New clsAccess
        
        strTables() = obj.clsSuper_GetTables( _
                        , CInt(dp.SourceDatabaseType), _
                        dp.SourceDatabaseFile, _
                        dp.SourceDatabaseFilePath)
                                
        If ErrorCode <> 0 Then GoTo PROC_EXIT
                                
        ''load the table names in lstTables
        For i = 1 To UBound(strTables())
            lstTables.AddItem strTables(i)
        Next
''=====================================================================

    Case Is = 11, 12, 13, 41, 42, 43  ''dBASE File and other formats
        ''''Steps
        ''''=====
        ''''1. Show File Open Dialog
        ''''2. Set SourceDatabse = text
        With dp
            cd.Filter = "dBASE files (*.dbf)|*.dbf|Paradox files (*.db)|*.db"
            cd.ShowOpen
            
            .SourceDatabaseFile = GetFileName(cd.FileName)
            .SourceDatabaseFilePath = GetFilePath(cd.FileName, True)
            txtSourceDatabaseFile.Text = cd.FileName
        End With
                  
        ''clear the fields list box
        lstSelectedFields.Clear
        
        ''clear the list view (fields details)
        For i = 1 To lvFields.ListItems.Count
            lvFields.ListItems.Remove 1
        Next

        ''clear the tables list box
        lstTables.Clear

        If Not (obj Is Nothing) Then Set obj = Nothing
        
        Set obj = New clsdBASE

        strTables() = obj.clsSuper_GetTables( _
                        , CInt(dp.SourceDatabaseType), _
                        dp.SourceDatabaseFile, _
                        dp.SourceDatabaseFilePath)
                        
        If ErrorCode <> 0 Then GoTo PROC_EXIT
                        
        ''load the table names in lstTables
        For i = 1 To UBound(strTables())
            lstTables.AddItem strTables(i)
        Next
        
        lstTables.Refresh
''=====================================================================
Case Is = 21, 22, 23, 24, 25 ''Excel Database
        
        With dp
            cd.Filter = "Excel Database (*.xls)|*.xls"
            cd.ShowOpen
            .SourceDatabaseFile = GetFileName(cd.FileName)
            .SourceDatabaseFilePath = GetFilePath(cd.FileName, True)
            txtSourceDatabaseFile.Text = cd.FileName
        End With
                   
        ''clear the fields list box
        lstSelectedFields.Clear
        
        ''clear the list view (fields details)
        For i = 1 To lvFields.ListItems.Count
            lvFields.ListItems.Remove 1
        Next

        ''clear the tables list box
        lstTables.Clear
        
        If Not (obj Is Nothing) Then Set obj = Nothing
        
        Set obj = New clsExcel
        
        strTables() = obj.clsSuper_GetTables( _
                        , CInt(dp.SourceDatabaseType), _
                        dp.SourceDatabaseFile, _
                        dp.SourceDatabaseFilePath)
                        
        If ErrorCode <> 0 Then GoTo PROC_EXIT
                        
        ''load the table names in lstTables
        For i = 1 To UBound(strTables())
            If strTables(i) <> "" Then lstTables.AddItem strTables(i)
        Next
''=====================================================================
    Case Is = 31 ''Html File
       ''
''=====================================================================

End Select

PROC_EXIT:

    If ErrorCode <> 0 Then
        MsgBox ErrorCode
        Set obj = Nothing
    End If

    Exit Sub

ehCmdSourceBrowse:

    If Err.Number = 32755 Then Exit Sub
    Resume Next
End Sub

Private Sub cmdFieldsDelete_Click()
''remove selected item
Dim i As Integer

With lstSelectedFields
    If .ListCount > 0 Then
        If .ListIndex >= 0 Then
            ''uncheck the deleted item in the fields list view
            For i = 1 To lvFields.ListItems.Count
                If lvFields.ListItems(i).Text = .List(.ListIndex) Then
                    lvFields.ListItems(i).Checked = False
                End If
            Next
            ''remove the item from the fields list box
            lstSelectedFields.RemoveItem lstSelectedFields.ListIndex
        End If
    End If
End With
End Sub

Private Sub Form_Load()

Me.Caption = Me.Caption & " v" & App.Major & "." & App.Minor

With cmbSourceDatabase
    .Clear
    
    ''file based databases
    
    .AddItem "Access"
    .ItemData(.NewIndex) = 1
    
    .AddItem "dBASE"
    .ItemData(.NewIndex) = 11
    
    .AddItem "Excel"
    .ItemData(.NewIndex) = 21
    
    .AddItem "HTML"
    .ItemData(.NewIndex) = 31
    
    .AddItem "Paradox"
    .ItemData(.NewIndex) = 41
    
    .AddItem "Text"
    .ItemData(.NewIndex) = 51
    
    '''''''''''''''''''''''''''''
    ''non-file based databases
    '''''''''''''''''''''''''''''
    
    .AddItem "ORACLE"
    .ItemData(.NewIndex) = 101
    
    .AddItem "SQLServer"
    .ItemData(.NewIndex) = 111
    
    .ListIndex = 0
End With

With cmbDestDatabase
    .Clear
    
    .AddItem "Access 2000"
    .ItemData(.NewIndex) = 1
    
    .AddItem "dBASE III"
    .ItemData(.NewIndex) = 11
    
    .AddItem "dBASE IV"
    .ItemData(.NewIndex) = 12
    
    .AddItem "dBASE 5.0"
    .ItemData(.NewIndex) = 13
    
    .AddItem "Excel 3.0"
    .ItemData(.NewIndex) = 21
    
    .AddItem "Excel 4.0"
    .ItemData(.NewIndex) = 22
    
    .AddItem "Excel 5.0"
    .ItemData(.NewIndex) = 23
    
'    .AddItem "Excel 97"
'    .ItemData(.NewIndex) = 24
    
    .AddItem "Excel 8.0"
    .ItemData(.NewIndex) = 25
    
    .AddItem "HTML"
    .ItemData(.NewIndex) = 31
    
    .AddItem "Paradox 3.x"
    .ItemData(.NewIndex) = 41
    
    .AddItem "Paradox 4.x"
    .ItemData(.NewIndex) = 42
    
    .AddItem "Paradox 5.x"
    .ItemData(.NewIndex) = 43
    
    .AddItem "Text"
    .ItemData(.NewIndex) = 51
    
    ''non-file based database
    
    .AddItem "ORACLE"
    .ItemData(.NewIndex) = 101
    
    .AddItem "SQLServer"
    .ItemData(.NewIndex) = 111
    
    .ListIndex = 0
End With

Set rs = New ADODB.Recordset

End Sub

Public Sub LoadFields()

Dim i As Integer

With lvFields
    If .ColumnHeaders.Count = 0 Then
        .ColumnHeaders.Add , , "Field Name", 3000
        .ColumnHeaders.Add , , "Field Type", 2000
        .ColumnHeaders.Add , , "Field Size", 1900
    End If


    i = 1
    For Each flds In rs.Fields
        Set li = .ListItems.Add(i, , flds.Name)
        Set lsi = .ListItems(i).ListSubItems.Add(1, , FindFieldType(flds.Type))
        ''empty table(s) will give trouble here so check the
        ''record count
        If rs.RecordCount > 0 Then
            Set lsi = .ListItems(i).ListSubItems.Add(2, , flds.ActualSize)
        End If
        i = i + 1
    Next
    
End With

End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

If Not (obj Is Nothing) Then Set obj = Nothing

If Not (rs Is Nothing) Then
    If rs.State <> 0 Then rs.Close
    Set rs = Nothing
End If

Set frmConvert = Nothing

End Sub

Private Sub lstTables_DblClick()

''if no item selected exit sub
With lstTables
    If .ListCount < 0 And .ListIndex < 0 Then Exit Sub
    
    ''Table name shouldn't contain spaces
    If InStr(1, .List(.ListIndex), " ") Then
        MsgBox "Unable to open this table", vbInformation, "DitETL"
        Exit Sub
    End If
    
    ''get the table name
    dp.SourceDatabaseTableName = .List(.ListIndex)
End With

''clear list view
ClearFields

''clear lstSelectedfields
lstSelectedFields.Clear

If rs.State = 1 Then rs.Close

Set rs = obj.clsSuper_GetFields(, CInt(dp.SourceDatabaseType), _
                                dp.SourceDatabaseFile, _
                                dp.SourceDatabaseFilePath, _
                                dp.SourceDatabaseTableName)
If ErrorCode <> 0 Then
    MsgBox ErrorCode
    Exit Sub
End If

''display number of records found in the database
lblTotal.Caption = obj.clsSuper_GetTotalRecords

''fill list-view with fields name,type,size
LoadFields
lvFields.Refresh
''close the recordset
rs.Close

End Sub

Private Sub lvFields_ItemCheck(ByVal Item As MSComctlLib.ListItem)

Dim i As Integer

With lstSelectedFields
    For i = 0 To .ListCount
        If .List(i) = Item Then
            If Not Item.Checked Then
                .RemoveItem i
                Exit Sub
            Else
                Exit Sub
            End If
        End If
    Next
    .AddItem Item
    .ItemData(.ListCount - 1) = Item.Index - 1
End With

End Sub

Private Sub udFieldsUpDown_DownClick()
Dim fi As String
With lstSelectedFields
    If .ListCount > 0 Then
        If .ListIndex < .ListCount - 1 And .ListIndex >= 0 Then
            fi = .List(.ListIndex)
            .List(.ListIndex) = .List(.ListIndex + 1)
            .List(.ListIndex + 1) = fi
            .ListIndex = .ListIndex + 1
        End If
    End If
End With
End Sub

Private Sub udFieldsUpDown_UpClick()
Dim fi As String
With lstSelectedFields
    If .ListCount > 0 Then
        If .ListIndex > 0 Then
            fi = .List(.ListIndex)
            .List(.ListIndex) = .List(.ListIndex - 1)
            .List(.ListIndex - 1) = fi
            .ListIndex = .ListIndex - 1
        End If
    End If
End With
End Sub

Public Sub ClearFieldSelection()
''clear list view fields selection
Dim i As Integer
With lvFields
    For i = 1 To .ListItems.Count
        .ListItems(i).Checked = False
    Next
End With
End Sub

Public Sub SortFields(bOrder As Boolean)
''this sub will sort the selected fields
''either in alphabet / reverse alphabet
''order
''bOrder is used to identify the sorting order i.e alphabet/reverse order

Dim i As Integer
Dim j As Integer
Dim strTemp As String

With lstSelectedFields
    If bOrder Then
        ''alphabet order
        For i = 0 To .ListCount - 1
            For j = 0 To .ListCount - 1
                If StrComp(.List(i), .List(j)) < 0 Then
                    strTemp = .List(j)
                    .List(j) = .List(i)
                    .List(i) = strTemp
                End If
            Next
        Next
        
    Else
        ''reverse alphabet order
        For i = 0 To .ListCount - 1
            For j = 0 To .ListCount - 1
                If StrComp(.List(i), .List(j)) > 0 Then
                    strTemp = .List(j)
                    .List(j) = .List(i)
                    .List(i) = strTemp
                End If
            Next
        Next
        
    End If
End With

End Sub

Public Sub ClearFields()
''clear list view fields selection
Dim i As Integer
With lvFields
    For i = 1 To .ListItems.Count
        .ListItems.Remove 1
    Next
End With
End Sub
