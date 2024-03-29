VERSION 5.00
Begin VB.Form FMain 
   Caption         =   "PropKeyHReader"
   ClientHeight    =   8655
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   14535
   Icon            =   "FMain.frx":0000
   LinkTopic       =   "FMain"
   ScaleHeight     =   8655
   ScaleWidth      =   14535
   StartUpPosition =   3  'Windows-Standard
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7935
      Left            =   3600
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   0
      Top             =   0
      Width           =   9975
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7710
      ItemData        =   "FMain.frx":1782
      Left            =   0
      List            =   "FMain.frx":1784
      TabIndex        =   1
      Top             =   0
      Width           =   3495
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileopen 
         Caption         =   "Open..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "Save As..."
      End
      Begin VB.Menu mnuFileSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "&Exit"
         Shortcut        =   ^E
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "Edit"
      Begin VB.Menu mnuEditStatAllDifDatatypes 
         Caption         =   "Stats all diff. datatypes"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuEditCopyAll 
         Caption         =   "Copy All for Excel-Import"
         Shortcut        =   ^I
      End
      Begin VB.Menu mnuEditCreateVBEnum 
         Caption         =   "Create VB-Enum"
      End
      Begin VB.Menu mnuEditCreateSubInit 
         Caption         =   "Create Sub Init()"
         Begin VB.Menu mnuEditCreateSubInit1 
            Caption         =   "Part 1"
         End
         Begin VB.Menu mnuEditCreateSubInit2 
            Caption         =   "Part 2"
         End
         Begin VB.Menu mnuEditCreateSubInit3 
            Caption         =   "Part 3"
         End
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "?"
      Begin VB.Menu mnuHelpInfo 
         Caption         =   "Info"
      End
   End
End
Attribute VB_Name = "FMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'For this project you need the following files not contained in this package:
'Modules:
' * MErr:    Err_CorrectErrorHandling\Modules\MErr.bas
' * MPtr:    Ptr_Pointers\Modules\MPtr.bas
' * MShell:  IO_PathFileName\Modules\MShell.bas
' * MString: Sys_Strings\Modules\MString.bas
'Classes:
' * List:           List_GenericNLinq\Classes\List.cls
' * OpenFileDialog: Win_Dialogs\Classes\OpenFileDialog.cls
' * PathFileName:   IO_PathFileName\Classes\PathFileName.cls
' * SaveFileDialog: Win_Dialogs\Classes\SaveFileDialog.cls
' * StringBuilder:  Sys_StringBuilder\Classes\StringBuilder.cls

Private m_FirstActivate As Boolean

Private Sub Form_Load()
    m_FirstActivate = True
    Me.Caption = Me.Caption & " v" & App.Major & "." & App.Minor & "." & App.Revision
End Sub
'
Private Sub Form_Activate()
    If m_FirstActivate Then
        If MApp.AutomaticOpenNReadPropKeyHFile(True) Then
            UpdateView
        End If
    End If
    m_FirstActivate = False
End Sub

Public Sub UpdateCaption()
    Me.Caption = "PropKeyHReader v" & App.Major & "." & App.Minor & "." & App.Revision & " - " & MApp.DocumentTsv.FileName.Value
End Sub

Public Sub UpdateView()
    UpdateCaption
    Text1.Text = vbNullChar
    List1.Clear
    Dim i As Long, le As List
    Dim PLists As List: Set PLists = MApp.PropertyLists
    For i = 0 To PLists.Count - 1
        Set le = PLists.Item(i)
        List1.AddItem le.Name & "(" & le.Count & ")"
    Next
    If List1.ListCount > 0 Then List1.ListIndex = 0
End Sub

Private Sub List1_Click()
    Dim i As Long: i = List1.ListIndex
    If i < 0 Or MApp.PropertyLists.Count <= i Then Exit Sub
    Dim le As List: Set le = MApp.PropertyLists.Item(i)
    Text1.Text = MApp.PropertyList_ToStr(le)
End Sub

Private Sub Form_Resize()
    Dim L As Single: L = 0
    Dim T As Single: T = List1.Top
    Dim W As Single: W = List1.Width
    Dim H As Single: H = Me.ScaleHeight - T
    If W > 0 And H > 0 Then List1.Move L, T, W, H
    L = W
    W = Me.ScaleWidth - L
    If W > 0 And H > 0 Then Text1.Move L, T, W, H
End Sub

' ############################## ' menu File  ' ############################## '
Private Sub mnuFileNew_Click()
    MApp.DataClear
    UpdateView
End Sub

Private Sub mnuFileOpen_Click()
    Dim OFD As OpenFileDialog: Set OFD = New OpenFileDialog
    With OFD
        .InitialDirectory = App.Path
        .Filter = "Header-Files (*.h)|*.h|Tab sepated values (*.tsvdb)|*.tsvdb"
        If .ShowDialog() = vbCancel Then Exit Sub
        Dim aFile As PathFileName: Set aFile = MNew.PathFileName(.FileName)
    End With
    Dim bOK As Boolean: bOK = MApp.ReadFile(aFile)
    If bOK Then Me.UpdateView: Exit Sub
    MsgBox "Could not read the file: " & vbCrLf & aFile.Value
End Sub

Private Sub mnuFileSave_Click()
    If LCase(MApp.DocumentTsv.FileName.Extension) = ".h" Then
        mnuFileSaveAs_Click
    Else
        MApp.WriteFile MApp.DocumentTsv.FileName
        UpdateView
    End If
End Sub

Private Sub mnuFileSaveAs_Click()
    Dim SFD As SaveFileDialog: Set SFD = New SaveFileDialog
    With SFD
        .AddExtension = True
        .InitialDirectory = App.Path
        .Filter = "Tab separated values (*.tsvdb)|*.tsvdb"
        If .ShowDialog = vbCancel Then Exit Sub
        Dim tsvdb As PathFileName: Set tsvdb = MNew.PathFileName(SFD.FileName)
        If Not MApp.WriteFile(tsvdb) Then
            MsgBox "Error writing file: " & vbCrLf & tsvdb.Value
        End If
        UpdateView
    End With
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

' ############################## ' menu Edit  ' ############################## '
Private Sub mnuEditStatAllDifDatatypes_Click()
    Dim datatypes As List: Set datatypes = MApp.StatsAllDifDatatypes
    If datatypes Is Nothing Then
        MsgBox "Error: failed to create statistics failed"
        Exit Sub
    End If
    Dim w1 As Long, w2 As Long, w3 As Long
    Dim s0 As String, s As String
    Dim i As Long, pkhe As PropKeyHEntry
    For i = 0 To datatypes.Count - 1
        Set pkhe = datatypes.Item(i)
        w1 = Max(w1, Len(pkhe.Name))
        w2 = Max(w2, Len(pkhe.DataType))
    Next
    For i = 0 To datatypes.Count - 1
        Set pkhe = datatypes.Item(i)
        s0 = PadRight(pkhe.Name, w1) & " || " & PadRight(pkhe.DataType, w2) & " :: " & pkhe.PKVarTyp
        s = s & s0 & vbCrLf
    Next
    Text1.Text = s
        
End Sub

Private Sub mnuEditCopyAll_Click()
    MApp.ClipboardCopyAll
End Sub

Private Sub mnuEditCreateVBEnum_Click()
    Text1.Text = MApp.DocumentTsv.ToModuleEnum
End Sub

Private Sub mnuEditCreateSubInit1_Click()
    'just because a VB.TextBox is not capable to hold so much text
    Text1.Text = MApp.DocumentTsv.ToModuleInit1
End Sub

Private Sub mnuEditCreateSubInit2_Click()
    Text1.Text = MApp.DocumentTsv.ToModuleInit2
End Sub

Private Sub mnuEditCreateSubInit3_Click()
    Text1.Text = MApp.DocumentTsv.ToModuleInit3
End Sub

' ############################## ' menu Help  ' ############################## '
Private Sub mnuHelpInfo_Click()
    MsgBox App.CompanyName & " " & Me.Caption & vbCrLf & App.FileDescription '& vbCrLf & " v" & App.Major & "." & App.Minor & "." & App.Revision
End Sub
