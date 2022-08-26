VERSION 5.00
Begin VB.Form FMain 
   Caption         =   "PropKeyHReader"
   ClientHeight    =   8655
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   14535
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
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
      Left            =   3120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   0
      Top             =   0
      Width           =   11295
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
      ItemData        =   "Form1.frx":1782
      Left            =   0
      List            =   "Form1.frx":1784
      TabIndex        =   1
      Top             =   0
      Width           =   3135
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "Open propkey.h"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "Save As..."
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "Edit"
      Begin VB.Menu mnuEditCopyAll 
         Caption         =   "Copy All for Excel-Import"
         Shortcut        =   ^C
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
        If AutomaticOpenNReadPropKeyHFile(True) Then
            UpdateView
        End If
    End If
    m_FirstActivate = False
End Sub

Public Sub UpdateView() 'pkl As List)
    List1.Clear
    Text1.Text = vbNullChar
    Dim i As Long, le As List
    'For Each le In MApp.PropertyLists.GetEnumerator
    For i = 0 To MApp.PropertyLists.Count - 1
        Set le = MApp.PropertyLists.Item(i)
        List1.AddItem le.Name
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
    Dim t As Single: t = List1.Top
    Dim W As Single: W = List1.Width
    Dim H As Single: H = Me.ScaleHeight - t
    If W > 0 And H > 0 Then List1.Move L, t, W, H
    L = W
    W = Me.ScaleWidth - L
    If W > 0 And H > 0 Then Text1.Move L, t, W, H
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
        Dim PropKeyH As PathFileName: Set PropKeyH = MNew.PathFileName(.FileName)
    End With
    Dim bOK As Boolean
    If PropKeyH.Extension = ".h" Then
        bOK = MApp.ReadFileH(PropKeyH)
    ElseIf PropKeyH.Extension = ".tsvdb" Then
        bOK = MApp.ReadFileTsvdb(PropKeyH)
    End If
    If bOK Then Me.UpdateView: Exit Sub
    MsgBox "Could not read the file: " & vbCrLf & PropKeyH.value
End Sub

Private Sub mnuFileSaveAs_Click()
    Dim SFD As SaveFileDialog: Set SFD = New SaveFileDialog
    With SFD
        .AddExtension = True
        .InitialDirectory = App.Path
        .Filter = "Tab separated values (*.tsvdb)|*.tsvdb"
        If .ShowDialog = vbCancel Then Exit Sub
        Dim tsvdb As PathFileName: Set tsvdb = MNew.PathFileName(SFD.FileName)
        If LCase(tsvdb.Extension) <> ".tsvdb" Then tsvdb.Extension = "tsvdb"
        MApp.WriteTSVDB tsvdb
    End With
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

' ############################## ' menu Edit  ' ############################## '
Private Sub mnuEditCopyAll_Click()
    Dim s As String: s = MApp.PropertyLists_ToStr
    Clipboard.SetText s
End Sub

' ############################## ' menu Help  ' ############################## '
Private Sub mnuHelpInfo_Click()
    MsgBox App.CompanyName & " " & Me.Caption & vbCrLf & App.FileDescription
End Sub
