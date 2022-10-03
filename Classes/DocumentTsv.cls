VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DocumentTsv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_PFN           As PathFileName
Private m_PropertyLists As List '(Of List(Of PropKeyHEntry))

Private m_Count         As Long

Friend Sub New_(aPFN As PathFileName, Optional ByVal aPropertyLists As List = Nothing)
    Set m_PFN = aPFN: Set m_PropertyLists = aPropertyLists
End Sub
Private Function GetCount() As Long
    Dim i As Long, c As Long, le As List
    For i = 0 To m_PropertyLists.Count - 1
        Set le = m_PropertyLists.Item(i)
        c = c + le.Count
    Next
    GetCount = c
End Function

Public Property Get FileName() As PathFileName
    Set FileName = m_PFN
End Property

Public Property Get PropertyLists() As List
    Set PropertyLists = m_PropertyLists
End Property

Public Property Get ItemPropKeyHEntry(ByVal Index As Long) As PropKeyHEntry
    If (0 <= Index) And (Index < m_Count) Then
        Dim le As List
        Dim i As Long, ci As Long: ci = Index
        For i = 0 To m_PropertyLists.Count - 1
            Set le = m_PropertyLists.Item(i)
            ci = ci - le.Count
            If ci < 0 Then
                Index = le.Count + ci
                Set ItemPropKeyHEntry = le.Item(Index)
                Exit Function
            End If
        Next
    End If
End Property

Public Function Read() As Boolean
Try: On Error GoTo Catch
    Dim lines() As String
    If Not m_PFN.TryReadAllLines(lines) Then
        MsgBox "Could not read the file: " & vbCrLf & m_PFN.Value
        Exit Function
    End If
    Set m_PropertyLists = MNew.List(EDataType.vbObject)
    Dim i As Long, u As Long: u = UBound(lines)
    Dim line As String
    Dim pkl As List
    Dim pke As PropKeyHEntry
    For i = 0 To u
        line = lines(i)
        Dim data() As String: data = Split(line, vbTab)
        Dim s As String: s = Left(line, 1)
        If s = "#" Then
            'just a comment go on
        ElseIf s = vbTab Then
            Set pke = PropKeyHEntry_ParseFromTsvDB(data)
            If Not pkl Is Nothing Then
                pkl.Add pke
            End If
        Else
            Set pkl = PropertyLists.Add(MNew.List(EDataType.vbObject))
            pkl.Name = Trim(line)
        End If
    Next
    If Not m_PropertyLists Is Nothing Then m_Count = GetCount
    Read = True
    GoTo Finally
Catch:
    MsgBox "Could not read from file: " & vbCrLf & m_PFN.Value
Finally:
    m_PFN.CloseFile
End Function

Private Function PropKeyHEntry_ParseFromTsvDB(data() As String) As PropKeyHEntry
    Dim i As Long, u As Long: u = UBound(data)
    Dim pke As PropKeyHEntry: Set pke = New PropKeyHEntry
    With pke
        i = i + 1 'the first column is empty, because it contains the name of the group of entries
        If i <= u Then .Name = data(i):     i = i + 1  ' System.Audio.ChannelCount
        If i <= u Then .PKEYName = data(i): i = i + 1  ' PKEY_Audio_ChannelCount
        If i <= u Then .DataType = data(i): i = i + 1  ' UInt32
        If i <= u Then .PKVarTyp = data(i): i = i + 1  ' VT_UI4
        If i <= u Then .FormatID = data(i): i = i + 1  ' FMTID_AudioSummaryInformation
        If i <= u Then .FmtGuid = data(i):  i = i + 1  ' 64440490-4C8B-11D1-8B70-080036B11A03
        If i <= u Then .PIDName = data(i):  i = i + 1  ' PIDASI_CHANNEL_COUNT
        If i <= u Then .PIDValue = data(i): i = i + 1  ' 7
        If i <= u Then .Descript = data(i): i = i + 1  ' Indicates the channel count for the audio file. Values: 1 (mono), 2 (stereo).
    End With
    Set PropKeyHEntry_ParseFromTsvDB = pke
End Function


Public Function WriteTsvDB() As Boolean
Try: On Error GoTo Catch
    m_PFN.WriteStr PropertyListsToStr
    WriteTsvDB = True
    GoTo Finally
Catch:
    MErr.MessError TypeName(Me), "WriteTSVDB", "Could not write to file: " & vbCrLf & m_PFN.Value
Finally:
    m_PFN.CloseFile
End Function

Public Function PropertyListsToStr() As String
'    Dim sb As New StringBuilder
'    Dim i As Long, pkl As List
'    For i = 0 To m_PropertyLists.Count - 1
'        Set pkl = m_PropertyLists.Item(i)
'        If pkl Is Nothing Then Exit For
'        sb.AppendLine pkl.Name
'        Dim j As Long, pkhe As PropKeyHEntry
'        For j = 0 To pkl.Count - 1
'            Set pkhe = pkl.Item(j)
'            If Not pkhe Is Nothing Then
'                sb.AppendLine vbTab & pkhe.ToStr
'            End If
'        Next
'    Next
'    PropertyListsToStr = sb.ToStr
End Function

Public Sub ToListbox(aCmbList)
    'aCmbList = ComboBox or ListBox
    aCmbList.Clear
    Dim i As Long, le As List
    For i = 0 To m_PropertyLists.Count - 1
        Set le = m_PropertyLists.Item(i)
        aCmbList.AddItem le.Name & "(" & le.Count & ")"
    Next
    If aCmbList.ListCount > 0 Then aCmbList.ListIndex = 0
End Sub

Public Sub ListOfPKHEntryToListBox(aCmbList)
    aCmbList.Clear
    Dim i As Long, j As Long, le As List, pk As PropKeyHEntry
    For i = 0 To m_PropertyLists.Count - 1
        Set le = m_PropertyLists.Item(i)
        For j = 0 To le.Count - 1
            Set pk = le.Item(j)
            aCmbList.AddItem pk.Name
        Next
    Next
    'If aCmbList.ListCount > 0 Then aCmbList.ListIndex = 0
End Sub

