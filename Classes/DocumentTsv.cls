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

Friend Sub New_(aPFN As PathFileName, Optional ByVal aPropertyLists As List = Nothing)
    Set m_PFN = aPFN: Set m_PropertyLists = aPropertyLists
End Sub

Public Property Get FileName() As PathFileName
    Set FileName = m_PFN
End Property

Public Function ToTsvDB() As DocumentTsv
    Set ToTsvDB = MNew.DocumentTsv(m_PFN, m_PropertyLists)
End Function

Public Property Get PropertyLists() As List
    Set PropertyLists = m_PropertyLists
End Property

Public Function Read() As Boolean
Try: On Error GoTo Catch
    Dim lines() As String
    If Not m_PFN.TryReadAllLines(lines) Then
        MsgBox "Could not read the file: " & vbCrLf & m_PFN.Value
        Exit Function
    End If
    Set m_PropertyLists = MNew.List(vbObject)
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
            Set pke = MParserPkh.PropKeyHEntry_ParseFromTsvDB(data)
            If Not pkl Is Nothing Then
                pkl.Add pke
            End If
        Else
            Set pkl = PropertyLists.Add(MNew.List(vbObject))
            pkl.Name = Trim(line)
        End If
    Next
    Read = True
Catch:
Finally:
    m_PFN.CloseFile
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
    Dim sb As New StringBuilder
    Dim i As Long, pkl As List
    For i = 0 To m_PropertyLists.Count - 1
        Set pkl = m_PropertyLists.Item(i)
        If pkl Is Nothing Then Exit For
        sb.AppendLine pkl.Name
        Dim j As Long, pkhe As PropKeyHEntry
        For j = 0 To pkl.Count - 1
            Set pkhe = pkl.Item(j)
            If Not pkhe Is Nothing Then
                sb.AppendLine vbTab & pkhe.ToStr
            End If
        Next
    Next
    PropertyListsToStr = sb.ToStr
End Function


