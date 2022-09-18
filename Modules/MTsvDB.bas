Attribute VB_Name = "MTsvDB"
Option Explicit
'
'Public Function ReadFileTsvdb(pfn As PathFileName, PropertyLists_out As List) As Boolean
'Try: On Error GoTo Catch
'    Dim lines() As String
'    If Not pfn.TryReadAllLines(lines) Then
'        MsgBox "Could not read the file: " & vbCrLf & pfn.Value
'        Exit Function
'    End If
'    Set PropertyLists = MNew.List(vbObject) 'Of List
'    Dim i As Long, u As Long: u = UBound(lines)
'    Dim line As String
'    Dim pkl As List ': Set pkl = MNew.List(vbObject)
'    Dim pke As PropKeyHEntry
'    For i = 0 To u
'        line = lines(i)
'        Dim data() As String: data = Split(line, vbTab)
'        Dim s As String: s = Left(line, 1)
'        If s = "#" Then
'            'just a comment go on
'        ElseIf s = vbTab Then
'        'If Left(line, 1) = vbTab Then
'            Set pke = MNew.PropKeyHEntry_Parse(data)
'            If Not pkl Is Nothing Then
'                pkl.Add pke
'            End If
'        Else
'            Set pkl = PropertyLists.Add(MNew.List(vbObject))
'            pkl.Name = Trim(line)
'        End If
'    Next
'    Set TsvbdPFN = pfn
'    ReadFileTsvdb = True
'Catch:
'    '
'End Function
''
''Public Function ReadTsvdb(lines() As String) As Boolean
''    '
''End Function
'
'Public Function WriteTsvDB(tsvdb As PathFileName, PropertyLists As List) As Boolean
'    'If Not tsvdb.Exists Then Exit Function
'Try: On Error GoTo Catch
'    tsvdb.WriteStr PropertyLists_ToStr(PropertyLists)
'    GoTo Finally
'Catch:
'    MErr.MessError "MApp", "WriteTSVDB", "Could not write to file: " & vbCrLf & tsvdb.Value
'Finally:
'End Function
'
'Function PropertyLists_ToStr(this As List) As String 'This=PropertyLists As List(Of List(Of PropKeyHEntry))
'    'this = PropertyLists
'    Dim sb As New StringBuilder
'    Dim i As Long, pkl As List
'    For i = 0 To this.Count - 1
'        Set pkl = this.Item(i)
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
'    PropertyLists_ToStr = sb.ToStr
'End Function
'
