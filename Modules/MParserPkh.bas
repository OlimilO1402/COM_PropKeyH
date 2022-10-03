Attribute VB_Name = "MParserPkh"
Option Explicit

Public Function TryParse(aPFN As PathFileName, PKLists_out As List) As Boolean
Try: On Error GoTo Catch
    If LCase(aPFN.Extension) <> ".h" Then Exit Function
    Dim lines() As String
    If Not aPFN.TryReadAllLines(lines) Then
        MsgBox "Could not read the file: " & vbCrLf & aPFN.Value
        aPFN.CloseFile
        Exit Function
    End If
    Set PKLists_out = MNew.List(EDataType.vbObject) '(Of List(Of PropKeyHEntry))
    TryParse = ReadLines(PKLists_out, lines)
Catch:
Finally:
    aPFN.CloseFile
End Function

Public Function PropKeyHEntry_ParseFromTsvDB(data() As String) As PropKeyHEntry
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

Private Function ReadLines(PropLists As List, lines() As String) As Boolean
    'OK we try to read until begin of one PKList, PKLists starting always with "//--------"
    'then we read entry by entry
Try: On Error GoTo Catch
    Dim line As String
    Dim i As Long, u As Long: u = UBound(lines)
    Do While i < u
        line = Trim$(lines(i)): i = i + 1
        If Len(line) Then
            If IsPKListStart(line) Then
                line = Trim$(lines(i)): i = i + 1
                If Len(line) Then
                    Dim PropListName As String: PropListName = Trim$(Mid$(line, 3))
                    Dim pkl As List: Set pkl = PropLists.Add(MNew.List(EDataType.vbObject))  'Of PropKeyHEntry
                    pkl.Name = PropListName
                    Do While i < u
                        line = Trim$(lines(i)): i = i + 1
                        If Len(line) Then
                            If IsPKEntryStart(line) Then
                                'OK now jump one line back
                                i = i - 1
                                Do While i < u
                                    line = lines(i): i = i + 1: If i > u Then Exit Function
                                    If IsPKListStart(line) Then Exit Do
                                    If IsPKEntryStart(line) Then
                                        i = i - 1
                                        Dim pkhe As PropKeyHEntry: Set pkhe = Parse_PropKeyHEntry(lines, i, u)
                                        If Not pkhe Is Nothing Then pkl.Add pkhe
                                        i = i + 1
                                    End If
                                Loop
                                If IsPKListStart(line) Then
                                    i = i - 1
                                    Exit Do
                                End If
                            End If
                        End If
                    Loop
                End If
            End If
        End If
    Loop
    ReadLines = True
Catch:
End Function

Private Function IsPKListStart(line As String) As Boolean
    IsPKListStart = Left(line, 10) = "//--------"
End Function

Private Function IsPKEntryStart(line As String) As Boolean
    IsPKEntryStart = Left(line, 10) = "//  Name: "
End Function

Private Function Parse_PropKeyHEntry(lines() As String, i As Long, u As Long) As PropKeyHEntry
    Dim pkhe As PropKeyHEntry: Set pkhe = New PropKeyHEntry
    Set Parse_PropKeyHEntry = pkhe
    Dim line As String, sa() As String
    
    line = lines(i): i = i + 1: If i > u Then Exit Function
    
    If Left(line, 10) = "//  Name: " Then
        line = Mid(line, 10)
        line = Trim(line): sa = Split(line, "--")
        If UBound(sa) >= 0 Then
            pkhe.Name = Trim(sa(0))
            If UBound(sa) >= 1 Then
                pkhe.PKEYName = Trim(sa(1))
            End If
        End If
    Else
        If IsPKListStart(line) Or IsPKEntryStart(line) Then
            i = i - 1
            Exit Function
        End If
    End If
    
    line = lines(i): i = i + 1: If i > u Then Exit Function
    
    If Left(line, 10) = "//  Type: " Then
        line = Mid(line, 10)
        line = Trim(line): sa = Split(line, "--")
        If UBound(sa) >= 0 Then
            pkhe.DataType = Trim(sa(0))
            If UBound(sa) >= 1 Then
                pkhe.PKVarTyp = Trim(sa(1))
            End If
        End If
    Else
        If IsPKListStart(line) Or IsPKEntryStart(line) Then
            i = i - 1
            Exit Function
        End If
    End If
    
    line = lines(i): i = i + 1: If i > u Then Exit Function
    Dim saa() As String
    
    If Left(line, 14) = "//  FormatID: " Then
        line = Mid(line, 14)
        line = Trim(line): sa = Split(line, ",")
        'now have a look if FMTID is even contained
        'or Guid only
        If UBound(sa) >= 0 Then
            line = Trim(sa(0))
            If Left(line, 1) = "(" Then
                saa = Split(line, ") ")
                If UBound(saa) >= 0 Then
                    pkhe.FormatID = Mid(Trim(saa(0)), 2)
                    If UBound(saa) >= 1 Then
                        pkhe.FmtGuid = Trim(saa(1))
                    End If
                End If
            Else
                pkhe.FmtGuid = Trim(line)
            End If
        End If
        If UBound(sa) >= 1 Then
            line = Trim(sa(1))
            saa = Split(line, " (")
            If UBound(saa) >= 0 Then
                pkhe.PIDValue = Trim(saa(0))
                If UBound(saa) >= 1 Then
                    pkhe.PIDName = Left(Trim(saa(1)), Len(saa(1)) - 1)
                End If
            End If
        End If
    Else
        If IsPKListStart(line) Or IsPKEntryStart(line) Then
            i = i - 1
            Exit Function
        End If
    End If
    
    'ZeileNr nochmal um 1 erhöhen
    i = i + 1
    line = lines(i)
    If Left(line, 4) = "//  " Then
        pkhe.Descript = Trim(Mid(line, 4))
    Else
        If IsPKListStart(line) Or IsPKEntryStart(line) Then
            i = i - 1
            Exit Function
        End If
    End If
End Function


