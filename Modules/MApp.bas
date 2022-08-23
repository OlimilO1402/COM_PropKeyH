Attribute VB_Name = "MApp"
Option Explicit
'Public PropKeyH As PathFileName
Public PropertyLists As List '(Of List (Of PropKeyHEntry))

Sub Main()
'    If AutomaticOpenNReadPropKeyHFile(True) Then
    FMain.Show
'        FMain.UpdateView PropertyLists
'    End If
End Sub

Sub DataClear()
    Set PropertyLists = MNew.List(vbObject)
End Sub

Public Function AutomaticOpenNReadPropKeyHFile(Optional bLoud As Boolean = False) As Boolean
    Dim PropKeyH As PathFileName
    Set PropKeyH = MNew.PathFileName(App.Path & "\propkey.h")
    If Not PropKeyH.Exists Then
        Set PropKeyH = MNew.PathFileName(App.Path & "\Resources\propkey.h")
        If Not PropKeyH.Exists Then
            If bLoud Then
                MsgBox "File not found: " & vbCrLf & _
                        App.Path & "\propkey.h" & vbCrLf & _
                        App.Path & "\Resources\propkey.h"
            End If
            Exit Function
        End If
    End If
    AutomaticOpenNReadPropKeyHFile = MApp.ReadFile(PropKeyH)
    'Dim s As String
    's = MApp.PropertyLists_ToStr
    'Text1.Text = s
    'Clipboard.SetText s
End Function

Public Function ReadFile(pfn As PathFileName) As Boolean
    
    Dim lines() As String
    If Not pfn.TryReadAllLines(lines) Then
        MsgBox "Could not read the file: " & vbCrLf & pfn.value
        Exit Function
    End If
    Set PropertyLists = MNew.List(vbObject) 'Of List
    ReadFile = ReadLines(lines)
    
End Function

Public Function WriteTSVDB(tsvdb As PathFileName) As Boolean
    'If Not tsvdb.Exists Then Exit Function
Try: On Error GoTo Catch
    tsvdb.WriteStr PropertyLists_ToStr
    GoTo Finally
Catch:
    MErr.MessError "MApp", "WriteTSVDB", "Could not write to file: " & vbCrLf & tsvdb.value
Finally:
End Function

Public Function ReadLines(lines() As String) As Boolean
'OK hierin lesen wir bis zum Beginn einer PKListe die startet immer mit "//--------"
'dann lesen wir die Liste Entry für Entry

    Dim line As String
    
    Dim i As Long, u As Long: u = UBound(lines)
    Do While i < u
        line = Trim$(lines(i)): i = i + 1
        If Len(line) Then
            If IsPKListStart(line) Then
                line = Trim$(lines(i)): i = i + 1
                If Len(line) Then
                    Dim PropListName As String: PropListName = Trim$(Mid$(line, 3))
                    Dim pkl As List: Set pkl = PropertyLists.Add(MNew.List(vbObject))  'Of PropKeyHEntry
                    pkl.Name = PropListName
                    'Debug.Print PropListName
                    Do While i < u
                        line = Trim$(lines(i)): i = i + 1
                        If Len(line) Then
                            If IsPKEntryStart(line) Then
                                'OK jetzt nochmal um eine Zeile zurückgehen
                                i = i - 1
                                Do While i < u
                                    line = lines(i): i = i + 1: If i > u Then Exit Function
                                    If IsPKListStart(line) Then Exit Do
                                    If IsPKEntryStart(line) Then
                                        i = i - 1
                                        Dim pkhe As PropKeyHEntry: Set pkhe = Parse_PropKeyHEntry(lines, i, u)
                                        If Not pkhe Is Nothing Then pkl.Add pkhe
                                        'If Not pkhe Is Nothing Then Debug.Print pkhe.ToStr
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
End Function

Public Function IsPKListStart(line As String) As Boolean
    IsPKListStart = Left(line, 10) = "//--------"
End Function

Public Function IsPKEntryStart(line As String) As Boolean
    IsPKEntryStart = Left(line, 10) = "//  Name: "
End Function

Public Function Parse_PropKeyHEntry(lines() As String, i As Long, u As Long) As PropKeyHEntry
    Dim pkhe As PropKeyHEntry: Set pkhe = New PropKeyHEntry
    Set Parse_PropKeyHEntry = pkhe
    Dim line As String, sa() As String
    
    line = lines(i): i = i + 1: If i > u Then Exit Function
    
    If Left(line, 10) = "//  Name: " Then
        line = Mid(line, 10)
        line = Trim(line): sa = Split(line, "--")
        If UBound(sa) >= 0 Then
            pkhe.Name = Trim(sa(0))
            'If pkhe.Name = "System.Audio.IsVariableBitRate" Then
            '    Debug.Print pkhe.Name
            'End If
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
        'jetzt zuerst schauen ob überhaupt FMTID enthalten ist
        'oder nur die Guid
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

Function PropertyLists_ToStr() As String
    Dim sb As New StringBuilder
    Dim i As Long, pkl As List
    For i = 0 To PropertyLists.Count - 1
        Set pkl = PropertyLists.Item(i)
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
    PropertyLists_ToStr = sb.ToStr
End Function

Public Function PropertyList_ToStr(pkl As List) As String
    If pkl.Count = 0 Then Exit Function
    'Dim sb As StringBuilder: Set sb = New StringBuilder
    Dim s As String
    Dim i As Long, pkhe As PropKeyHEntry
    'For Each pkhe In pkl.GetEnumerator
    For i = 0 To pkl.Count - 1
        Set pkhe = pkl.Item(i)
        'sb.AppendLine pkhe.ToStr
        s = s & pkhe.ToStr & vbCrLf
    Next
    'PropertyList_ToStr = sb.ToStr
    PropertyList_ToStr = s
End Function
