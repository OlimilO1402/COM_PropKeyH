VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DocumentPkh"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_PFN           As PathFileName
Private m_PropertyLists As List '(Of List(Of PropKeyHEntry))

Friend Sub New_(aPFN As PathFileName, Optional ByVal aPropertyLists As List = Nothing)
    Set m_PFN = aPFN: Set m_PropertyLists = aPropertyLists
    If UCase(m_PFN.Extension) <> "h" Then m_PFN.Extension = "h"
End Sub

Public Property Get FileName() As PathFileName
    Set FileName = m_PFN
End Property

Public Function ToTsvDB() As DocumentTsv
    Set ToTsvDB = MNew.DocumentTsv(m_PFN, m_PropertyLists)
End Function

Public Property Get PropertyLists() As List '(Of List(Of PropKeyHEntry))
    Set PropertyLists = m_PropertyLists
End Property

Public Function Read() As Boolean
    Dim lines() As String
    If Not m_PFN.TryReadAllLines(lines) Then
        MsgBox "Could not read the file: " & vbCrLf & m_PFN.Value
        Exit Function
    End If
    Set m_PropertyLists = MNew.List(vbObject) '(Of List(Of PropKeyHEntry))
    Read = ReadLines(lines)
End Function

Private Function ReadLines(lines() As String) As Boolean
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

