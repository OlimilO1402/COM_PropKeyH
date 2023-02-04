Attribute VB_Name = "MApp"
Option Explicit
Private m_DocumentTsv As DocumentTsv

Sub Main()
    MEPropertyKey.Init
    FMain.Show
End Sub

Public Property Get PropertyLists() As List '(Of List(Of PropKeyHEntry))
    Set PropertyLists = m_DocumentTsv.PropertyLists
End Property

Public Property Get DocumentTsv() As DocumentTsv
    Set DocumentTsv = m_DocumentTsv
End Property

Public Property Get EnvironHomePathDocs() As String
    EnvironHomePathDocs = Environ("HOMEDRIVE") & Environ("HOMEPATH") & "\Documents\"
End Property

Sub DataClear()
    m_DocumentTsv.PropertyLists.Clear
End Sub

Public Function Max(V1, V2)
    If V1 > V2 Then Max = V1 Else Max = V2
End Function

Public Function AutomaticOpenNReadPropKeyHFile(Optional bLoud As Boolean = False) As Boolean
    Dim aPFN As PathFileName
    If Not TrySearchFile(aPFN, bLoud) Then Exit Function
    Dim PKLists As List, Ex As String: Ex = LCase(aPFN.Extension)
    If Ex = ".h" Then
        Debug.Print aPFN.Value
        AutomaticOpenNReadPropKeyHFile = MParserPkh.TryParse(aPFN, PKLists)
        Set m_DocumentTsv = MNew.DocumentTsv(aPFN, PKLists)
    ElseIf Ex = ".tsvdb" Then
        Set m_DocumentTsv = MNew.DocumentTsv(aPFN, PKLists)
        AutomaticOpenNReadPropKeyHFile = m_DocumentTsv.Read
    End If
End Function

Private Function TrySearchFile(PFN_out As PathFileName, Optional bLoud As Boolean = False) As Boolean
    Dim sEr As String: sEr = "File not found: "
    Dim PNm As String, FNm As String: FNm = "propkey.h"
    
    Set PFN_out = New PathFileName
    
    'we are confident to find the right file, so at first we assume to be successfull:
    TrySearchFile = True
    
    PNm = App.Path & "\"
    PFN_out.Value = PNm & FNm
    If PFN_out.Exists Then Exit Function
    sEr = sEr & vbCrLf & PFN_out.Value
    
    PNm = PNm & "Resources\"
    PFN_out.Value = PNm & FNm
    If PFN_out.Exists Then Exit Function
    sEr = sEr & vbCrLf & PFN_out.Value
    
    PNm = PNm & "Strings\"
    PFN_out.Value = PNm & FNm
    If PFN_out.Exists Then Exit Function
    sEr = sEr & vbCrLf & PFN_out.Value
    
    'now we search in User-path
    PNm = EnvironHomePathDocs
    PFN_out.Value = PNm & FNm
    If PFN_out.Exists Then Exit Function
    sEr = sEr & vbCrLf & PFN_out.Value
    
    'OK found, now write the file in Resources to the user path to read from there
    Dim fc As String: fc = StrConv(LoadResData(1, "CUSTOM"), vbUnicode)
    If PFN_out.WriteStr(fc) Then
        PFN_out.CloseFile
        Exit Function
    Else
        If bLoud Then MsgBox sEr & vbCrLf & "Could not write to file: " & vbCrLf & PFN_out.Value
    End If
    
    'no we search for tsvdb-file
    FNm = "PropKey.tsvdb"
    sEr = "File not found: "
    
    PNm = App.Path & "\"
    PFN_out.Value = PNm & FNm
    If PFN_out.Exists Then Exit Function
    sEr = sEr & vbCrLf & PFN_out.Value
    
    PNm = PNm & "Resources\"
    PFN_out.Value = PNm & FNm
    If PFN_out.Exists Then Exit Function
    sEr = sEr & vbCrLf & PFN_out.Value
    
    PNm = PNm & "Strings\"
    PFN_out.Value = PNm & FNm
    If PFN_out.Exists Then Exit Function
    sEr = sEr & vbCrLf & PFN_out.Value
    
    'jetzt noch den UserPfad untersuchen
    PNm = EnvironHomePathDocs
    PFN_out.Value = PNm & FNm
    If PFN_out.Exists Then Exit Function
    sEr = sEr & vbCrLf & PFN_out.Value
    
    If bLoud Then MsgBox sEr & vbCrLf & "Could not write to file: " & vbCrLf & PFN_out.Value
    
    TrySearchFile = False
End Function

Public Function ReadFile(aPFN As PathFileName) As Boolean
    Dim PKLists As List
    Dim Ex As String: Ex = LCase(aPFN.Extension)
    If Ex = ".h" Then
        ReadFile = MParserPkh.TryParse(aPFN, PKLists)
        If ReadFile Then
            Set m_DocumentTsv = MNew.DocumentTsv(aPFN, PKLists)
        Else
            MsgBox "Errors reading file: " & vbCrLf & aPFN.Value
            Exit Function
        End If
    ElseIf Ex = ".tsvdb" Then
        Set m_DocumentTsv = MNew.DocumentTsv(aPFN)
        ReadFile = m_DocumentTsv.Read
    End If
End Function

Public Function WriteFile(tsvdb_pfn As PathFileName) As Boolean
    Debug.Print tsvdb_pfn.Value
    m_DocumentTsv.FileName.Value = tsvdb_pfn.Value
    WriteFile = m_DocumentTsv.WriteTsvDB
End Function

Public Sub ClipboardCopyAll()
    If m_DocumentTsv Is Nothing Then Set m_DocumentTsv = MNew.DocumentTsv(Nothing, PropertyLists)
    Dim s As String: s = m_DocumentTsv.PropertyListsToStr
    Clipboard.SetText s
End Sub

Public Function PropertyList_ToStr(pkl As List) As String
    If pkl.Count = 0 Then Exit Function
    Dim s As String
    Dim i As Long, pkhe As PropKeyHEntry
    For i = 0 To pkl.Count - 1
        Set pkhe = pkl.Item(i)
        s = s & pkhe.ToStr & vbCrLf
    Next
    PropertyList_ToStr = s
End Function

Public Function StatsAllDifDatatypes() As List
    Dim datatypes As List: Set datatypes = MNew.List(EDataType.vbObject, , True)
    Dim i As Long, pkl As List
    For i = 0 To PropertyLists.Count - 1
        Set pkl = PropertyLists.Item(i)
        Dim j As Long, pkhe As PropKeyHEntry
        Dim dtyp As String
        For j = 0 To pkl.Count - 1
            Set pkhe = pkl.Item(j)
            If Not pkhe Is Nothing Then
                dtyp = pkhe.DataType & " " & pkhe.PKVarTyp
                If Not datatypes.ContainsKey(dtyp) Then
                    datatypes.Add pkhe, dtyp
                End If
            End If
        Next
    Next
    Set StatsAllDifDatatypes = datatypes
End Function


