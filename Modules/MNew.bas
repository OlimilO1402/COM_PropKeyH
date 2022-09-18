Attribute VB_Name = "MNew"
Option Explicit

Public Function PathFileName(ByVal aPathFileName As String, _
                    Optional ByVal aFileName As String, _
                    Optional ByVal aExt As String) As PathFileName
    Set PathFileName = New PathFileName: PathFileName.New_ aPathFileName, aFileName, aExt
End Function

Public Function List(Of_T As EDataType, _
                     Optional ArrColStrTypList, _
                     Optional ByVal IsHashed As Boolean = False, _
                     Optional ByVal Capacity As Long = 32, _
                     Optional ByVal GrowRate As Single = 2, _
                     Optional ByVal GrowSize As Long = 0) As List
    Set List = New List: List.New_ Of_T, ArrColStrTypList, IsHashed, Capacity, GrowRate, GrowSize
End Function

Public Function PropKeyHEntry_Parse(data() As String) As PropKeyHEntry
    Dim i As Long, u As Long: u = UBound(data)
    Dim pke As PropKeyHEntry: Set pke = New PropKeyHEntry
    With pke
        i = i + 1 'the first column is empty, because it contains the name of the group of entries
        If i <= u Then .Name = data(i): i = i + 1 ' System.Audio.ChannelCount
        If i <= u Then .PKEYName = data(i): i = i + 1  ' PKEY_Audio_ChannelCount
        If i <= u Then .DataType = data(i): i = i + 1  ' UInt32
        If i <= u Then .PKVarTyp = data(i): i = i + 1  ' VT_UI4
        If i <= u Then .FormatID = data(i): i = i + 1  ' FMTID_AudioSummaryInformation
        If i <= u Then .FmtGuid = data(i): i = i + 1   ' 64440490-4C8B-11D1-8B70-080036B11A03
        If i <= u Then .PIDName = data(i): i = i + 1   ' PIDASI_CHANNEL_COUNT
        If i <= u Then .PIDValue = data(i): i = i + 1  ' 7
        If i <= u Then .Descript = data(i): i = i + 1  ' Indicates the channel count for the audio file. Values: 1 (mono), 2 (stereo).
    End With
    Set PropKeyHEntry_Parse = pke
End Function

'Public Function DocumentPkh(aPFN As PathFileName, Optional ByVal aPropertyLists As List = Nothing) As DocumentPkh
'    Set DocumentPkh = New DocumentPkh: DocumentPkh.New_ aPFN, aPropertyLists
'End Function

Public Function DocumentTsv(aPFN As PathFileName, Optional ByVal aPropertyLists As List = Nothing) As DocumentTsv
    Set DocumentTsv = New DocumentTsv: DocumentTsv.New_ aPFN, aPropertyLists
End Function

'Public Function DocumentTsvPL(ByVal aPropertyLists As List) As DocumentTsv
'    Set DocumentTsv = New DocumentTsv: DocumentTsv.NewPL aPropertyLists
'End Function


