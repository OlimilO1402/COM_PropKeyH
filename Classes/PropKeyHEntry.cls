VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PropKeyHEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Name     As String ' e.g.: System.Audio.ChannelCount
Public PKEYName As String ' e.g.: PKEY_Audio_ChannelCount
Public DataType As String ' e.g.: UInt32
Public PKVarTyp As String ' e.g.: VT_UI4
Public FormatID As String ' e.g.: FMTID_AudioSummaryInformation
Public FmtGuid  As String ' e.g.: 64440490-4C8B-11D1-8B70-080036B11A03
Public PIDName  As String ' e.g.: PIDASI_CHANNEL_COUNT
Public PIDValue As String ' e.g.: 7
Public Descript As String ' e.g.: Indicates the channel count for the audio file. Values: 1 (mono), 2 (stereo).

'//-----------------------------------------------------------------------------
'// Audio properties
'
'//  Name:     System.Audio.ChannelCount -- PKEY_Audio_ChannelCount
'//  Type:     UInt32 -- VT_UI4
'//  FormatID: (FMTID_AudioSummaryInformation) 64440490-4C8B-11D1-8B70-080036B11A03, 7 (PIDASI_CHANNEL_COUNT)
'//
'//  Indicates the channel count for the audio file.  Values: 1 (mono), 2 (stereo).
'DEFINE_PROPERTYKEY(PKEY_Audio_ChannelCount, 0x64440490, 0x4C8B, 0x11D1, 0x8B, 0x70, 0x08, 0x00, 0x36, 0xB1, 0x1A, 0x03, 7);
'
'// Possible discrete values for PKEY_Audio_ChannelCount are:
'#define AUDIO_CHANNELCOUNT_MONO             1ul
'#define AUDIO_CHANNELCOUNT_STEREO           2ul
'
'//  Name:     System.Audio.Compression -- PKEY_Audio_Compression
'//  Type:     String -- VT_LPWSTR  (For variants: VT_BSTR)
'//  FormatID: (FMTID_AudioSummaryInformation) 64440490-4C8B-11D1-8B70-080036B11A03, 10 (PIDASI_COMPRESSION)
'//
'//
'DEFINE_PROPERTYKEY(PKEY_Audio_Compression, 0x64440490, 0x4C8B, 0x11D1, 0x8B, 0x70, 0x08, 0x00, 0x36, 0xB1, 0x1A, 0x03, 10);
'
'//  Name:     System.Audio.EncodingBitrate -- PKEY_Audio_EncodingBitrate
'//  Type:     UInt32 -- VT_UI4
'//  FormatID: (FMTID_AudioSummaryInformation) 64440490-4C8B-11D1-8B70-080036B11A03, 4 (PIDASI_AVG_DATA_RATE)
'//
'//  Indicates the average data rate in Hz for the audio file in "bits per second".
'DEFINE_PROPERTYKEY(PKEY_Audio_EncodingBitrate, 0x64440490, 0x4C8B, 0x11D1, 0x8B, 0x70, 0x08, 0x00, 0x36, 0xB1, 0x1A, 0x03, 4);

Friend Function ToStr() As String
    ToStr = Name & vbTab & _
            PKEYName & vbTab & _
            DataType & vbTab & _
            PKVarTyp & vbTab & _
            FormatID & vbTab & _
            FmtGuid & vbTab & _
            PIDName & vbTab & _
            PIDValue & vbTab & _
            Descript '& vbCrLf
End Function

Public Function Name_ToEnumConstname() As String
    Name_ToEnumConstname = Replace(Me.Name, ".", "_")
End Function

Public Function ToEnumElement(ByVal MaxNameLen As Long) As String
    'the name and maybe other infos too?
    ToEnumElement = Name_ToEnumConstname ' MString.PadRight(Me.Name_ToEnumConstname, MaxNameLen) & " ' " & Me.FmtGuid & ", " & MString.PadRight(CStr(Me.PIDValue), 5) & " | " & Me.DataType
End Function

Public Function ToSubInit() As String
    'm_PropKeys(i) = New_PROPERTYKEY("", 0): i = i + 1
    'm_PKeys(i) = PKEY("{64440490-4C8B-11D1-8B70-080036B11A03}", 7): i = i + 1
    ToSubInit = "m_PKeys(i) = NamedPKEY(""" & Me.Name & """, """ & Me.FmtGuid & """, " & Me.PIDValue & "): i = i + 1"
End Function
'Byte              :: VT_UI1
'Int16             :: VT_I2
'Int32             :: VT_I4
'UInt16            :: VT_UI2
'UInt32            :: VT_UI4
'UInt64            :: VT_UI8
'Boolean           :: VT_BOOL
'Stream            :: VT_STREAM
'Blob              :: VT_BLOB
'Guid              :: VT_CLSID
'Null              :: VT_NULL
'Object            :: VT_UNKNOWN
'Clipboard         :: VT_CF
'Double            :: VT_R8
'DateTime          :: VT_FILETIME                (For variants: VT_DATE            )

'Any               :: VT_NULL                                                         Legacy code may treat this as VT_UNKNOWN.
'Buffer            :: VT_VECTOR | VT_UI1         (For variants: VT_ARRAY | VT_UI1  )
'String            :: VT_LPWSTR                  (For variants: VT_BSTR            )  Legacy code may treat this as VT_LPSTR.
'String            :: VT_LPWSTR                  (For variants: VT_BSTR            )
'String            :: VT_LPWSTR                  (For variants: VT_BSTR            )  Legacy code may treat this as VT_BSTR.
'Multivalue String :: VT_VECTOR | VT_LPWSTR      (For variants: VT_ARRAY | VT_BSTR )
'Multivalue UInt64 :: VT_VECTOR | VT_UI8         (For variants: VT_ARRAY | VT_UI8  )
'Multivalue String :: VT_VECTOR | VT_LPWSTR      (For variants: VT_ARRAY | VT_BSTR )  Legacy code may treat this as VT_LPSTR.
'Multivalue Guid   :: VT_VECTOR | VT_CLSID       (For variants: VT_ARRAY | VT_CLSID)
'Multivalue Double :: VT_VECTOR | VT_R8          (For variants: VT_ARRAY | VT_R8   )
'Multivalue UInt32 :: VT_VECTOR | VT_UI4         (For variants: VT_ARRAY | VT_UI4  )
'Multivalue Any    :: VT_VECTOR | VT_NULL        (For variants: VT_ARRAY | VT_NULL )

Public Property Get EVbVarType() As EVbVarType
    'Dim vt As EVbVarType
    'If CheckEVbVarType(vt) Then EVbVarType = vt
    EVbVarType = MEVbVarType.EVbVarType_Parse(DataType)
End Property

Private Function CheckEVbVarType(vt_out As EVbVarType) As Boolean
    Dim vt0 As EVbVarType: vt0 = DataTypeToEVbVarType
    Dim vt1 As EVbVarType: vt1 = PKVarTypToEVbVarType
    CheckEVbVarType = vt0 = vt1
    vt_out = vt0
End Function

Private Function DataTypeToEVbVarType() As EVbVarType
    DataTypeToEVbVarType = MEVbVarType.EVbVarType_Parse(DataType)
End Function

Private Function PKVarTypToEVbVarType() As EVbVarType
    Dim s As String: s = Trim(PKVarTyp)
    Dim vt As EVbVarType
    vt = MEVbVarType.EVbVarType_Parse(s)
    If vt <> vbEmpty Then PKVarTypToEVbVarType = vt: Exit Function
    Dim sa() As String: sa = Split(PKVarTyp, " ")
    Dim i As Long, u As Long: u = UBound(sa)
    
    'i = 0
    If i > u Then PKVarTypToEVbVarType = vt: Exit Function
    vt = vt Or MEVbVarType.EVbVarType_Parse(sa(i))
    
    i = i + 1
    If i > u Then PKVarTypToEVbVarType = vt: Exit Function
    If sa(i) = "|" Then
    
        i = i + 1
        If i > u Then PKVarTypToEVbVarType = vt: Exit Function
        vt = vt Or MEVbVarType.EVbVarType_Parse(sa(i))
        Exit Function
        
    End If
    
    i = i + 1
    If i > u Then PKVarTypToEVbVarType = vt: Exit Function
    If sa(i) = "(For" Then
    
        i = i + 1
        If i > u Then PKVarTypToEVbVarType = vt: Exit Function
        If sa(i) = "variants:" Then
            vt = vt Or MEVbVarType.EVbVarType_Parse(sa(i))
        End If
        
        i = i + 1
        If i > u Then PKVarTypToEVbVarType = vt: Exit Function
        If sa(i) = "|" Then
        
            i = i + 1
            If i > u Then PKVarTypToEVbVarType = vt: Exit Function
            vt = vt Or MEVbVarType.EVbVarType_Parse(sa(i))
            Exit Function
            
        End If
    
    End If

End Function

'Public Function ToVVariant() As VVariant
'    Dim vt As EVbVarType: vt = MEVbVarType.EVbVarType_Parse(Me.PKVarTyp)
'    Set ToVVariant = MNew.VVariantVt(vt)
'End Function
'
'Public Function ToPossFileProperty() As FileProperty
'    Set ToPossFileProperty = MNew.FileProperty(Me.ToPropKey, Me.ToVVariant)
'End Function
