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

Public Name     As String ' System.Audio.ChannelCount
Public PKEYName As String ' PKEY_Audio_ChannelCount
Public DataType As String ' UInt32
Public PKVarTyp As String ' VT_UI4
Public FormatID As String ' FMTID_AudioSummaryInformation
Public FmtGuid  As String ' 64440490-4C8B-11D1-8B70-080036B11A03
Public PIDName  As String ' PIDASI_CHANNEL_COUNT
Public PIDValue As String ' 7
Public Descript As String ' Indicates the channel count for the audio file. Values: 1 (mono), 2 (stereo).

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
