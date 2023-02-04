Attribute VB_Name = "MEPropertyKey"
Option Explicit
'This module, all the elements enum types and functions are created
Public Enum EPropertyKeys
    ' Audio properties
    System_Audio_ChannelCount
    System_Audio_Compression
    System_Audio_EncodingBitrate
    System_Audio_Format
    System_Audio_IsVariableBitRate
    System_Audio_PeakValue
    System_Audio_SampleRate
    System_Audio_SampleSize
    System_Audio_StreamName
    System_Audio_StreamNumber
    ' Calendar properties
    System_Calendar_Duration
    System_Calendar_IsOnline
    System_Calendar_IsRecurring
    System_Calendar_Location
    System_Calendar_OptionalAttendeeAddresses
    System_Calendar_OptionalAttendeeNames
    System_Calendar_OrganizerAddress
    System_Calendar_OrganizerName
    System_Calendar_ReminderTime
    System_Calendar_RequiredAttendeeAddresses
    System_Calendar_RequiredAttendeeNames
    System_Calendar_Resources
    System_Calendar_ResponseStatus
    System_Calendar_ShowTimeAs
    System_Calendar_ShowTimeAsText
    ' Communication properties
    System_Communication_AccountName
    System_Communication_DateItemExpires
    System_Communication_Direction
    System_Communication_FollowupIconIndex
    System_Communication_HeaderItem
    System_Communication_PolicyTag
    System_Communication_SecurityFlags
    System_Communication_Suffix
    System_Communication_TaskStatus
    System_Communication_TaskStatusText
    ' Computer properties
    System_Computer_DecoratedFreeSpace
    ' Contact properties
    System_Contact_AccountPictureDynamicVideo
    System_Contact_AccountPictureLarge
    System_Contact_AccountPictureSmall
    System_Contact_Anniversary
    System_Contact_AssistantName
    System_Contact_AssistantTelephone
    System_Contact_Birthday
    System_Contact_BusinessAddress
    System_Contact_BusinessAddress1Country
    System_Contact_BusinessAddress1Locality
    System_Contact_BusinessAddress1PostalCode
    System_Contact_BusinessAddress1Region
    System_Contact_BusinessAddress1Street
    System_Contact_BusinessAddress2Country
    System_Contact_BusinessAddress2Locality
    System_Contact_BusinessAddress2PostalCode
    System_Contact_BusinessAddress2Region
    System_Contact_BusinessAddress2Street
    System_Contact_BusinessAddress3Country
    System_Contact_BusinessAddress3Locality
    System_Contact_BusinessAddress3PostalCode
    System_Contact_BusinessAddress3Region
    System_Contact_BusinessAddress3Street
    System_Contact_BusinessAddressCity
    System_Contact_BusinessAddressCountry
    System_Contact_BusinessAddressPostalCode
    System_Contact_BusinessAddressPostOfficeBox
    System_Contact_BusinessAddressState
    System_Contact_BusinessAddressStreet
    System_Contact_BusinessEmailAddresses
    System_Contact_BusinessFaxNumber
    System_Contact_BusinessHomePage
    System_Contact_BusinessTelephone
    System_Contact_CallbackTelephone
    System_Contact_CarTelephone
    System_Contact_Children
    System_Contact_CompanyMainTelephone
    System_Contact_ConnectedServiceDisplayName
    System_Contact_ConnectedServiceIdentities
    System_Contact_ConnectedServiceName
    System_Contact_ConnectedServiceSupportedActions
    System_Contact_DataSuppliers
    System_Contact_Department
    System_Contact_DisplayBusinessPhoneNumbers
    System_Contact_DisplayHomePhoneNumbers
    System_Contact_DisplayMobilePhoneNumbers
    System_Contact_DisplayOtherPhoneNumbers
    System_Contact_EmailAddress
    System_Contact_EmailAddress2
    System_Contact_EmailAddress3
    System_Contact_EmailAddresses
    System_Contact_EmailName
    System_Contact_FileAsName
    System_Contact_FirstName
    System_Contact_FullName
    System_Contact_Gender
    System_Contact_GenderValue
    System_Contact_Hobbies
    System_Contact_HomeAddress
    System_Contact_HomeAddress1Country
    System_Contact_HomeAddress1Locality
    System_Contact_HomeAddress1PostalCode
    System_Contact_HomeAddress1Region
    System_Contact_HomeAddress1Street
    System_Contact_HomeAddress2Country
    System_Contact_HomeAddress2Locality
    System_Contact_HomeAddress2PostalCode
    System_Contact_HomeAddress2Region
    System_Contact_HomeAddress2Street
    System_Contact_HomeAddress3Country
    System_Contact_HomeAddress3Locality
    System_Contact_HomeAddress3PostalCode
    System_Contact_HomeAddress3Region
    System_Contact_HomeAddress3Street
    System_Contact_HomeAddressCity
    System_Contact_HomeAddressCountry
    System_Contact_HomeAddressPostalCode
    System_Contact_HomeAddressPostOfficeBox
    System_Contact_HomeAddressState
    System_Contact_HomeAddressStreet
    System_Contact_HomeEmailAddresses
    System_Contact_HomeFaxNumber
    System_Contact_HomeTelephone
    System_Contact_IMAddress
    System_Contact_Initials
    System_Contact_JA_CompanyNamePhonetic
    System_Contact_JA_FirstNamePhonetic
    System_Contact_JA_LastNamePhonetic
    System_Contact_JobInfo1CompanyAddress
    System_Contact_JobInfo1CompanyName
    System_Contact_JobInfo1Department
    System_Contact_JobInfo1Manager
    System_Contact_JobInfo1OfficeLocation
    System_Contact_JobInfo1Title
    System_Contact_JobInfo1YomiCompanyName
    System_Contact_JobInfo2CompanyAddress
    System_Contact_JobInfo2CompanyName
    System_Contact_JobInfo2Department
    System_Contact_JobInfo2Manager
    System_Contact_JobInfo2OfficeLocation
    System_Contact_JobInfo2Title
    System_Contact_JobInfo2YomiCompanyName
    System_Contact_JobInfo3CompanyAddress
    System_Contact_JobInfo3CompanyName
    System_Contact_JobInfo3Department
    System_Contact_JobInfo3Manager
    System_Contact_JobInfo3OfficeLocation
    System_Contact_JobInfo3Title
    System_Contact_JobInfo3YomiCompanyName
    System_Contact_JobTitle
    System_Contact_Label
    System_Contact_LastName
    System_Contact_MailingAddress
    System_Contact_MiddleName
    System_Contact_MobileTelephone
    System_Contact_NickName
    System_Contact_OfficeLocation
    System_Contact_OtherAddress
    System_Contact_OtherAddress1Country
    System_Contact_OtherAddress1Locality
    System_Contact_OtherAddress1PostalCode
    System_Contact_OtherAddress1Region
    System_Contact_OtherAddress1Street
    System_Contact_OtherAddress2Country
    System_Contact_OtherAddress2Locality
    System_Contact_OtherAddress2PostalCode
    System_Contact_OtherAddress2Region
    System_Contact_OtherAddress2Street
    System_Contact_OtherAddress3Country
    System_Contact_OtherAddress3Locality
    System_Contact_OtherAddress3PostalCode
    System_Contact_OtherAddress3Region
    System_Contact_OtherAddress3Street
    System_Contact_OtherAddressCity
    System_Contact_OtherAddressCountry
    System_Contact_OtherAddressPostalCode
    System_Contact_OtherAddressPostOfficeBox
    System_Contact_OtherAddressState
    System_Contact_OtherAddressStreet
    System_Contact_OtherEmailAddresses
    System_Contact_PagerTelephone
    System_Contact_PersonalTitle
    System_Contact_PhoneNumbersCanonical
    System_Contact_Prefix
    System_Contact_PrimaryAddressCity
    System_Contact_PrimaryAddressCountry
    System_Contact_PrimaryAddressPostalCode
    System_Contact_PrimaryAddressPostOfficeBox
    System_Contact_PrimaryAddressState
    System_Contact_PrimaryAddressStreet
    System_Contact_PrimaryEmailAddress
    System_Contact_PrimaryTelephone
    System_Contact_Profession
    System_Contact_SpouseName
    System_Contact_Suffix
    System_Contact_TelexNumber
    System_Contact_TTYTDDTelephone
    System_Contact_WebPage
    System_Contact_Webpage2
    System_Contact_Webpage3
    ' Core properties
    System_AcquisitionID
    System_ApplicationDefinedProperties
    System_ApplicationName
    System_AppZoneIdentifier
    System_Author
    System_CachedFileUpdaterContentIdForConflictResolution
    System_CachedFileUpdaterContentIdForStream
    System_Capacity
    System_Category
    System_Comment
    System_Company
    System_ComputerName
    System_ContainedItems
    System_ContentStatus
    System_ContentType
    System_Copyright
    System_CreatorAppId
    System_CreatorOpenWithUIOptions
    System_DataObjectFormat
    System_DateAccessed
    System_DateAcquired
    System_DateArchived
    System_DateCompleted
    System_DateCreated
    System_DateImported
    System_DateModified
    System_DefaultSaveLocationDisplay
    System_DueDate
    System_EndDate
    System_ExpandoProperties
    System_FileAllocationSize
    System_FileAttributes
    System_FileCount
    System_FileDescription
    System_FileExtension
    System_FileFRN
    System_FileName
    System_FileOfflineAvailabilityStatus
    System_FileOwner
    System_FilePlaceholderStatus
    System_FileVersion
    System_FindData
    System_FlagColor
    System_FlagColorText
    System_FlagStatus
    System_FlagStatusText
    System_FolderKind
    System_FolderNameDisplay
    System_FreeSpace
    System_FullText
    System_HighKeywords
    System_Identity
    System_Identity_Blob
    System_Identity_DisplayName
    System_Identity_InternetSid
    System_Identity_IsMeIdentity
    System_Identity_KeyProviderContext
    System_Identity_KeyProviderName
    System_Identity_LogonStatusString
    System_Identity_PrimaryEmailAddress
    System_Identity_PrimarySid
    System_Identity_ProviderData
    System_Identity_ProviderID
    System_Identity_QualifiedUserName
    System_Identity_UniqueID
    System_Identity_UserName
    System_IdentityProvider_Name
    System_IdentityProvider_Picture
    System_ImageParsingName
    System_Importance
    System_ImportanceText
    System_IsAttachment
    System_IsDefaultNonOwnerSaveLocation
    System_IsDefaultSaveLocation
    System_IsDeleted
    System_IsEncrypted
    System_IsFlagged
    System_IsFlaggedComplete
    System_IsIncomplete
    System_IsLocationSupported
    System_IsPinnedToNameSpaceTree
    System_IsRead
    System_IsSearchOnlyItem
    System_IsSendToTarget
    System_IsShared
    System_ItemAuthors
    System_ItemClassType
    System_ItemDate
    System_ItemFolderNameDisplay
    System_ItemFolderPathDisplay
    System_ItemFolderPathDisplayNarrow
    System_ItemName
    System_ItemNameDisplay
    System_ItemNameDisplayWithoutExtension
    System_ItemNamePrefix
    System_ItemNameSortOverride
    System_ItemParticipants
    System_ItemPathDisplay
    System_ItemPathDisplayNarrow
    System_ItemSubType
    System_ItemType
    System_ItemTypeText
    System_ItemUrl
    System_Keywords
    System_Kind
    System_KindText
    System_Language
    System_LastSyncError
    System_LastSyncWarning
    System_LastWriterPackageFamilyName
    System_LowKeywords
    System_MediumKeywords
    System_MileageInformation
    System_MIMEType
    System_Null
    System_OfflineAvailability
    System_OfflineStatus
    System_OriginalFileName
    System_OwnerSID
    System_ParentalRating
    System_ParentalRatingReason
    System_ParentalRatingsOrganization
    System_ParsingBindContext
    System_ParsingName
    System_ParsingPath
    System_PerceivedType
    System_PercentFull
    System_Priority
    System_PriorityText
    System_Project
    System_ProviderItemID
    System_Rating
    System_RatingText
    System_RemoteConflictingFile
    System_Security_AllowedEnterpriseDataProtectionIdentities
    System_Security_EncryptionOwners
    System_Security_EncryptionOwnersDisplay
    System_Sensitivity
    System_SensitivityText
    System_SFGAOFlags
    System_SharedWith
    System_ShareUserRating
    System_SharingStatus
    System_Shell_OmitFromView
    System_SimpleRating
    System_Size
    System_SoftwareUsed
    System_SourceItem
    System_SourcePackageFamilyName
    System_StartDate
    System_Status
    System_StorageProviderCallerVersionInformation
    System_StorageProviderError
    System_StorageProviderFileChecksum
    System_StorageProviderFileFlags
    System_StorageProviderFileIdentifier
    System_StorageProviderFileRemoteUri
    System_StorageProviderFileVersion
    System_StorageProviderFileVersionWaterline
    System_StorageProviderId
    System_StorageProviderShareStatuses
    System_StorageProviderSharingStatus
    System_StorageProviderStatus
    System_Subject
    System_SyncTransferStatus
    System_Thumbnail
    System_ThumbnailCacheId
    System_ThumbnailStream
    System_Title
    System_TitleSortOverride
    System_TotalFileSize
    System_Trademarks
    System_TransferOrder
    System_TransferPosition
    System_TransferSize
    System_VolumeId
    System_ZoneIdentifier
    ' Devices properties
    System_Device_PrinterURL
    System_DeviceInterface_Bluetooth_DeviceAddress
    System_DeviceInterface_Bluetooth_Flags
    System_DeviceInterface_Bluetooth_LastConnectedTime
    System_DeviceInterface_Bluetooth_Manufacturer
    System_DeviceInterface_Bluetooth_ModelNumber
    System_DeviceInterface_Bluetooth_ProductId
    System_DeviceInterface_Bluetooth_ProductVersion
    System_DeviceInterface_Bluetooth_ServiceGuid
    System_DeviceInterface_Bluetooth_VendorId
    System_DeviceInterface_Bluetooth_VendorIdSource
    System_DeviceInterface_Hid_IsReadOnly
    System_DeviceInterface_Hid_ProductId
    System_DeviceInterface_Hid_UsageId
    System_DeviceInterface_Hid_UsagePage
    System_DeviceInterface_Hid_VendorId
    System_DeviceInterface_Hid_VersionNumber
    System_DeviceInterface_PrinterDriverDirectory
    System_DeviceInterface_PrinterDriverName
    System_DeviceInterface_PrinterEnumerationFlag
    System_DeviceInterface_PrinterName
    System_DeviceInterface_PrinterPortName
    System_DeviceInterface_Proximity_SupportsNfc
    System_DeviceInterface_Serial_PortName
    System_DeviceInterface_Serial_UsbProductId
    System_DeviceInterface_Serial_UsbVendorId
    System_DeviceInterface_WinUsb_DeviceInterfaceClasses
    System_DeviceInterface_WinUsb_UsbClass
    System_DeviceInterface_WinUsb_UsbProductId
    System_DeviceInterface_WinUsb_UsbProtocol
    System_DeviceInterface_WinUsb_UsbSubClass
    System_DeviceInterface_WinUsb_UsbVendorId
    System_Devices_Aep_AepId
    System_Devices_Aep_Bluetooth_Cod_Major
    System_Devices_Aep_Bluetooth_Cod_Minor
    System_Devices_Aep_Bluetooth_Cod_Services_Audio
    System_Devices_Aep_Bluetooth_Cod_Services_Capturing
    System_Devices_Aep_Bluetooth_Cod_Services_Information
    System_Devices_Aep_Bluetooth_Cod_Services_LimitedDiscovery
    System_Devices_Aep_Bluetooth_Cod_Services_Networking
    System_Devices_Aep_Bluetooth_Cod_Services_ObjectXfer
    System_Devices_Aep_Bluetooth_Cod_Services_Positioning
    System_Devices_Aep_Bluetooth_Cod_Services_Rendering
    System_Devices_Aep_Bluetooth_Cod_Services_Telephony
    System_Devices_Aep_Bluetooth_LastSeenTime
    System_Devices_Aep_Bluetooth_Le_AddressType
    System_Devices_Aep_Bluetooth_Le_Appearance
    System_Devices_Aep_Bluetooth_Le_Appearance_Category
    System_Devices_Aep_Bluetooth_Le_Appearance_Subcategory
    System_Devices_Aep_Bluetooth_Le_IsConnectable
    System_Devices_Aep_CanPair
    System_Devices_Aep_Category
    System_Devices_Aep_ContainerId
    System_Devices_Aep_DeviceAddress
    System_Devices_Aep_IsConnected
    System_Devices_Aep_IsPaired
    System_Devices_Aep_IsPresent
    System_Devices_Aep_Manufacturer
    System_Devices_Aep_ModelId
    System_Devices_Aep_ModelName
    System_Devices_Aep_PointOfService_ConnectionTypes
    System_Devices_Aep_ProtocolId
    System_Devices_Aep_SignalStrength
    System_Devices_AepContainer_CanPair
    System_Devices_AepContainer_Categories
    System_Devices_AepContainer_Children
    System_Devices_AepContainer_ContainerId
    System_Devices_AepContainer_DialProtocol_InstalledApplications
    System_Devices_AepContainer_IsPaired
    System_Devices_AepContainer_IsPresent
    System_Devices_AepContainer_Manufacturer
    System_Devices_AepContainer_ModelIds
    System_Devices_AepContainer_ModelName
    System_Devices_AepContainer_ProtocolIds
    System_Devices_AepContainer_SupportedUriSchemes
    System_Devices_AepContainer_SupportsAudio
    System_Devices_AepContainer_SupportsCapturing
    System_Devices_AepContainer_SupportsImages
    System_Devices_AepContainer_SupportsInformation
    System_Devices_AepContainer_SupportsLimitedDiscovery
    System_Devices_AepContainer_SupportsNetworking
    System_Devices_AepContainer_SupportsObjectTransfer
    System_Devices_AepContainer_SupportsPositioning
    System_Devices_AepContainer_SupportsRendering
    System_Devices_AepContainer_SupportsTelephony
    System_Devices_AepContainer_SupportsVideo
    System_Devices_AepService_AepId
    System_Devices_AepService_Bluetooth_CacheMode
    System_Devices_AepService_Bluetooth_ServiceGuid
    System_Devices_AepService_Bluetooth_TargetDevice
    System_Devices_AepService_ContainerId
    System_Devices_AepService_FriendlyName
    System_Devices_AepService_IoT_ServiceInterfaces
    System_Devices_AepService_ParentAepIsPaired
    System_Devices_AepService_ProtocolId
    System_Devices_AepService_ServiceClassId
    System_Devices_AepService_ServiceId
    System_Devices_AppPackageFamilyName
    System_Devices_AudioDevice_Microphone_SensitivityInDbfs
    System_Devices_AudioDevice_Microphone_SignalToNoiseRatioInDb
    System_Devices_AudioDevice_RawProcessingSupported
    System_Devices_AudioDevice_SpeechProcessingSupported
    System_Devices_BatteryLife
    System_Devices_BatteryPlusCharging
    System_Devices_BatteryPlusChargingText
    System_Devices_Category
    System_Devices_CategoryGroup
    System_Devices_CategoryIds
    System_Devices_CategoryPlural
    System_Devices_ChargingState
    System_Devices_Children
    System_Devices_ClassGuid
    System_Devices_CompatibleIds
    System_Devices_Connected
    System_Devices_ContainerId
    System_Devices_DefaultTooltip
    System_Devices_DeviceCapabilities
    System_Devices_DeviceCharacteristics
    System_Devices_DeviceDescription1
    System_Devices_DeviceDescription2
    System_Devices_DeviceHasProblem
    System_Devices_DeviceInstanceId
    System_Devices_DeviceManufacturer
    System_Devices_DevObjectType
    System_Devices_DialProtocol_InstalledApplications
    System_Devices_DiscoveryMethod
    System_Devices_Dnssd_Domain
    System_Devices_Dnssd_FullName
    System_Devices_Dnssd_HostName
    System_Devices_Dnssd_InstanceName
    System_Devices_Dnssd_NetworkAdapterId
    System_Devices_Dnssd_PortNumber
    System_Devices_Dnssd_Priority
    System_Devices_Dnssd_ServiceName
    System_Devices_Dnssd_TextAttributes
    System_Devices_Dnssd_Ttl
    System_Devices_Dnssd_Weight
    System_Devices_FriendlyName
    System_Devices_FunctionPaths
    System_Devices_GlyphIcon
    System_Devices_HardwareIds
    System_Devices_Icon
    System_Devices_InLocalMachineContainer
    System_Devices_InterfaceClassGuid
    System_Devices_InterfaceEnabled
    System_Devices_InterfacePaths
    System_Devices_IpAddress
    System_Devices_IsDefault
    System_Devices_IsNetworkConnected
    System_Devices_IsShared
    System_Devices_IsSoftwareInstalling
    System_Devices_LaunchDeviceStageFromExplorer
    System_Devices_LocalMachine
    System_Devices_LocationPaths
    System_Devices_Manufacturer
    System_Devices_MetadataPath
    System_Devices_MicrophoneArray_Geometry
    System_Devices_MissedCalls
    System_Devices_ModelId
    System_Devices_ModelName
    System_Devices_ModelNumber
    System_Devices_NetworkedTooltip
    System_Devices_NetworkName
    System_Devices_NetworkType
    System_Devices_NewPictures
    System_Devices_Notification
    System_Devices_Notifications_LowBattery
    System_Devices_Notifications_MissedCall
    System_Devices_Notifications_NewMessage
    System_Devices_Notifications_NewVoicemail
    System_Devices_Notifications_StorageFull
    System_Devices_Notifications_StorageFullLinkText
    System_Devices_NotificationStore
    System_Devices_NotWorkingProperly
    System_Devices_Paired
    System_Devices_Parent
    System_Devices_PhysicalDeviceLocation
    System_Devices_PlaybackPositionPercent
    System_Devices_PlaybackState
    System_Devices_PlaybackTitle
    System_Devices_Present
    System_Devices_PresentationUrl
    System_Devices_PrimaryCategory
    System_Devices_RemainingDuration
    System_Devices_RestrictedInterface
    System_Devices_Roaming
    System_Devices_SafeRemovalRequired
    System_Devices_SchematicName
    System_Devices_ServiceAddress
    System_Devices_ServiceId
    System_Devices_SharedTooltip
    System_Devices_SignalStrength
    System_Devices_SmartCards_ReaderKind
    System_Devices_Status
    System_Devices_Status1
    System_Devices_Status2
    System_Devices_StorageCapacity
    System_Devices_StorageFreeSpace
    System_Devices_StorageFreeSpacePercent
    System_Devices_TextMessages
    System_Devices_Voicemail
    System_Devices_WiaDeviceType
    System_Devices_WiFi_InterfaceGuid
    System_Devices_WiFiDirect_DeviceAddress
    System_Devices_WiFiDirect_GroupId
    System_Devices_WiFiDirect_InformationElements
    System_Devices_WiFiDirect_InterfaceAddress
    System_Devices_WiFiDirect_InterfaceGuid
    System_Devices_WiFiDirect_IsConnected
    System_Devices_WiFiDirect_IsLegacyDevice
    System_Devices_WiFiDirect_IsMiracastLcpSupported
    System_Devices_WiFiDirect_IsVisible
    System_Devices_WiFiDirect_MiracastVersion
    System_Devices_WiFiDirect_Services
    System_Devices_WiFiDirect_SupportedChannelList
    System_Devices_WiFiDirectServices_AdvertisementId
    System_Devices_WiFiDirectServices_RequestServiceInformation
    System_Devices_WiFiDirectServices_ServiceAddress
    System_Devices_WiFiDirectServices_ServiceConfigMethods
    System_Devices_WiFiDirectServices_ServiceInformation
    System_Devices_WiFiDirectServices_ServiceName
    System_Devices_WinPhone8CameraFlags
    System_Devices_Wwan_InterfaceGuid
    System_Storage_Portable
    System_Storage_RemovableMedia
    System_Storage_SystemCritical
    ' Document properties
    System_Document_ByteCount
    System_Document_CharacterCount
    System_Document_ClientID
    System_Document_Contributor
    System_Document_DateCreated
    System_Document_DatePrinted
    System_Document_DateSaved
    System_Document_Division
    System_Document_DocumentID
    System_Document_HiddenSlideCount
    System_Document_LastAuthor
    System_Document_LineCount
    System_Document_Manager
    System_Document_MultimediaClipCount
    System_Document_NoteCount
    System_Document_PageCount
    System_Document_ParagraphCount
    System_Document_PresentationFormat
    System_Document_RevisionNumber
    System_Document_Security
    System_Document_SlideCount
    System_Document_Template
    System_Document_TotalEditingTime
    System_Document_Version
    System_Document_WordCount
    ' DRM properties
    System_DRM_DatePlayExpires
    System_DRM_DatePlayStarts
    System_DRM_Description
    System_DRM_IsDisabled
    System_DRM_IsProtected
    System_DRM_PlayCount
    ' GPS properties
    System_GPS_Altitude
    System_GPS_AltitudeDenominator
    System_GPS_AltitudeNumerator
    System_GPS_AltitudeRef
    System_GPS_AreaInformation
    System_GPS_Date
    System_GPS_DestBearing
    System_GPS_DestBearingDenominator
    System_GPS_DestBearingNumerator
    System_GPS_DestBearingRef
    System_GPS_DestDistance
    System_GPS_DestDistanceDenominator
    System_GPS_DestDistanceNumerator
    System_GPS_DestDistanceRef
    System_GPS_DestLatitude
    System_GPS_DestLatitudeDenominator
    System_GPS_DestLatitudeNumerator
    System_GPS_DestLatitudeRef
    System_GPS_DestLongitude
    System_GPS_DestLongitudeDenominator
    System_GPS_DestLongitudeNumerator
    System_GPS_DestLongitudeRef
    System_GPS_Differential
    System_GPS_DOP
    System_GPS_DOPDenominator
    System_GPS_DOPNumerator
    System_GPS_ImgDirection
    System_GPS_ImgDirectionDenominator
    System_GPS_ImgDirectionNumerator
    System_GPS_ImgDirectionRef
    System_GPS_Latitude
    System_GPS_LatitudeDecimal
    System_GPS_LatitudeDenominator
    System_GPS_LatitudeNumerator
    System_GPS_LatitudeRef
    System_GPS_Longitude
    System_GPS_LongitudeDecimal
    System_GPS_LongitudeDenominator
    System_GPS_LongitudeNumerator
    System_GPS_LongitudeRef
    System_GPS_MapDatum
    System_GPS_MeasureMode
    System_GPS_ProcessingMethod
    System_GPS_Satellites
    System_GPS_Speed
    System_GPS_SpeedDenominator
    System_GPS_SpeedNumerator
    System_GPS_SpeedRef
    System_GPS_Status
    System_GPS_Track
    System_GPS_TrackDenominator
    System_GPS_TrackNumerator
    System_GPS_TrackRef
    System_GPS_VersionID
    ' History properties
    System_History_VisitCount
    ' Image properties
    System_Image_BitDepth
    System_Image_ColorSpace
    System_Image_CompressedBitsPerPixel
    System_Image_CompressedBitsPerPixelDenominator
    System_Image_CompressedBitsPerPixelNumerator
    System_Image_Compression
    System_Image_CompressionText
    System_Image_Dimensions
    System_Image_HorizontalResolution
    System_Image_HorizontalSize
    System_Image_ImageID
    System_Image_ResolutionUnit
    System_Image_VerticalResolution
    System_Image_VerticalSize
    ' Journal properties
    System_Journal_Contacts
    System_Journal_EntryType
    ' LayoutPattern properties
    System_LayoutPattern_ContentViewModeForBrowse
    System_LayoutPattern_ContentViewModeForSearch
    ' Link properties
    System_History_SelectionCount
    System_History_TargetUrlHostName
    System_Link_Arguments
    System_Link_Comment
    System_Link_DateVisited
    System_Link_Description
    System_Link_FeedItemLocalId
    System_Link_Status
    System_Link_TargetExtension
    System_Link_TargetParsingPath
    System_Link_TargetSFGAOFlags
    System_Link_TargetUrlHostName
    System_Link_TargetUrlPath
    ' Media properties
    System_Media_AuthorUrl
    System_Media_AverageLevel
    System_Media_ClassPrimaryID
    System_Media_ClassSecondaryID
    System_Media_CollectionGroupID
    System_Media_CollectionID
    System_Media_ContentDistributor
    System_Media_ContentID
    System_Media_CreatorApplication
    System_Media_CreatorApplicationVersion
    System_Media_DateEncoded
    System_Media_DateReleased
    System_Media_DlnaProfileID
    System_Media_Duration
    System_Media_DVDID
    System_Media_EncodedBy
    System_Media_EncodingSettings
    System_Media_EpisodeNumber
    System_Media_FrameCount
    System_Media_MCDI
    System_Media_MetadataContentProvider
    System_Media_Producer
    System_Media_PromotionUrl
    System_Media_ProtectionType
    System_Media_ProviderRating
    System_Media_ProviderStyle
    System_Media_Publisher
    System_Media_SeasonNumber
    System_Media_SeriesName
    System_Media_SubscriptionContentId
    System_Media_SubTitle
    System_Media_ThumbnailLargePath
    System_Media_ThumbnailLargeUri
    System_Media_ThumbnailSmallPath
    System_Media_ThumbnailSmallUri
    System_Media_UniqueFileIdentifier
    System_Media_UserNoAutoInfo
    System_Media_UserWebUrl
    System_Media_Writer
    System_Media_Year
    ' Message properties
    System_Message_AttachmentContents
    System_Message_AttachmentNames
    System_Message_BccAddress
    System_Message_BccName
    System_Message_CcAddress
    System_Message_CcName
    System_Message_ConversationID
    System_Message_ConversationIndex
    System_Message_DateReceived
    System_Message_DateSent
    System_Message_Flags
    System_Message_FromAddress
    System_Message_FromName
    System_Message_HasAttachments
    System_Message_IsFwdOrReply
    System_Message_MessageClass
    System_Message_Participants
    System_Message_ProofInProgress
    System_Message_SenderAddress
    System_Message_SenderName
    System_Message_Store
    System_Message_ToAddress
    System_Message_ToDoFlags
    System_Message_ToDoTitle
    System_Message_ToName
    ' Music properties
    System_Music_AlbumArtist
    System_Music_AlbumArtistSortOverride
    System_Music_AlbumID
    System_Music_AlbumTitle
    System_Music_AlbumTitleSortOverride
    System_Music_Artist
    System_Music_ArtistSortOverride
    System_Music_BeatsPerMinute
    System_Music_Composer
    System_Music_ComposerSortOverride
    System_Music_Conductor
    System_Music_ContentGroupDescription
    System_Music_DiscNumber
    System_Music_DisplayArtist
    System_Music_Genre
    System_Music_InitialKey
    System_Music_IsCompilation
    System_Music_Lyrics
    System_Music_Mood
    System_Music_PartOfSet
    System_Music_Period
    System_Music_SynchronizedLyrics
    System_Music_TrackNumber
    ' Note properties
    System_Note_Color
    System_Note_ColorText
    ' Photo properties
    System_Photo_Aperture
    System_Photo_ApertureDenominator
    System_Photo_ApertureNumerator
    System_Photo_Brightness
    System_Photo_BrightnessDenominator
    System_Photo_BrightnessNumerator
    System_Photo_CameraManufacturer
    System_Photo_CameraModel
    System_Photo_CameraSerialNumber
    System_Photo_Contrast
    System_Photo_ContrastText
    System_Photo_DateTaken
    System_Photo_DigitalZoom
    System_Photo_DigitalZoomDenominator
    System_Photo_DigitalZoomNumerator
    System_Photo_Event
    System_Photo_EXIFVersion
    System_Photo_ExposureBias
    System_Photo_ExposureBiasDenominator
    System_Photo_ExposureBiasNumerator
    System_Photo_ExposureIndex
    System_Photo_ExposureIndexDenominator
    System_Photo_ExposureIndexNumerator
    System_Photo_ExposureProgram
    System_Photo_ExposureProgramText
    System_Photo_ExposureTime
    System_Photo_ExposureTimeDenominator
    System_Photo_ExposureTimeNumerator
    System_Photo_Flash
    System_Photo_FlashEnergy
    System_Photo_FlashEnergyDenominator
    System_Photo_FlashEnergyNumerator
    System_Photo_FlashManufacturer
    System_Photo_FlashModel
    System_Photo_FlashText
    System_Photo_FNumber
    System_Photo_FNumberDenominator
    System_Photo_FNumberNumerator
    System_Photo_FocalLength
    System_Photo_FocalLengthDenominator
    System_Photo_FocalLengthInFilm
    System_Photo_FocalLengthNumerator
    System_Photo_FocalPlaneXResolution
    System_Photo_FocalPlaneXResolutionDenominator
    System_Photo_FocalPlaneXResolutionNumerator
    System_Photo_FocalPlaneYResolution
    System_Photo_FocalPlaneYResolutionDenominator
    System_Photo_FocalPlaneYResolutionNumerator
    System_Photo_GainControl
    System_Photo_GainControlDenominator
    System_Photo_GainControlNumerator
    System_Photo_GainControlText
    System_Photo_ISOSpeed
    System_Photo_LensManufacturer
    System_Photo_LensModel
    System_Photo_LightSource
    System_Photo_MakerNote
    System_Photo_MakerNoteOffset
    System_Photo_MaxAperture
    System_Photo_MaxApertureDenominator
    System_Photo_MaxApertureNumerator
    System_Photo_MeteringMode
    System_Photo_MeteringModeText
    System_Photo_Orientation
    System_Photo_OrientationText
    System_Photo_PeopleNames
    System_Photo_PhotometricInterpretation
    System_Photo_PhotometricInterpretationText
    System_Photo_ProgramMode
    System_Photo_ProgramModeText
    System_Photo_RelatedSoundFile
    System_Photo_Saturation
    System_Photo_SaturationText
    System_Photo_Sharpness
    System_Photo_SharpnessText
    System_Photo_ShutterSpeed
    System_Photo_ShutterSpeedDenominator
    System_Photo_ShutterSpeedNumerator
    System_Photo_SubjectDistance
    System_Photo_SubjectDistanceDenominator
    System_Photo_SubjectDistanceNumerator
    System_Photo_TagViewAggregate
    System_Photo_TranscodedForSync
    System_Photo_WhiteBalance
    System_Photo_WhiteBalanceText
    ' PropGroup properties
    System_PropGroup_Advanced
    System_PropGroup_Audio
    System_PropGroup_Calendar
    System_PropGroup_Camera
    System_PropGroup_Contact
    System_PropGroup_Content
    System_PropGroup_Description
    System_PropGroup_FileSystem
    System_PropGroup_General
    System_PropGroup_GPS
    System_PropGroup_Image
    System_PropGroup_Media
    System_PropGroup_MediaAdvanced
    System_PropGroup_Message
    System_PropGroup_Music
    System_PropGroup_Origin
    System_PropGroup_PhotoAdvanced
    System_PropGroup_RecordedTV
    System_PropGroup_Video
    ' PropList properties
    System_InfoTipText
    System_PropList_ConflictPrompt
    System_PropList_ContentViewModeForBrowse
    System_PropList_ContentViewModeForSearch
    System_PropList_ExtendedTileInfo
    System_PropList_FileOperationPrompt
    System_PropList_FullDetails
    System_PropList_InfoTip
    System_PropList_NonPersonal
    System_PropList_PreviewDetails
    System_PropList_PreviewTitle
    System_PropList_QuickTip
    System_PropList_TileInfo
    System_PropList_XPDetailsPanel
    ' RecordedTV properties
    System_RecordedTV_ChannelNumber
    System_RecordedTV_Credits
    System_RecordedTV_DateContentExpires
    System_RecordedTV_EpisodeName
    System_RecordedTV_IsATSCContent
    System_RecordedTV_IsClosedCaptioningAvailable
    System_RecordedTV_IsDTVContent
    System_RecordedTV_IsHDContent
    System_RecordedTV_IsRepeatBroadcast
    System_RecordedTV_IsSAP
    System_RecordedTV_NetworkAffiliation
    System_RecordedTV_OriginalBroadcastDate
    System_RecordedTV_ProgramDescription
    System_RecordedTV_RecordingTime
    System_RecordedTV_StationCallSign
    System_RecordedTV_StationName
    ' Search properties
    System_Search_AutoSummary
    System_Search_ContainerHash
    System_Search_Contents
    System_Search_EntryID
    System_Search_ExtendedProperties
    System_Search_GatherTime
    System_Search_HitCount
    System_Search_IsClosedDirectory
    System_Search_IsFullyContained
    System_Search_QueryFocusedSummary
    System_Search_QueryFocusedSummaryWithFallback
    System_Search_QueryPropertyHits
    System_Search_Rank
    System_Search_Store
    System_Search_UrlToIndex
    System_Search_UrlToIndexWithModificationTime
    System_Supplemental_AlbumID
    System_Supplemental_ResourceId
    ' Shell properties
    System_DescriptionID
    System_InternalName
    System_LibraryLocationsCount
    System_Link_TargetSFGAOFlagsStrings
    System_Link_TargetUrl
    System_NamespaceCLSID
    System_Shell_SFGAOFlagsStrings
    System_StatusBarSelectedItemCount
    System_StatusBarViewItemCount
    ' Software properties
    System_AppUserModel_ExcludeFromShowInNewInstall
    System_AppUserModel_ID
    System_AppUserModel_IsDestListSeparator
    System_AppUserModel_IsDualMode
    System_AppUserModel_PreventPinning
    System_AppUserModel_RelaunchCommand
    System_AppUserModel_RelaunchDisplayNameResource
    System_AppUserModel_RelaunchIconResource
    System_AppUserModel_StartPinOption
    System_AppUserModel_ToastActivatorCLSID
    System_EdgeGesture_DisableTouchWhenFullscreen
    System_Software_DateLastUsed
    System_Software_ProductName
    ' Sync properties
    System_Sync_Comments
    System_Sync_ConflictDescription
    System_Sync_ConflictFirstLocation
    System_Sync_ConflictSecondLocation
    System_Sync_HandlerCollectionID
    System_Sync_HandlerID
    System_Sync_HandlerName
    System_Sync_HandlerType
    System_Sync_HandlerTypeLabel
    System_Sync_ItemID
    System_Sync_ItemName
    System_Sync_ProgressPercentage
    System_Sync_State
    System_Sync_Status
    ' Task properties
    System_Task_BillingInformation
    System_Task_CompletionStatus
    System_Task_Owner
    ' Video properties
    System_Video_Compression
    System_Video_Director
    System_Video_EncodingBitrate
    System_Video_FourCC
    System_Video_FrameHeight
    System_Video_FrameRate
    System_Video_FrameWidth
    System_Video_HorizontalAspectRatio
    System_Video_IsSpherical
    System_Video_IsStereo
    System_Video_Orientation
    System_Video_SampleSize
    System_Video_StreamName
    System_Video_StreamNumber
    System_Video_TotalBitrate
    System_Video_TranscodedForSync
    System_Video_VerticalAspectRatio
    ' Volume properties
    System_Volume_FileSystem
    System_Volume_IsMappedDrive
    System_Volume_IsRoot
    EPropertyKeys_Max
End Enum

' ----==== Type ====----
Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

Private Type PKEY 'PROPERTYKEY
    fmtid As GUID
    pid   As Long
End Type

Private Type NamedPKEY
    Name As String
    PKEY As PKEY
End Type

Private Declare Function CLSIDFromString Lib "ole32" (ByVal pString As Long, ByRef pCLSID As GUID) As Long
Private Declare Function StringFromCLSID Lib "ole32" (ByRef pCLSID As GUID, ByVal pString As Long) As Long

Private m_PKeys() As NamedPKEY 'PROPERTYKEY

Private Function NamedPKEY(ByVal aName As String, ByVal sGuid As String, ByVal aPID As Long) As NamedPKEY
    With NamedPKEY
        .Name = aName
        Dim hr As Long
        With .PKEY
            hr = CLSIDFromString(StrPtr(sGuid), .fmtid)
            .pid = aPID
        End With
    End With
End Function

Public Function GetNamedPKEY(e As EPropertyKeys) As NamedPKEY
    GetNamedPKEY = m_PKeys(e)
End Function

Private Function NamedPKEY_ToStr(this As NamedPKEY) As String
    Dim s As String: s = Space(80)
    With this
        With .PKEY
            Dim hr As Long
            hr = StringFromCLSID(.fmtid, StrPtr(s))
            s = Trim$(s)
            s = s & "-" & .pid
        End With
        s = .Name & " " & s
    End With
    NamedPKEY_ToStr = s
End Function

Public Function PropertyKey_ToStr(e As EPropertyKeys) As String
    PropertyKey_ToStr = NamedPKEY_ToStr(GetNamedPKEY(EPropertyKeys_Max))
End Function

Public Sub Init()
    ReDim m_PKeys(0 To EPropertyKeys.EPropertyKeys_Max - 1)
    Dim i As Long
    m_PKeys(i) = NamedPKEY("System.Audio.ChannelCount", "{64440490-4C8B-11D1-8B70-080036B11A03}", 7): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Audio.Compression", "{64440490-4C8B-11D1-8B70-080036B11A03}", 10): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Audio.EncodingBitrate", "{64440490-4C8B-11D1-8B70-080036B11A03}", 4): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Audio.Format", "{64440490-4C8B-11D1-8B70-080036B11A03}", 2): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Audio.IsVariableBitRate", "{E6822FEE-8C17-4D62-823C-8E9CFCBD1D5C}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Audio.PeakValue", "{2579E5D0-1116-4084-BD9A-9B4F7CB4DF5E}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Audio.SampleRate", "{64440490-4C8B-11D1-8B70-080036B11A03}", 5): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Audio.SampleSize", "{64440490-4C8B-11D1-8B70-080036B11A03}", 6): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Audio.StreamName", "{64440490-4C8B-11D1-8B70-080036B11A03}", 9): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Audio.StreamNumber", "{64440490-4C8B-11D1-8B70-080036B11A03}", 8): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Calendar.Duration", "{293CA35A-09AA-4DD2-B180-1FE245728A52}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Calendar.IsOnline", "{BFEE9149-E3E2-49A7-A862-C05988145CEC}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Calendar.IsRecurring", "{315B9C8D-80A9-4EF9-AE16-8E746DA51D70}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Calendar.Location", "{F6272D18-CECC-40B1-B26A-3911717AA7BD}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Calendar.OptionalAttendeeAddresses", "{D55BAE5A-3892-417A-A649-C6AC5AAAEAB3}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Calendar.OptionalAttendeeNames", "{09429607-582D-437F-84C3-DE93A2B24C3C}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Calendar.OrganizerAddress", "{744C8242-4DF5-456C-AB9E-014EFB9021E3}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Calendar.OrganizerName", "{AAA660F9-9865-458E-B484-01BC7FE3973E}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Calendar.ReminderTime", "{72FC5BA4-24F9-4011-9F3F-ADD27AFAD818}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Calendar.RequiredAttendeeAddresses", "{0BA7D6C3-568D-4159-AB91-781A91FB71E5}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Calendar.RequiredAttendeeNames", "{B33AF30B-F552-4584-936C-CB93E5CDA29F}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Calendar.Resources", "{00F58A38-C54B-4C40-8696-97235980EAE1}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Calendar.ResponseStatus", "{188C1F91-3C40-4132-9EC5-D8B03B72A8A2}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Calendar.ShowTimeAs", "{5BF396D4-5EB2-466F-BDE9-2FB3F2361D6E}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Calendar.ShowTimeAsText", "{53DA57CF-62C0-45C4-81DE-7610BCEFD7F5}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Communication.AccountName", "{E3E0584C-B788-4A5A-BB20-7F5A44C9ACDD}", 9): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Communication.DateItemExpires", "{428040AC-A177-4C8A-9760-F6F761227F9A}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Communication.Direction", "{8E531030-B960-4346-AE0D-66BC9A86FB94}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Communication.FollowupIconIndex", "{83A6347E-6FE4-4F40-BA9C-C4865240D1F4}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Communication.HeaderItem", "{C9C34F84-2241-4401-B607-BD20ED75AE7F}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Communication.PolicyTag", "{EC0B4191-AB0B-4C66-90B6-C6637CDEBBAB}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Communication.SecurityFlags", "{8619A4B6-9F4D-4429-8C0F-B996CA59E335}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Communication.Suffix", "{807B653A-9E91-43EF-8F97-11CE04EE20C5}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Communication.TaskStatus", "{BE1A72C6-9A1D-46B7-AFE7-AFAF8CEF4999}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Communication.TaskStatusText", "{A6744477-C237-475B-A075-54F34498292A}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Computer.DecoratedFreeSpace", "{9B174B35-40FF-11D2-A27E-00C04FC30871}", 7): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Contact.AccountPictureDynamicVideo", "{0B8BB018-2725-4B44-92BA-7933AEB2DDE7}", 2): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Contact.AccountPictureLarge", "{0B8BB018-2725-4B44-92BA-7933AEB2DDE7}", 3): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Contact.AccountPictureSmall", "{0B8BB018-2725-4B44-92BA-7933AEB2DDE7}", 4): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Contact.Anniversary", "{9AD5BADB-CEA7-4470-A03D-B84E51B9949E}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Contact.AssistantName", "{CD102C9C-5540-4A88-A6F6-64E4981C8CD1}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Contact.AssistantTelephone", "{9A93244D-A7AD-4FF8-9B99-45EE4CC09AF6}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Contact.Birthday", "{176DC63C-2688-4E89-8143-A347800F25E9}", 47): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Contact.BusinessAddress", "{730FB6DD-CF7C-426B-A03F-BD166CC9EE24}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Contact.BusinessAddress1Country", "{A7B6F596-D678-4BC1-B05F-0203D27E8AA1}", 119): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Contact.BusinessAddress1Locality", "{A7B6F596-D678-4BC1-B05F-0203D27E8AA1}", 117): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Contact.BusinessAddress1PostalCode", "{A7B6F596-D678-4BC1-B05F-0203D27E8AA1}", 120): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Contact.BusinessAddress1Region", "{A7B6F596-D678-4BC1-B05F-0203D27E8AA1}", 118): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Contact.BusinessAddress1Street", "{A7B6F596-D678-4BC1-B05F-0203D27E8AA1}", 116): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Contact.BusinessAddress2Country", "{A7B6F596-D678-4BC1-B05F-0203D27E8AA1}", 124): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Contact.BusinessAddress2Locality", "{A7B6F596-D678-4BC1-B05F-0203D27E8AA1}", 122): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Contact.BusinessAddress2PostalCode", "{A7B6F596-D678-4BC1-B05F-0203D27E8AA1}", 125): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Contact.BusinessAddress2Region", "{A7B6F596-D678-4BC1-B05F-0203D27E8AA1}", 123): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Contact.BusinessAddress2Street", "{A7B6F596-D678-4BC1-B05F-0203D27E8AA1}", 121): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Contact.BusinessAddress3Country", "{A7B6F596-D678-4BC1-B05F-0203D27E8AA1}", 129): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Contact.BusinessAddress3Locality", "{A7B6F596-D678-4BC1-B05F-0203D27E8AA1}", 127): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Contact.BusinessAddress3PostalCode", "{A7B6F596-D678-4BC1-B05F-0203D27E8AA1}", 130): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Contact.BusinessAddress3Region", "{A7B6F596-D678-4BC1-B05F-0203D27E8AA1}", 128): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Contact.BusinessAddress3Street", "{A7B6F596-D678-4BC1-B05F-0203D27E8AA1}", 126): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Contact.BusinessAddressCity", "{402B5934-EC5A-48C3-93E6-85E86A2D934E}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Contact.BusinessAddressCountry", "{B0B87314-FCF6-4FEB-8DFF-A50DA6AF561C}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Contact.BusinessAddressPostalCode", "{E1D4A09E-D758-4CD1-B6EC-34A8B5A73F80}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Contact.BusinessAddressPostOfficeBox", "{BC4E71CE-17F9-48D5-BEE9-021DF0EA5409}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Contact.BusinessAddressState", "{446F787F-10C4-41CB-A6C4-4D0343551597}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Contact.BusinessAddressStreet", "{DDD1460F-C0BF-4553-8CE4-10433C908FB0}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Contact.BusinessEmailAddresses", "{F271C659-7E5E-471F-BA25-7F77B286F836}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Contact.BusinessFaxNumber", "{91EFF6F3-2E27-42CA-933E-7C999FBE310B}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Contact.BusinessHomePage", "{56310920-2491-4919-99CE-EADB06FAFDB2}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Contact.BusinessTelephone", "{6A15E5A0-0A1E-4CD7-BB8C-D2F1B0C929BC}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Contact.CallbackTelephone", "{BF53D1C3-49E0-4F7F-8567-5A821D8AC542}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Contact.CarTelephone", "{8FDC6DEA-B929-412B-BA90-397A257465FE}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Contact.Children", "{D4729704-8EF1-43EF-9024-2BD381187FD5}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Contact.CompanyMainTelephone", "{8589E481-6040-473D-B171-7FA89C2708ED}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Contact.ConnectedServiceDisplayName", "{39B77F4F-A104-4863-B395-2DB2AD8F7BC1}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Contact.ConnectedServiceIdentities", "{80F41EB8-AFC4-4208-AA5F-CCE21A627281}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Contact.ConnectedServiceName", "{B5C84C9E-5927-46B5-A3CC-933C21B78469}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Contact.ConnectedServiceSupportedActions", "{A19FB7A9-024B-4371-A8BF-4D29C3E4E9C9}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Contact.DataSuppliers", "{9660C283-FC3A-4A08-A096-EED3AAC46DA2}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Contact.Department", "{FC9F7306-FF8F-4D49-9FB6-3FFE5C0951EC}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Contact.DisplayBusinessPhoneNumbers", "{364028DA-D895-41FE-A584-302B1BB70A76}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Contact.DisplayHomePhoneNumbers", "{5068BCDF-D697-4D85-8C53-1F1CDAB01763}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Contact.DisplayMobilePhoneNumbers", "{9CB0C358-9D7A-46B1-B466-DCC6F1A3D93D}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Contact.DisplayOtherPhoneNumbers", "{03089873-8EE8-4191-BD60-D31F72B7900B}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Contact.EmailAddress", "{F8FA7FA3-D12B-4785-8A4E-691A94F7A3E7}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Contact.EmailAddress2", "{38965063-EDC8-4268-8491-B7723172CF29}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Contact.EmailAddress3", "{644D37B4-E1B3-4BAD-B099-7E7C04966ACA}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Contact.EmailAddresses", "{84D8F337-981D-44B3-9615-C7596DBA17E3}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Contact.EmailName", "{CC6F4F24-6083-4BD4-8754-674D0DE87AB8}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Contact.FileAsName", "{F1A24AA7-9CA7-40F6-89EC-97DEF9FFE8DB}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Contact.FirstName", "{14977844-6B49-4AAD-A714-A4513BF60460}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Contact.FullName", "{635E9051-50A5-4BA2-B9DB-4ED056C77296}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Contact.Gender", "{3C8CEE58-D4F0-4CF9-B756-4E5D24447BCD}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Contact.GenderValue", "{3C8CEE58-D4F0-4CF9-B756-4E5D24447BCD}", 101): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Contact.Hobbies", "{5DC2253F-5E11-4ADF-9CFE-910DD01E3E70}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Contact.HomeAddress", "{98F98354-617A-46B8-8560-5B1B64BF1F89}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Contact.HomeAddress1Country", "{A7B6F596-D678-4BC1-B05F-0203D27E8AA1}", 104): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Contact.HomeAddress1Locality", "{A7B6F596-D678-4BC1-B05F-0203D27E8AA1}", 102): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Contact.HomeAddress1PostalCode", "{A7B6F596-D678-4BC1-B05F-0203D27E8AA1}", 105): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Contact.HomeAddress1Region", "{A7B6F596-D678-4BC1-B05F-0203D27E8AA1}", 103): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Contact.HomeAddress1Street", "{A7B6F596-D678-4BC1-B05F-0203D27E8AA1}", 101): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Contact.HomeAddress2Country", "{A7B6F596-D678-4BC1-B05F-0203D27E8AA1}", 109): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Contact.HomeAddress2Locality", "{A7B6F596-D678-4BC1-B05F-0203D27E8AA1}", 107): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Contact.HomeAddress2PostalCode", "{A7B6F596-D678-4BC1-B05F-0203D27E8AA1}", 110): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Contact.HomeAddress2Region", "{A7B6F596-D678-4BC1-B05F-0203D27E8AA1}", 108): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Contact.HomeAddress2Street", "{A7B6F596-D678-4BC1-B05F-0203D27E8AA1}", 106): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Contact.HomeAddress3Country", "{A7B6F596-D678-4BC1-B05F-0203D27E8AA1}", 114): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Contact.HomeAddress3Locality", "{A7B6F596-D678-4BC1-B05F-0203D27E8AA1}", 112): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Contact.HomeAddress3PostalCode", "{A7B6F596-D678-4BC1-B05F-0203D27E8AA1}", 115): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Contact.HomeAddress3Region", "{A7B6F596-D678-4BC1-B05F-0203D27E8AA1}", 113): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Contact.HomeAddress3Street", "{A7B6F596-D678-4BC1-B05F-0203D27E8AA1}", 111): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Contact.HomeAddressCity", "{176DC63C-2688-4E89-8143-A347800F25E9}", 65): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Contact.HomeAddressCountry", "{08A65AA1-F4C9-43DD-9DDF-A33D8E7EAD85}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Contact.HomeAddressPostalCode", "{8AFCC170-8A46-4B53-9EEE-90BAE7151E62}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Contact.HomeAddressPostOfficeBox", "{7B9F6399-0A3F-4B12-89BD-4ADC51C918AF}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Contact.HomeAddressState", "{C89A23D0-7D6D-4EB8-87D4-776A82D493E5}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Contact.HomeAddressStreet", "{0ADEF160-DB3F-4308-9A21-06237B16FA2A}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Contact.HomeEmailAddresses", "{56C90E9D-9D46-4963-886F-2E1CD9A694EF}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Contact.HomeFaxNumber", "{660E04D6-81AB-4977-A09F-82313113AB26}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Contact.HomeTelephone", "{176DC63C-2688-4E89-8143-A347800F25E9}", 20): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Contact.IMAddress", "{D68DBD8A-3374-4B81-9972-3EC30682DB3D}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Contact.Initials", "{F3D8F40D-50CB-44A2-9718-40CB9119495D}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Contact.JA.CompanyNamePhonetic", "{897B3694-FE9E-43E6-8066-260F590C0100}", 2): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Contact.JA.FirstNamePhonetic", "{897B3694-FE9E-43E6-8066-260F590C0100}", 3): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Contact.JA.LastNamePhonetic", "{897B3694-FE9E-43E6-8066-260F590C0100}", 4): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Contact.JobInfo1CompanyAddress", "{00F63DD8-22BD-4A5D-BA34-5CB0B9BDCB03}", 120): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Contact.JobInfo1CompanyName", "{00F63DD8-22BD-4A5D-BA34-5CB0B9BDCB03}", 102): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Contact.JobInfo1Department", "{00F63DD8-22BD-4A5D-BA34-5CB0B9BDCB03}", 106): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Contact.JobInfo1Manager", "{00F63DD8-22BD-4A5D-BA34-5CB0B9BDCB03}", 105): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Contact.JobInfo1OfficeLocation", "{00F63DD8-22BD-4A5D-BA34-5CB0B9BDCB03}", 104): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Contact.JobInfo1Title", "{00F63DD8-22BD-4A5D-BA34-5CB0B9BDCB03}", 103): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Contact.JobInfo1YomiCompanyName", "{00F63DD8-22BD-4A5D-BA34-5CB0B9BDCB03}", 101): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Contact.JobInfo2CompanyAddress", "{00F63DD8-22BD-4A5D-BA34-5CB0B9BDCB03}", 121): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Contact.JobInfo2CompanyName", "{00F63DD8-22BD-4A5D-BA34-5CB0B9BDCB03}", 108): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Contact.JobInfo2Department", "{00F63DD8-22BD-4A5D-BA34-5CB0B9BDCB03}", 113): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Contact.JobInfo2Manager", "{00F63DD8-22BD-4A5D-BA34-5CB0B9BDCB03}", 112): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Contact.JobInfo2OfficeLocation", "{00F63DD8-22BD-4A5D-BA34-5CB0B9BDCB03}", 110): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Contact.JobInfo2Title", "{00F63DD8-22BD-4A5D-BA34-5CB0B9BDCB03}", 109): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Contact.JobInfo2YomiCompanyName", "{00F63DD8-22BD-4A5D-BA34-5CB0B9BDCB03}", 107): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Contact.JobInfo3CompanyAddress", "{00F63DD8-22BD-4A5D-BA34-5CB0B9BDCB03}", 123): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Contact.JobInfo3CompanyName", "{00F63DD8-22BD-4A5D-BA34-5CB0B9BDCB03}", 115): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Contact.JobInfo3Department", "{00F63DD8-22BD-4A5D-BA34-5CB0B9BDCB03}", 119): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Contact.JobInfo3Manager", "{00F63DD8-22BD-4A5D-BA34-5CB0B9BDCB03}", 118): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Contact.JobInfo3OfficeLocation", "{00F63DD8-22BD-4A5D-BA34-5CB0B9BDCB03}", 117): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Contact.JobInfo3Title", "{00F63DD8-22BD-4A5D-BA34-5CB0B9BDCB03}", 116): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Contact.JobInfo3YomiCompanyName", "{00F63DD8-22BD-4A5D-BA34-5CB0B9BDCB03}", 114): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Contact.JobTitle", "{176DC63C-2688-4E89-8143-A347800F25E9}", 6): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Contact.Label", "{97B0AD89-DF49-49CC-834E-660974FD755B}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Contact.LastName", "{8F367200-C270-457C-B1D4-E07C5BCD90C7}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Contact.MailingAddress", "{C0AC206A-827E-4650-95AE-77E2BB74FCC9}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Contact.MiddleName", "{176DC63C-2688-4E89-8143-A347800F25E9}", 71): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Contact.MobileTelephone", "{176DC63C-2688-4E89-8143-A347800F25E9}", 35): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Contact.NickName", "{176DC63C-2688-4E89-8143-A347800F25E9}", 74): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Contact.OfficeLocation", "{176DC63C-2688-4E89-8143-A347800F25E9}", 7): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Contact.OtherAddress", "{508161FA-313B-43D5-83A1-C1ACCF68622C}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Contact.OtherAddress1Country", "{A7B6F596-D678-4BC1-B05F-0203D27E8AA1}", 134): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Contact.OtherAddress1Locality", "{A7B6F596-D678-4BC1-B05F-0203D27E8AA1}", 132): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Contact.OtherAddress1PostalCode", "{A7B6F596-D678-4BC1-B05F-0203D27E8AA1}", 135): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Contact.OtherAddress1Region", "{A7B6F596-D678-4BC1-B05F-0203D27E8AA1}", 133): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Contact.OtherAddress1Street", "{A7B6F596-D678-4BC1-B05F-0203D27E8AA1}", 131): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Contact.OtherAddress2Country", "{A7B6F596-D678-4BC1-B05F-0203D27E8AA1}", 139): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Contact.OtherAddress2Locality", "{A7B6F596-D678-4BC1-B05F-0203D27E8AA1}", 137): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Contact.OtherAddress2PostalCode", "{A7B6F596-D678-4BC1-B05F-0203D27E8AA1}", 140): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Contact.OtherAddress2Region", "{A7B6F596-D678-4BC1-B05F-0203D27E8AA1}", 138): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Contact.OtherAddress2Street", "{A7B6F596-D678-4BC1-B05F-0203D27E8AA1}", 136): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Contact.OtherAddress3Country", "{A7B6F596-D678-4BC1-B05F-0203D27E8AA1}", 144): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Contact.OtherAddress3Locality", "{A7B6F596-D678-4BC1-B05F-0203D27E8AA1}", 142): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Contact.OtherAddress3PostalCode", "{A7B6F596-D678-4BC1-B05F-0203D27E8AA1}", 145): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Contact.OtherAddress3Region", "{A7B6F596-D678-4BC1-B05F-0203D27E8AA1}", 143): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Contact.OtherAddress3Street", "{A7B6F596-D678-4BC1-B05F-0203D27E8AA1}", 141): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Contact.OtherAddressCity", "{6E682923-7F7B-4F0C-A337-CFCA296687BF}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Contact.OtherAddressCountry", "{8F167568-0AAE-4322-8ED9-6055B7B0E398}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Contact.OtherAddressPostalCode", "{95C656C1-2ABF-4148-9ED3-9EC602E3B7CD}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Contact.OtherAddressPostOfficeBox", "{8B26EA41-058F-43F6-AECC-4035681CE977}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Contact.OtherAddressState", "{71B377D6-E570-425F-A170-809FAE73E54E}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Contact.OtherAddressStreet", "{FF962609-B7D6-4999-862D-95180D529AEA}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Contact.OtherEmailAddresses", "{11D6336B-38C4-4EC9-84D6-EB38D0B150AF}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Contact.PagerTelephone", "{D6304E01-F8F5-4F45-8B15-D024A6296789}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Contact.PersonalTitle", "{176DC63C-2688-4E89-8143-A347800F25E9}", 69): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Contact.PhoneNumbersCanonical", "{D042D2A1-927E-40B5-A503-6EDBD42A517E}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Contact.Prefix", "{176DC63C-2688-4E89-8143-A347800F25E9}", 75): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Contact.PrimaryAddressCity", "{C8EA94F0-A9E3-4969-A94B-9C62A95324E0}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Contact.PrimaryAddressCountry", "{E53D799D-0F3F-466E-B2FF-74634A3CB7A4}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Contact.PrimaryAddressPostalCode", "{18BBD425-ECFD-46EF-B612-7B4A6034EDA0}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Contact.PrimaryAddressPostOfficeBox", "{DE5EF3C7-46E1-484E-9999-62C5308394C1}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Contact.PrimaryAddressState", "{F1176DFE-7138-4640-8B4C-AE375DC70A6D}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Contact.PrimaryAddressStreet", "{63C25B20-96BE-488F-8788-C09C407AD812}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Contact.PrimaryEmailAddress", "{176DC63C-2688-4E89-8143-A347800F25E9}", 48): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Contact.PrimaryTelephone", "{176DC63C-2688-4E89-8143-A347800F25E9}", 25): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Contact.Profession", "{7268AF55-1CE4-4F6E-A41F-B6E4EF10E4A9}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Contact.SpouseName", "{9D2408B6-3167-422B-82B0-F583B7A7CFE3}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Contact.Suffix", "{176DC63C-2688-4E89-8143-A347800F25E9}", 73): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Contact.TelexNumber", "{C554493C-C1F7-40C1-A76C-EF8C0614003E}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Contact.TTYTDDTelephone", "{AAF16BAC-2B55-45E6-9F6D-415EB94910DF}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Contact.WebPage", "{E3E0584C-B788-4A5A-BB20-7F5A44C9ACDD}", 18): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Contact.Webpage2", "{00F63DD8-22BD-4A5D-BA34-5CB0B9BDCB03}", 124): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Contact.Webpage3", "{00F63DD8-22BD-4A5D-BA34-5CB0B9BDCB03}", 125): i = i + 1
    m_PKeys(i) = NamedPKEY("System.AcquisitionID", "{65A98875-3C80-40AB-ABBC-EFDAF77DBEE2}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.ApplicationDefinedProperties", "{CDBFC167-337E-41D8-AF7C-8C09205429C7}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.ApplicationName", "{F29F85E0-4FF9-1068-AB91-08002B27B3D9}", 18): i = i + 1
    m_PKeys(i) = NamedPKEY("System.AppZoneIdentifier", "{502CFEAB-47EB-459C-B960-E6D8728F7701}", 102): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Author", "{F29F85E0-4FF9-1068-AB91-08002B27B3D9}", 4): i = i + 1
    m_PKeys(i) = NamedPKEY("System.CachedFileUpdaterContentIdForConflictResolution", "{FCEFF153-E839-4CF3-A9E7-EA22832094B8}", 114): i = i + 1
    m_PKeys(i) = NamedPKEY("System.CachedFileUpdaterContentIdForStream", "{FCEFF153-E839-4CF3-A9E7-EA22832094B8}", 113): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Capacity", "{9B174B35-40FF-11D2-A27E-00C04FC30871}", 3): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Category", "{D5CDD502-2E9C-101B-9397-08002B2CF9AE}", 2): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Comment", "{F29F85E0-4FF9-1068-AB91-08002B27B3D9}", 6): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Company", "{D5CDD502-2E9C-101B-9397-08002B2CF9AE}", 15): i = i + 1
    m_PKeys(i) = NamedPKEY("System.ComputerName", "{28636AA6-953D-11D2-B5D6-00C04FD918D0}", 5): i = i + 1
    m_PKeys(i) = NamedPKEY("System.ContainedItems", "{28636AA6-953D-11D2-B5D6-00C04FD918D0}", 29): i = i + 1
    m_PKeys(i) = NamedPKEY("System.ContentStatus", "{D5CDD502-2E9C-101B-9397-08002B2CF9AE}", 27): i = i + 1
    m_PKeys(i) = NamedPKEY("System.ContentType", "{D5CDD502-2E9C-101B-9397-08002B2CF9AE}", 26): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Copyright", "{64440492-4C8B-11D1-8B70-080036B11A03}", 11): i = i + 1
    m_PKeys(i) = NamedPKEY("System.CreatorAppId", "{C2EA046E-033C-4E91-BD5B-D4942F6BBE49}", 2): i = i + 1
    m_PKeys(i) = NamedPKEY("System.CreatorOpenWithUIOptions", "{C2EA046E-033C-4E91-BD5B-D4942F6BBE49}", 3): i = i + 1
    m_PKeys(i) = NamedPKEY("System.DataObjectFormat", "{1E81A3F8-A30F-4247-B9EE-1D0368A9425C}", 2): i = i + 1
    m_PKeys(i) = NamedPKEY("System.DateAccessed", "{B725F130-47EF-101A-A5F1-02608C9EEBAC}", 16): i = i + 1
    m_PKeys(i) = NamedPKEY("System.DateAcquired", "{2CBAA8F5-D81F-47CA-B17A-F8D822300131}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.DateArchived", "{43F8D7B7-A444-4F87-9383-52271C9B915C}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.DateCompleted", "{72FAB781-ACDA-43E5-B155-B2434F85E678}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.DateCreated", "{B725F130-47EF-101A-A5F1-02608C9EEBAC}", 15): i = i + 1
    m_PKeys(i) = NamedPKEY("System.DateImported", "{14B81DA1-0135-4D31-96D9-6CBFC9671A99}", 18258): i = i + 1
    m_PKeys(i) = NamedPKEY("System.DateModified", "{B725F130-47EF-101A-A5F1-02608C9EEBAC}", 14): i = i + 1
    m_PKeys(i) = NamedPKEY("System.DefaultSaveLocationDisplay", "{5D76B67F-9B3D-44BB-B6AE-25DA4F638A67}", 10): i = i + 1
    m_PKeys(i) = NamedPKEY("System.DueDate", "{3F8472B5-E0AF-4DB2-8071-C53FE76AE7CE}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.EndDate", "{C75FAA05-96FD-49E7-9CB4-9F601082D553}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.ExpandoProperties", "{6FA20DE6-D11C-4D9D-A154-64317628C12D}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.FileAllocationSize", "{B725F130-47EF-101A-A5F1-02608C9EEBAC}", 18): i = i + 1
    m_PKeys(i) = NamedPKEY("System.FileAttributes", "{B725F130-47EF-101A-A5F1-02608C9EEBAC}", 13): i = i + 1
    m_PKeys(i) = NamedPKEY("System.FileCount", "{28636AA6-953D-11D2-B5D6-00C04FD918D0}", 12): i = i + 1
    m_PKeys(i) = NamedPKEY("System.FileDescription", "{0CEF7D53-FA64-11D1-A203-0000F81FEDEE}", 3): i = i + 1
    m_PKeys(i) = NamedPKEY("System.FileExtension", "{E4F10A3C-49E6-405D-8288-A23BD4EEAA6C}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.FileFRN", "{B725F130-47EF-101A-A5F1-02608C9EEBAC}", 21): i = i + 1
    m_PKeys(i) = NamedPKEY("System.FileName", "{41CF5AE0-F75A-4806-BD87-59C7D9248EB9}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.FileOfflineAvailabilityStatus", "{FCEFF153-E839-4CF3-A9E7-EA22832094B8}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.FileOwner", "{9B174B34-40FF-11D2-A27E-00C04FC30871}", 4): i = i + 1
    m_PKeys(i) = NamedPKEY("System.FilePlaceholderStatus", "{B2F9B9D6-FEC4-4DD5-94D7-8957488C807B}", 2): i = i + 1
    m_PKeys(i) = NamedPKEY("System.FileVersion", "{0CEF7D53-FA64-11D1-A203-0000F81FEDEE}", 4): i = i + 1
    m_PKeys(i) = NamedPKEY("System.FindData", "{28636AA6-953D-11D2-B5D6-00C04FD918D0}", 0): i = i + 1
    m_PKeys(i) = NamedPKEY("System.FlagColor", "{67DF94DE-0CA7-4D6F-B792-053A3E4F03CF}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.FlagColorText", "{45EAE747-8E2A-40AE-8CBF-CA52ABA6152A}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.FlagStatus", "{E3E0584C-B788-4A5A-BB20-7F5A44C9ACDD}", 12): i = i + 1
    m_PKeys(i) = NamedPKEY("System.FlagStatusText", "{DC54FD2E-189D-4871-AA01-08C2F57A4ABC}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.FolderKind", "{FCEFF153-E839-4CF3-A9E7-EA22832094B8}", 101): i = i + 1
    m_PKeys(i) = NamedPKEY("System.FolderNameDisplay", "{B725F130-47EF-101A-A5F1-02608C9EEBAC}", 25): i = i + 1
    m_PKeys(i) = NamedPKEY("System.FreeSpace", "{9B174B35-40FF-11D2-A27E-00C04FC30871}", 2): i = i + 1
    m_PKeys(i) = NamedPKEY("System.FullText", "{1E3EE840-BC2B-476C-8237-2ACD1A839B22}", 6): i = i + 1
    m_PKeys(i) = NamedPKEY("System.HighKeywords", "{F29F85E0-4FF9-1068-AB91-08002B27B3D9}", 24): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Identity", "{A26F4AFC-7346-4299-BE47-EB1AE613139F}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Identity.Blob", "{8C3B93A4-BAED-1A83-9A32-102EE313F6EB}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Identity.DisplayName", "{7D683FC9-D155-45A8-BB1F-89D19BCB792F}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Identity.InternetSid", "{6D6D5D49-265D-4688-9F4E-1FDD33E7CC83}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Identity.IsMeIdentity", "{A4108708-09DF-4377-9DFC-6D99986D5A67}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Identity.KeyProviderContext", "{A26F4AFC-7346-4299-BE47-EB1AE613139F}", 17): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Identity.KeyProviderName", "{A26F4AFC-7346-4299-BE47-EB1AE613139F}", 16): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Identity.LogonStatusString", "{F18DEDF3-337F-42C0-9E03-CEE08708A8C3}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Identity.PrimaryEmailAddress", "{FCC16823-BAED-4F24-9B32-A0982117F7FA}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Identity.PrimarySid", "{2B1B801E-C0C1-4987-9EC5-72FA89814787}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Identity.ProviderData", "{A8A74B92-361B-4E9A-B722-7C4A7330A312}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Identity.ProviderID", "{74A7DE49-FA11-4D3D-A006-DB7E08675916}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Identity.QualifiedUserName", "{DA520E51-F4E9-4739-AC82-02E0A95C9030}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Identity.UniqueID", "{E55FC3B0-2B60-4220-918E-B21E8BF16016}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Identity.UserName", "{C4322503-78CA-49C6-9ACC-A68E2AFD7B6B}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.IdentityProvider.Name", "{B96EFF7B-35CA-4A35-8607-29E3A54C46EA}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.IdentityProvider.Picture", "{2425166F-5642-4864-992F-98FD98F294C3}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.ImageParsingName", "{D7750EE0-C6A4-48EC-B53E-B87B52E6D073}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Importance", "{E3E0584C-B788-4A5A-BB20-7F5A44C9ACDD}", 11): i = i + 1
    m_PKeys(i) = NamedPKEY("System.ImportanceText", "{A3B29791-7713-4E1D-BB40-17DB85F01831}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.IsAttachment", "{F23F425C-71A1-4FA8-922F-678EA4A60408}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.IsDefaultNonOwnerSaveLocation", "{5D76B67F-9B3D-44BB-B6AE-25DA4F638A67}", 5): i = i + 1
    m_PKeys(i) = NamedPKEY("System.IsDefaultSaveLocation", "{5D76B67F-9B3D-44BB-B6AE-25DA4F638A67}", 3): i = i + 1
    m_PKeys(i) = NamedPKEY("System.IsDeleted", "{5CDA5FC8-33EE-4FF3-9094-AE7BD8868C4D}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.IsEncrypted", "{90E5E14E-648B-4826-B2AA-ACAF790E3513}", 10): i = i + 1
    m_PKeys(i) = NamedPKEY("System.IsFlagged", "{5DA84765-E3FF-4278-86B0-A27967FBDD03}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.IsFlaggedComplete", "{A6F360D2-55F9-48DE-B909-620E090A647C}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.IsIncomplete", "{346C8BD1-2E6A-4C45-89A4-61B78E8E700F}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.IsLocationSupported", "{5D76B67F-9B3D-44BB-B6AE-25DA4F638A67}", 8): i = i + 1
    m_PKeys(i) = NamedPKEY("System.IsPinnedToNameSpaceTree", "{5D76B67F-9B3D-44BB-B6AE-25DA4F638A67}", 2): i = i + 1
    m_PKeys(i) = NamedPKEY("System.IsRead", "{E3E0584C-B788-4A5A-BB20-7F5A44C9ACDD}", 10): i = i + 1
    m_PKeys(i) = NamedPKEY("System.IsSearchOnlyItem", "{5D76B67F-9B3D-44BB-B6AE-25DA4F638A67}", 4): i = i + 1
    m_PKeys(i) = NamedPKEY("System.IsSendToTarget", "{28636AA6-953D-11D2-B5D6-00C04FD918D0}", 33): i = i + 1
    m_PKeys(i) = NamedPKEY("System.IsShared", "{EF884C5B-2BFE-41BB-AAE5-76EEDF4F9902}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.ItemAuthors", "{D0A04F0A-462A-48A4-BB2F-3706E88DBD7D}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.ItemClassType", "{048658AD-2DB8-41A4-BBB6-AC1EF1207EB1}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.ItemDate", "{F7DB74B4-4287-4103-AFBA-F1B13DCD75CF}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.ItemFolderNameDisplay", "{B725F130-47EF-101A-A5F1-02608C9EEBAC}", 2): i = i + 1
    m_PKeys(i) = NamedPKEY("System.ItemFolderPathDisplay", "{E3E0584C-B788-4A5A-BB20-7F5A44C9ACDD}", 6): i = i + 1
    m_PKeys(i) = NamedPKEY("System.ItemFolderPathDisplayNarrow", "{DABD30ED-0043-4789-A7F8-D013A4736622}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.ItemName", "{6B8DA074-3B5C-43BC-886F-0A2CDCE00B6F}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.ItemNameDisplay", "{B725F130-47EF-101A-A5F1-02608C9EEBAC}", 10): i = i + 1
    m_PKeys(i) = NamedPKEY("System.ItemNameDisplayWithoutExtension", "{B725F130-47EF-101A-A5F1-02608C9EEBAC}", 24): i = i + 1
    m_PKeys(i) = NamedPKEY("System.ItemNamePrefix", "{D7313FF1-A77A-401C-8C99-3DBDD68ADD36}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.ItemNameSortOverride", "{B725F130-47EF-101A-A5F1-02608C9EEBAC}", 23): i = i + 1
    m_PKeys(i) = NamedPKEY("System.ItemParticipants", "{D4D0AA16-9948-41A4-AA85-D97FF9646993}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.ItemPathDisplay", "{E3E0584C-B788-4A5A-BB20-7F5A44C9ACDD}", 7): i = i + 1
    m_PKeys(i) = NamedPKEY("System.ItemPathDisplayNarrow", "{28636AA6-953D-11D2-B5D6-00C04FD918D0}", 8): i = i + 1
    m_PKeys(i) = NamedPKEY("System.ItemSubType", "{28636AA6-953D-11D2-B5D6-00C04FD918D0}", 37): i = i + 1
    m_PKeys(i) = NamedPKEY("System.ItemType", "{28636AA6-953D-11D2-B5D6-00C04FD918D0}", 11): i = i + 1
    m_PKeys(i) = NamedPKEY("System.ItemTypeText", "{B725F130-47EF-101A-A5F1-02608C9EEBAC}", 4): i = i + 1
    m_PKeys(i) = NamedPKEY("System.ItemUrl", "{49691C90-7E17-101A-A91C-08002B2ECDA9}", 9): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Keywords", "{F29F85E0-4FF9-1068-AB91-08002B27B3D9}", 5): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Kind", "{1E3EE840-BC2B-476C-8237-2ACD1A839B22}", 3): i = i + 1
    m_PKeys(i) = NamedPKEY("System.KindText", "{F04BEF95-C585-4197-A2B7-DF46FDC9EE6D}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Language", "{D5CDD502-2E9C-101B-9397-08002B2CF9AE}", 28): i = i + 1
    m_PKeys(i) = NamedPKEY("System.LastSyncError", "{FCEFF153-E839-4CF3-A9E7-EA22832094B8}", 107): i = i + 1
    m_PKeys(i) = NamedPKEY("System.LastSyncWarning", "{FCEFF153-E839-4CF3-A9E7-EA22832094B8}", 128): i = i + 1
    m_PKeys(i) = NamedPKEY("System.LastWriterPackageFamilyName", "{502CFEAB-47EB-459C-B960-E6D8728F7701}", 101): i = i + 1
    m_PKeys(i) = NamedPKEY("System.LowKeywords", "{F29F85E0-4FF9-1068-AB91-08002B27B3D9}", 25): i = i + 1
    m_PKeys(i) = NamedPKEY("System.MediumKeywords", "{F29F85E0-4FF9-1068-AB91-08002B27B3D9}", 26): i = i + 1
    m_PKeys(i) = NamedPKEY("System.MileageInformation", "{FDF84370-031A-4ADD-9E91-0D775F1C6605}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.MIMEType", "{0B63E350-9CCC-11D0-BCDB-00805FCCCE04}", 5): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Null", "{00000000-0000-0000-0000-000000000000}", 0): i = i + 1
    m_PKeys(i) = NamedPKEY("System.OfflineAvailability", "{A94688B6-7D9F-4570-A648-E3DFC0AB2B3F}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.OfflineStatus", "{6D24888F-4718-4BDA-AFED-EA0FB4386CD8}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.OriginalFileName", "{0CEF7D53-FA64-11D1-A203-0000F81FEDEE}", 6): i = i + 1
    m_PKeys(i) = NamedPKEY("System.OwnerSID", "{5D76B67F-9B3D-44BB-B6AE-25DA4F638A67}", 6): i = i + 1
    m_PKeys(i) = NamedPKEY("System.ParentalRating", "{64440492-4C8B-11D1-8B70-080036B11A03}", 21): i = i + 1
    m_PKeys(i) = NamedPKEY("System.ParentalRatingReason", "{10984E0A-F9F2-4321-B7EF-BAF195AF4319}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.ParentalRatingsOrganization", "{A7FE0840-1344-46F0-8D37-52ED712A4BF9}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.ParsingBindContext", "{DFB9A04D-362F-4CA3-B30B-0254B17B5B84}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.ParsingName", "{28636AA6-953D-11D2-B5D6-00C04FD918D0}", 24): i = i + 1
    m_PKeys(i) = NamedPKEY("System.ParsingPath", "{28636AA6-953D-11D2-B5D6-00C04FD918D0}", 30): i = i + 1
    m_PKeys(i) = NamedPKEY("System.PerceivedType", "{28636AA6-953D-11D2-B5D6-00C04FD918D0}", 9): i = i + 1
    m_PKeys(i) = NamedPKEY("System.PercentFull", "{9B174B35-40FF-11D2-A27E-00C04FC30871}", 5): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Priority", "{9C1FCF74-2D97-41BA-B4AE-CB2E3661A6E4}", 5): i = i + 1
    m_PKeys(i) = NamedPKEY("System.PriorityText", "{D98BE98B-B86B-4095-BF52-9D23B2E0A752}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Project", "{39A7F922-477C-48DE-8BC8-B28441E342E3}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.ProviderItemID", "{F21D9941-81F0-471A-ADEE-4E74B49217ED}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Rating", "{64440492-4C8B-11D1-8B70-080036B11A03}", 9): i = i + 1
    m_PKeys(i) = NamedPKEY("System.RatingText", "{90197CA7-FD8F-4E8C-9DA3-B57E1E609295}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.RemoteConflictingFile", "{FCEFF153-E839-4CF3-A9E7-EA22832094B8}", 115): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Security.AllowedEnterpriseDataProtectionIdentities", "{38D43380-D418-4830-84D5-46935A81C5C6}", 32): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Security.EncryptionOwners", "{5F5AFF6A-37E5-4780-97EA-80C7565CF535}", 34): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Security.EncryptionOwnersDisplay", "{DE621B8F-E125-43A3-A32D-5665446D632A}", 25): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Sensitivity", "{F8D3F6AC-4874-42CB-BE59-AB454B30716A}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.SensitivityText", "{D0C7F054-3F72-4725-8527-129A577CB269}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.SFGAOFlags", "{28636AA6-953D-11D2-B5D6-00C04FD918D0}", 25): i = i + 1
    m_PKeys(i) = NamedPKEY("System.SharedWith", "{EF884C5B-2BFE-41BB-AAE5-76EEDF4F9902}", 200): i = i + 1
    m_PKeys(i) = NamedPKEY("System.ShareUserRating", "{64440492-4C8B-11D1-8B70-080036B11A03}", 12): i = i + 1
    m_PKeys(i) = NamedPKEY("System.SharingStatus", "{EF884C5B-2BFE-41BB-AAE5-76EEDF4F9902}", 300): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Shell.OmitFromView", "{DE35258C-C695-4CBC-B982-38B0AD24CED0}", 2): i = i + 1
    m_PKeys(i) = NamedPKEY("System.SimpleRating", "{A09F084E-AD41-489F-8076-AA5BE3082BCA}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Size", "{B725F130-47EF-101A-A5F1-02608C9EEBAC}", 12): i = i + 1
    m_PKeys(i) = NamedPKEY("System.SoftwareUsed", "{14B81DA1-0135-4D31-96D9-6CBFC9671A99}", 305): i = i + 1
    m_PKeys(i) = NamedPKEY("System.SourceItem", "{668CDFA5-7A1B-4323-AE4B-E527393A1D81}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.SourcePackageFamilyName", "{FFAE9DB7-1C8D-43FF-818C-84403AA3732D}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.StartDate", "{48FD6EC8-8A12-4CDF-A03E-4EC5A511EDDE}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Status", "{000214A1-0000-0000-C000-000000000046}", 9): i = i + 1
    m_PKeys(i) = NamedPKEY("System.StorageProviderCallerVersionInformation", "{B2F9B9D6-FEC4-4DD5-94D7-8957488C807B}", 7): i = i + 1
    m_PKeys(i) = NamedPKEY("System.StorageProviderError", "{FCEFF153-E839-4CF3-A9E7-EA22832094B8}", 109): i = i + 1
    m_PKeys(i) = NamedPKEY("System.StorageProviderFileChecksum", "{B2F9B9D6-FEC4-4DD5-94D7-8957488C807B}", 5): i = i + 1
    m_PKeys(i) = NamedPKEY("System.StorageProviderFileFlags", "{B2F9B9D6-FEC4-4DD5-94D7-8957488C807B}", 8): i = i + 1
    m_PKeys(i) = NamedPKEY("System.StorageProviderFileIdentifier", "{B2F9B9D6-FEC4-4DD5-94D7-8957488C807B}", 3): i = i + 1
    m_PKeys(i) = NamedPKEY("System.StorageProviderFileRemoteUri", "{FCEFF153-E839-4CF3-A9E7-EA22832094B8}", 112): i = i + 1
    m_PKeys(i) = NamedPKEY("System.StorageProviderFileVersion", "{B2F9B9D6-FEC4-4DD5-94D7-8957488C807B}", 4): i = i + 1
    m_PKeys(i) = NamedPKEY("System.StorageProviderFileVersionWaterline", "{B2F9B9D6-FEC4-4DD5-94D7-8957488C807B}", 6): i = i + 1
    m_PKeys(i) = NamedPKEY("System.StorageProviderId", "{FCEFF153-E839-4CF3-A9E7-EA22832094B8}", 108): i = i + 1
    m_PKeys(i) = NamedPKEY("System.StorageProviderShareStatuses", "{FCEFF153-E839-4CF3-A9E7-EA22832094B8}", 111): i = i + 1
    m_PKeys(i) = NamedPKEY("System.StorageProviderSharingStatus", "{FCEFF153-E839-4CF3-A9E7-EA22832094B8}", 117): i = i + 1
    m_PKeys(i) = NamedPKEY("System.StorageProviderStatus", "{FCEFF153-E839-4CF3-A9E7-EA22832094B8}", 110): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Subject", "{F29F85E0-4FF9-1068-AB91-08002B27B3D9}", 3): i = i + 1
    m_PKeys(i) = NamedPKEY("System.SyncTransferStatus", "{FCEFF153-E839-4CF3-A9E7-EA22832094B8}", 103): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Thumbnail", "{F29F85E0-4FF9-1068-AB91-08002B27B3D9}", 17): i = i + 1
    m_PKeys(i) = NamedPKEY("System.ThumbnailCacheId", "{446D16B1-8DAD-4870-A748-402EA43D788C}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.ThumbnailStream", "{F29F85E0-4FF9-1068-AB91-08002B27B3D9}", 27): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Title", "{F29F85E0-4FF9-1068-AB91-08002B27B3D9}", 2): i = i + 1
    m_PKeys(i) = NamedPKEY("System.TitleSortOverride", "{F0F7984D-222E-4AD2-82AB-1DD8EA40E57E}", 300): i = i + 1
    m_PKeys(i) = NamedPKEY("System.TotalFileSize", "{28636AA6-953D-11D2-B5D6-00C04FD918D0}", 14): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Trademarks", "{0CEF7D53-FA64-11D1-A203-0000F81FEDEE}", 9): i = i + 1
    m_PKeys(i) = NamedPKEY("System.TransferOrder", "{FCEFF153-E839-4CF3-A9E7-EA22832094B8}", 106): i = i + 1
    m_PKeys(i) = NamedPKEY("System.TransferPosition", "{FCEFF153-E839-4CF3-A9E7-EA22832094B8}", 104): i = i + 1
    m_PKeys(i) = NamedPKEY("System.TransferSize", "{FCEFF153-E839-4CF3-A9E7-EA22832094B8}", 105): i = i + 1
    m_PKeys(i) = NamedPKEY("System.VolumeId", "{446D16B1-8DAD-4870-A748-402EA43D788C}", 104): i = i + 1
    m_PKeys(i) = NamedPKEY("System.ZoneIdentifier", "{502CFEAB-47EB-459C-B960-E6D8728F7701}", 100): i = i + 1
    ' end part1
    Init2 i
End Sub

Private Sub Init2(ByRef i As Long)
    m_PKeys(i) = NamedPKEY("System.Device.PrinterURL", "{0B48F35A-BE6E-4F17-B108-3C4073D1669A}", 15): i = i + 1
    m_PKeys(i) = NamedPKEY("System.DeviceInterface.Bluetooth.DeviceAddress", "{2BD67D8B-8BEB-48D5-87E0-6CDA3428040A}", 1): i = i + 1
    m_PKeys(i) = NamedPKEY("System.DeviceInterface.Bluetooth.Flags", "{2BD67D8B-8BEB-48D5-87E0-6CDA3428040A}", 3): i = i + 1
    m_PKeys(i) = NamedPKEY("System.DeviceInterface.Bluetooth.LastConnectedTime", "{2BD67D8B-8BEB-48D5-87E0-6CDA3428040A}", 11): i = i + 1
    m_PKeys(i) = NamedPKEY("System.DeviceInterface.Bluetooth.Manufacturer", "{2BD67D8B-8BEB-48D5-87E0-6CDA3428040A}", 4): i = i + 1
    m_PKeys(i) = NamedPKEY("System.DeviceInterface.Bluetooth.ModelNumber", "{2BD67D8B-8BEB-48D5-87E0-6CDA3428040A}", 5): i = i + 1
    m_PKeys(i) = NamedPKEY("System.DeviceInterface.Bluetooth.ProductId", "{2BD67D8B-8BEB-48D5-87E0-6CDA3428040A}", 8): i = i + 1
    m_PKeys(i) = NamedPKEY("System.DeviceInterface.Bluetooth.ProductVersion", "{2BD67D8B-8BEB-48D5-87E0-6CDA3428040A}", 9): i = i + 1
    m_PKeys(i) = NamedPKEY("System.DeviceInterface.Bluetooth.ServiceGuid", "{2BD67D8B-8BEB-48D5-87E0-6CDA3428040A}", 2): i = i + 1
    m_PKeys(i) = NamedPKEY("System.DeviceInterface.Bluetooth.VendorId", "{2BD67D8B-8BEB-48D5-87E0-6CDA3428040A}", 7): i = i + 1
    m_PKeys(i) = NamedPKEY("System.DeviceInterface.Bluetooth.VendorIdSource", "{2BD67D8B-8BEB-48D5-87E0-6CDA3428040A}", 6): i = i + 1
    m_PKeys(i) = NamedPKEY("System.DeviceInterface.Hid.IsReadOnly", "{CBF38310-4A17-4310-A1EB-247F0B67593B}", 4): i = i + 1
    m_PKeys(i) = NamedPKEY("System.DeviceInterface.Hid.ProductId", "{CBF38310-4A17-4310-A1EB-247F0B67593B}", 6): i = i + 1
    m_PKeys(i) = NamedPKEY("System.DeviceInterface.Hid.UsageId", "{CBF38310-4A17-4310-A1EB-247F0B67593B}", 3): i = i + 1
    m_PKeys(i) = NamedPKEY("System.DeviceInterface.Hid.UsagePage", "{CBF38310-4A17-4310-A1EB-247F0B67593B}", 2): i = i + 1
    m_PKeys(i) = NamedPKEY("System.DeviceInterface.Hid.VendorId", "{CBF38310-4A17-4310-A1EB-247F0B67593B}", 5): i = i + 1
    m_PKeys(i) = NamedPKEY("System.DeviceInterface.Hid.VersionNumber", "{CBF38310-4A17-4310-A1EB-247F0B67593B}", 7): i = i + 1
    m_PKeys(i) = NamedPKEY("System.DeviceInterface.PrinterDriverDirectory", "{847C66DE-B8D6-4AF9-ABC3-6F4F926BC039}", 14): i = i + 1
    m_PKeys(i) = NamedPKEY("System.DeviceInterface.PrinterDriverName", "{AFC47170-14F5-498C-8F30-B0D19BE449C6}", 11): i = i + 1
    m_PKeys(i) = NamedPKEY("System.DeviceInterface.PrinterEnumerationFlag", "{A00742A1-CD8C-4B37-95AB-70755587767A}", 3): i = i + 1
    m_PKeys(i) = NamedPKEY("System.DeviceInterface.PrinterName", "{0A7B84EF-0C27-463F-84EF-06C5070001BE}", 10): i = i + 1
    m_PKeys(i) = NamedPKEY("System.DeviceInterface.PrinterPortName", "{EEC7B761-6F94-41B1-949F-C729720DD13C}", 12): i = i + 1
    m_PKeys(i) = NamedPKEY("System.DeviceInterface.Proximity.SupportsNfc", "{FB3842CD-9E2A-4F83-8FCC-4B0761139AE9}", 2): i = i + 1
    m_PKeys(i) = NamedPKEY("System.DeviceInterface.Serial.PortName", "{4C6BF15C-4C03-4AAC-91F5-64C0F852BCF4}", 4): i = i + 1
    m_PKeys(i) = NamedPKEY("System.DeviceInterface.Serial.UsbProductId", "{4C6BF15C-4C03-4AAC-91F5-64C0F852BCF4}", 3): i = i + 1
    m_PKeys(i) = NamedPKEY("System.DeviceInterface.Serial.UsbVendorId", "{4C6BF15C-4C03-4AAC-91F5-64C0F852BCF4}", 2): i = i + 1
    m_PKeys(i) = NamedPKEY("System.DeviceInterface.WinUsb.DeviceInterfaceClasses", "{95E127B5-79CC-4E83-9C9E-8422187B3E0E}", 7): i = i + 1
    m_PKeys(i) = NamedPKEY("System.DeviceInterface.WinUsb.UsbClass", "{95E127B5-79CC-4E83-9C9E-8422187B3E0E}", 4): i = i + 1
    m_PKeys(i) = NamedPKEY("System.DeviceInterface.WinUsb.UsbProductId", "{95E127B5-79CC-4E83-9C9E-8422187B3E0E}", 3): i = i + 1
    m_PKeys(i) = NamedPKEY("System.DeviceInterface.WinUsb.UsbProtocol", "{95E127B5-79CC-4E83-9C9E-8422187B3E0E}", 6): i = i + 1
    m_PKeys(i) = NamedPKEY("System.DeviceInterface.WinUsb.UsbSubClass", "{95E127B5-79CC-4E83-9C9E-8422187B3E0E}", 5): i = i + 1
    m_PKeys(i) = NamedPKEY("System.DeviceInterface.WinUsb.UsbVendorId", "{95E127B5-79CC-4E83-9C9E-8422187B3E0E}", 2): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.Aep.AepId", "{3B2CE006-5E61-4FDE-BAB8-9B8AAC9B26DF}", 8): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.Aep.Bluetooth.Cod.Major", "{5FBD34CD-561A-412E-BA98-478A6B0FEF1D}", 2): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.Aep.Bluetooth.Cod.Minor", "{5FBD34CD-561A-412E-BA98-478A6B0FEF1D}", 3): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.Aep.Bluetooth.Cod.Services.Audio", "{5FBD34CD-561A-412E-BA98-478A6B0FEF1D}", 10): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.Aep.Bluetooth.Cod.Services.Capturing", "{5FBD34CD-561A-412E-BA98-478A6B0FEF1D}", 8): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.Aep.Bluetooth.Cod.Services.Information", "{5FBD34CD-561A-412E-BA98-478A6B0FEF1D}", 12): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.Aep.Bluetooth.Cod.Services.LimitedDiscovery", "{5FBD34CD-561A-412E-BA98-478A6B0FEF1D}", 4): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.Aep.Bluetooth.Cod.Services.Networking", "{5FBD34CD-561A-412E-BA98-478A6B0FEF1D}", 6): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.Aep.Bluetooth.Cod.Services.ObjectXfer", "{5FBD34CD-561A-412E-BA98-478A6B0FEF1D}", 9): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.Aep.Bluetooth.Cod.Services.Positioning", "{5FBD34CD-561A-412E-BA98-478A6B0FEF1D}", 5): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.Aep.Bluetooth.Cod.Services.Rendering", "{5FBD34CD-561A-412E-BA98-478A6B0FEF1D}", 7): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.Aep.Bluetooth.Cod.Services.Telephony", "{5FBD34CD-561A-412E-BA98-478A6B0FEF1D}", 11): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.Aep.Bluetooth.LastSeenTime", "{2BD67D8B-8BEB-48D5-87E0-6CDA3428040A}", 12): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.Aep.Bluetooth.Le.AddressType", "{995EF0B0-7EB3-4A8B-B9CE-068BB3F4AF69}", 4): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.Aep.Bluetooth.Le.Appearance", "{995EF0B0-7EB3-4A8B-B9CE-068BB3F4AF69}", 1): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.Aep.Bluetooth.Le.Appearance.Category", "{995EF0B0-7EB3-4A8B-B9CE-068BB3F4AF69}", 5): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.Aep.Bluetooth.Le.Appearance.Subcategory", "{995EF0B0-7EB3-4A8B-B9CE-068BB3F4AF69}", 6): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.Aep.Bluetooth.Le.IsConnectable", "{995EF0B0-7EB3-4A8B-B9CE-068BB3F4AF69}", 8): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.Aep.CanPair", "{E7C3FB29-CAA7-4F47-8C8B-BE59B330D4C5}", 3): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.Aep.Category", "{A35996AB-11CF-4935-8B61-A6761081ECDF}", 17): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.Aep.ContainerId", "{E7C3FB29-CAA7-4F47-8C8B-BE59B330D4C5}", 2): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.Aep.DeviceAddress", "{A35996AB-11CF-4935-8B61-A6761081ECDF}", 12): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.Aep.IsConnected", "{A35996AB-11CF-4935-8B61-A6761081ECDF}", 7): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.Aep.IsPaired", "{A35996AB-11CF-4935-8B61-A6761081ECDF}", 16): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.Aep.IsPresent", "{A35996AB-11CF-4935-8B61-A6761081ECDF}", 9): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.Aep.Manufacturer", "{A35996AB-11CF-4935-8B61-A6761081ECDF}", 5): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.Aep.ModelId", "{A35996AB-11CF-4935-8B61-A6761081ECDF}", 4): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.Aep.ModelName", "{A35996AB-11CF-4935-8B61-A6761081ECDF}", 3): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.Aep.PointOfService.ConnectionTypes", "{D4BF61B3-442E-4ADA-882D-FA7B70C832D9}", 6): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.Aep.ProtocolId", "{3B2CE006-5E61-4FDE-BAB8-9B8AAC9B26DF}", 5): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.Aep.SignalStrength", "{A35996AB-11CF-4935-8B61-A6761081ECDF}", 6): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.AepContainer.CanPair", "{0BBA1EDE-7566-4F47-90EC-25FC567CED2A}", 3): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.AepContainer.Categories", "{0BBA1EDE-7566-4F47-90EC-25FC567CED2A}", 9): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.AepContainer.Children", "{0BBA1EDE-7566-4F47-90EC-25FC567CED2A}", 2): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.AepContainer.ContainerId", "{0BBA1EDE-7566-4F47-90EC-25FC567CED2A}", 12): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.AepContainer.DialProtocol.InstalledApplications", "{6AF55D45-38DB-4495-ACB0-D4728A3B8314}", 6): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.AepContainer.IsPaired", "{0BBA1EDE-7566-4F47-90EC-25FC567CED2A}", 4): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.AepContainer.IsPresent", "{0BBA1EDE-7566-4F47-90EC-25FC567CED2A}", 11): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.AepContainer.Manufacturer", "{0BBA1EDE-7566-4F47-90EC-25FC567CED2A}", 6): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.AepContainer.ModelIds", "{0BBA1EDE-7566-4F47-90EC-25FC567CED2A}", 8): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.AepContainer.ModelName", "{0BBA1EDE-7566-4F47-90EC-25FC567CED2A}", 7): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.AepContainer.ProtocolIds", "{0BBA1EDE-7566-4F47-90EC-25FC567CED2A}", 13): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.AepContainer.SupportedUriSchemes", "{6AF55D45-38DB-4495-ACB0-D4728A3B8314}", 5): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.AepContainer.SupportsAudio", "{6AF55D45-38DB-4495-ACB0-D4728A3B8314}", 2): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.AepContainer.SupportsCapturing", "{6AF55D45-38DB-4495-ACB0-D4728A3B8314}", 11): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.AepContainer.SupportsImages", "{6AF55D45-38DB-4495-ACB0-D4728A3B8314}", 4): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.AepContainer.SupportsInformation", "{6AF55D45-38DB-4495-ACB0-D4728A3B8314}", 14): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.AepContainer.SupportsLimitedDiscovery", "{6AF55D45-38DB-4495-ACB0-D4728A3B8314}", 7): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.AepContainer.SupportsNetworking", "{6AF55D45-38DB-4495-ACB0-D4728A3B8314}", 9): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.AepContainer.SupportsObjectTransfer", "{6AF55D45-38DB-4495-ACB0-D4728A3B8314}", 12): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.AepContainer.SupportsPositioning", "{6AF55D45-38DB-4495-ACB0-D4728A3B8314}", 8): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.AepContainer.SupportsRendering", "{6AF55D45-38DB-4495-ACB0-D4728A3B8314}", 10): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.AepContainer.SupportsTelephony", "{6AF55D45-38DB-4495-ACB0-D4728A3B8314}", 13): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.AepContainer.SupportsVideo", "{6AF55D45-38DB-4495-ACB0-D4728A3B8314}", 3): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.AepService.AepId", "{C9C141A9-1B4C-4F17-A9D1-F298538CADB8}", 6): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.AepService.Bluetooth.CacheMode", "{9744311E-7951-4B2E-B6F0-ECB293CAC119}", 5): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.AepService.Bluetooth.ServiceGuid", "{A399AAC7-C265-474E-B073-FFCE57721716}", 2): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.AepService.Bluetooth.TargetDevice", "{9744311E-7951-4B2E-B6F0-ECB293CAC119}", 6): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.AepService.ContainerId", "{71724756-3E74-4432-9B59-E7B2F668A593}", 4): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.AepService.FriendlyName", "{71724756-3E74-4432-9B59-E7B2F668A593}", 2): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.AepService.IoT.ServiceInterfaces", "{79D94E82-4D79-45AA-821A-74858B4E4CA6}", 2): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.AepService.ParentAepIsPaired", "{C9C141A9-1B4C-4F17-A9D1-F298538CADB8}", 7): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.AepService.ProtocolId", "{C9C141A9-1B4C-4F17-A9D1-F298538CADB8}", 5): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.AepService.ServiceClassId", "{71724756-3E74-4432-9B59-E7B2F668A593}", 3): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.AepService.ServiceId", "{C9C141A9-1B4C-4F17-A9D1-F298538CADB8}", 2): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.AppPackageFamilyName", "{51236583-0C4A-4FE8-B81F-166AEC13F510}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.AudioDevice.Microphone.SensitivityInDbfs", "{8943B373-388C-4395-B557-BC6DBAFFAFDB}", 3): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.AudioDevice.Microphone.SignalToNoiseRatioInDb", "{8943B373-388C-4395-B557-BC6DBAFFAFDB}", 4): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.AudioDevice.RawProcessingSupported", "{8943B373-388C-4395-B557-BC6DBAFFAFDB}", 2): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.AudioDevice.SpeechProcessingSupported", "{FB1DE864-E06D-47F4-82A6-8A0AEF44493C}", 2): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.BatteryLife", "{49CD1F76-5626-4B17-A4E8-18B4AA1A2213}", 10): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.BatteryPlusCharging", "{49CD1F76-5626-4B17-A4E8-18B4AA1A2213}", 22): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.BatteryPlusChargingText", "{49CD1F76-5626-4B17-A4E8-18B4AA1A2213}", 23): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.Category", "{78C34FC8-104A-4ACA-9EA4-524D52996E57}", 91): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.CategoryGroup", "{78C34FC8-104A-4ACA-9EA4-524D52996E57}", 94): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.CategoryIds", "{78C34FC8-104A-4ACA-9EA4-524D52996E57}", 90): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.CategoryPlural", "{78C34FC8-104A-4ACA-9EA4-524D52996E57}", 92): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.ChargingState", "{49CD1F76-5626-4B17-A4E8-18B4AA1A2213}", 11): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.Children", "{4340A6C5-93FA-4706-972C-7B648008A5A7}", 9): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.ClassGuid", "{A45C254E-DF1C-4EFD-8020-67D146A850E0}", 10): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.CompatibleIds", "{A45C254E-DF1C-4EFD-8020-67D146A850E0}", 4): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.Connected", "{78C34FC8-104A-4ACA-9EA4-524D52996E57}", 55): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.ContainerId", "{8C7ED206-3F8A-4827-B3AB-AE9E1FAEFC6C}", 2): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.DefaultTooltip", "{880F70A2-6082-47AC-8AAB-A739D1A300C3}", 153): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.DeviceCapabilities", "{A45C254E-DF1C-4EFD-8020-67D146A850E0}", 17): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.DeviceCharacteristics", "{A45C254E-DF1C-4EFD-8020-67D146A850E0}", 29): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.DeviceDescription1", "{78C34FC8-104A-4ACA-9EA4-524D52996E57}", 81): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.DeviceDescription2", "{78C34FC8-104A-4ACA-9EA4-524D52996E57}", 82): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.DeviceHasProblem", "{540B947E-8B40-45BC-A8A2-6A0B894CBDA2}", 6): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.DeviceInstanceId", "{78C34FC8-104A-4ACA-9EA4-524D52996E57}", 256): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.DeviceManufacturer", "{A45C254E-DF1C-4EFD-8020-67D146A850E0}", 13): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.DevObjectType", "{13673F42-A3D6-49F6-B4DA-AE46E0C5237C}", 2): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.DialProtocol.InstalledApplications", "{6845CC72-1B71-48C3-AF86-B09171A19B14}", 3): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.DiscoveryMethod", "{78C34FC8-104A-4ACA-9EA4-524D52996E57}", 52): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.Dnssd.Domain", "{BF79C0AB-BB74-4CEE-B070-470B5AE202EA}", 3): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.Dnssd.FullName", "{BF79C0AB-BB74-4CEE-B070-470B5AE202EA}", 5): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.Dnssd.HostName", "{BF79C0AB-BB74-4CEE-B070-470B5AE202EA}", 7): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.Dnssd.InstanceName", "{BF79C0AB-BB74-4CEE-B070-470B5AE202EA}", 4): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.Dnssd.NetworkAdapterId", "{BF79C0AB-BB74-4CEE-B070-470B5AE202EA}", 11): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.Dnssd.PortNumber", "{BF79C0AB-BB74-4CEE-B070-470B5AE202EA}", 12): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.Dnssd.Priority", "{BF79C0AB-BB74-4CEE-B070-470B5AE202EA}", 9): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.Dnssd.ServiceName", "{BF79C0AB-BB74-4CEE-B070-470B5AE202EA}", 2): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.Dnssd.TextAttributes", "{BF79C0AB-BB74-4CEE-B070-470B5AE202EA}", 6): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.Dnssd.Ttl", "{BF79C0AB-BB74-4CEE-B070-470B5AE202EA}", 10): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.Dnssd.Weight", "{BF79C0AB-BB74-4CEE-B070-470B5AE202EA}", 8): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.FriendlyName", "{656A3BB3-ECC0-43FD-8477-4AE0404A96CD}", 12288): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.FunctionPaths", "{D08DD4C0-3A9E-462E-8290-7B636B2576B9}", 3): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.GlyphIcon", "{51236583-0C4A-4FE8-B81F-166AEC13F510}", 123): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.HardwareIds", "{A45C254E-DF1C-4EFD-8020-67D146A850E0}", 3): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.Icon", "{78C34FC8-104A-4ACA-9EA4-524D52996E57}", 57): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.InLocalMachineContainer", "{8C7ED206-3F8A-4827-B3AB-AE9E1FAEFC6C}", 4): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.InterfaceClassGuid", "{026E516E-B814-414B-83CD-856D6FEF4822}", 4): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.InterfaceEnabled", "{026E516E-B814-414B-83CD-856D6FEF4822}", 3): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.InterfacePaths", "{D08DD4C0-3A9E-462E-8290-7B636B2576B9}", 2): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.IpAddress", "{656A3BB3-ECC0-43FD-8477-4AE0404A96CD}", 12297): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.IsDefault", "{78C34FC8-104A-4ACA-9EA4-524D52996E57}", 86): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.IsNetworkConnected", "{78C34FC8-104A-4ACA-9EA4-524D52996E57}", 85): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.IsShared", "{78C34FC8-104A-4ACA-9EA4-524D52996E57}", 84): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.IsSoftwareInstalling", "{83DA6326-97A6-4088-9453-A1923F573B29}", 9): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.LaunchDeviceStageFromExplorer", "{78C34FC8-104A-4ACA-9EA4-524D52996E57}", 77): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.LocalMachine", "{78C34FC8-104A-4ACA-9EA4-524D52996E57}", 70): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.LocationPaths", "{A45C254E-DF1C-4EFD-8020-67D146A850E0}", 37): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.Manufacturer", "{656A3BB3-ECC0-43FD-8477-4AE0404A96CD}", 8192): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.MetadataPath", "{78C34FC8-104A-4ACA-9EA4-524D52996E57}", 71): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.MicrophoneArray.Geometry", "{A1829EA2-27EB-459E-935D-B2FAD7B07762}", 2): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.MissedCalls", "{49CD1F76-5626-4B17-A4E8-18B4AA1A2213}", 5): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.ModelId", "{80D81EA6-7473-4B0C-8216-EFC11A2C4C8B}", 2): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.ModelName", "{656A3BB3-ECC0-43FD-8477-4AE0404A96CD}", 8194): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.ModelNumber", "{656A3BB3-ECC0-43FD-8477-4AE0404A96CD}", 8195): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.NetworkedTooltip", "{880F70A2-6082-47AC-8AAB-A739D1A300C3}", 152): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.NetworkName", "{49CD1F76-5626-4B17-A4E8-18B4AA1A2213}", 7): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.NetworkType", "{49CD1F76-5626-4B17-A4E8-18B4AA1A2213}", 8): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.NewPictures", "{49CD1F76-5626-4B17-A4E8-18B4AA1A2213}", 4): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.Notification", "{06704B0C-E830-4C81-9178-91E4E95A80A0}", 3): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.Notifications.LowBattery", "{C4C07F2B-8524-4E66-AE3A-A6235F103BEB}", 2): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.Notifications.MissedCall", "{6614EF48-4EFE-4424-9EDA-C79F404EDF3E}", 2): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.Notifications.NewMessage", "{2BE9260A-2012-4742-A555-F41B638B7DCB}", 2): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.Notifications.NewVoicemail", "{59569556-0A08-4212-95B9-FAE2AD6413DB}", 2): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.Notifications.StorageFull", "{A0E00EE1-F0C7-4D41-B8E7-26A7BD8D38B0}", 2): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.Notifications.StorageFullLinkText", "{A0E00EE1-F0C7-4D41-B8E7-26A7BD8D38B0}", 3): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.NotificationStore", "{06704B0C-E830-4C81-9178-91E4E95A80A0}", 2): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.NotWorkingProperly", "{78C34FC8-104A-4ACA-9EA4-524D52996E57}", 83): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.Paired", "{78C34FC8-104A-4ACA-9EA4-524D52996E57}", 56): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.Parent", "{4340A6C5-93FA-4706-972C-7B648008A5A7}", 8): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.PhysicalDeviceLocation", "{540B947E-8B40-45BC-A8A2-6A0B894CBDA2}", 9): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.PlaybackPositionPercent", "{3633DE59-6825-4381-A49B-9F6BA13A1471}", 5): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.PlaybackState", "{3633DE59-6825-4381-A49B-9F6BA13A1471}", 2): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.PlaybackTitle", "{3633DE59-6825-4381-A49B-9F6BA13A1471}", 3): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.Present", "{540B947E-8B40-45BC-A8A2-6A0B894CBDA2}", 5): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.PresentationUrl", "{656A3BB3-ECC0-43FD-8477-4AE0404A96CD}", 8198): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.PrimaryCategory", "{D08DD4C0-3A9E-462E-8290-7B636B2576B9}", 10): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.RemainingDuration", "{3633DE59-6825-4381-A49B-9F6BA13A1471}", 4): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.RestrictedInterface", "{026E516E-B814-414B-83CD-856D6FEF4822}", 6): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.Roaming", "{49CD1F76-5626-4B17-A4E8-18B4AA1A2213}", 9): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.SafeRemovalRequired", "{AFD97640-86A3-4210-B67C-289C41AABE55}", 2): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.SchematicName", "{026E516E-B814-414B-83CD-856D6FEF4822}", 9): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.ServiceAddress", "{656A3BB3-ECC0-43FD-8477-4AE0404A96CD}", 16384): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.ServiceId", "{656A3BB3-ECC0-43FD-8477-4AE0404A96CD}", 16385): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.SharedTooltip", "{880F70A2-6082-47AC-8AAB-A739D1A300C3}", 151): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.SignalStrength", "{49CD1F76-5626-4B17-A4E8-18B4AA1A2213}", 2): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.SmartCards.ReaderKind", "{D6B5B883-18BD-4B4D-B2EC-9E38AFFEDA82}", 2): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.Status", "{D08DD4C0-3A9E-462E-8290-7B636B2576B9}", 259): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.Status1", "{D08DD4C0-3A9E-462E-8290-7B636B2576B9}", 257): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.Status2", "{D08DD4C0-3A9E-462E-8290-7B636B2576B9}", 258): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.StorageCapacity", "{49CD1F76-5626-4B17-A4E8-18B4AA1A2213}", 12): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.StorageFreeSpace", "{49CD1F76-5626-4B17-A4E8-18B4AA1A2213}", 13): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.StorageFreeSpacePercent", "{49CD1F76-5626-4B17-A4E8-18B4AA1A2213}", 14): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.TextMessages", "{49CD1F76-5626-4B17-A4E8-18B4AA1A2213}", 3): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.Voicemail", "{49CD1F76-5626-4B17-A4E8-18B4AA1A2213}", 6): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.WiaDeviceType", "{6BDD1FC6-810F-11D0-BEC7-08002BE2092F}", 2): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.WiFi.InterfaceGuid", "{EF1167EB-CBFC-4341-A568-A7C91A68982C}", 2): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.WiFiDirect.DeviceAddress", "{1506935D-E3E7-450F-8637-82233EBE5F6E}", 13): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.WiFiDirect.GroupId", "{1506935D-E3E7-450F-8637-82233EBE5F6E}", 4): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.WiFiDirect.InformationElements", "{1506935D-E3E7-450F-8637-82233EBE5F6E}", 12): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.WiFiDirect.InterfaceAddress", "{1506935D-E3E7-450F-8637-82233EBE5F6E}", 2): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.WiFiDirect.InterfaceGuid", "{1506935D-E3E7-450F-8637-82233EBE5F6E}", 3): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.WiFiDirect.IsConnected", "{1506935D-E3E7-450F-8637-82233EBE5F6E}", 5): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.WiFiDirect.IsLegacyDevice", "{1506935D-E3E7-450F-8637-82233EBE5F6E}", 7): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.WiFiDirect.IsMiracastLcpSupported", "{1506935D-E3E7-450F-8637-82233EBE5F6E}", 9): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.WiFiDirect.IsVisible", "{1506935D-E3E7-450F-8637-82233EBE5F6E}", 6): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.WiFiDirect.MiracastVersion", "{1506935D-E3E7-450F-8637-82233EBE5F6E}", 8): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.WiFiDirect.Services", "{1506935D-E3E7-450F-8637-82233EBE5F6E}", 10): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.WiFiDirect.SupportedChannelList", "{1506935D-E3E7-450F-8637-82233EBE5F6E}", 11): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.WiFiDirectServices.AdvertisementId", "{31B37743-7C5E-4005-93E6-E953F92B82E9}", 5): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.WiFiDirectServices.RequestServiceInformation", "{31B37743-7C5E-4005-93E6-E953F92B82E9}", 7): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.WiFiDirectServices.ServiceAddress", "{31B37743-7C5E-4005-93E6-E953F92B82E9}", 2): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.WiFiDirectServices.ServiceConfigMethods", "{31B37743-7C5E-4005-93E6-E953F92B82E9}", 6): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.WiFiDirectServices.ServiceInformation", "{31B37743-7C5E-4005-93E6-E953F92B82E9}", 4): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.WiFiDirectServices.ServiceName", "{31B37743-7C5E-4005-93E6-E953F92B82E9}", 3): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.WinPhone8CameraFlags", "{B7B4D61C-5A64-4187-A52E-B1539F359099}", 2): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Devices.Wwan.InterfaceGuid", "{FF1167EB-CBFC-4341-A568-A7C91A68982C}", 2): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Storage.Portable", "{4D1EBEE8-0803-4774-9842-B77DB50265E9}", 2): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Storage.RemovableMedia", "{4D1EBEE8-0803-4774-9842-B77DB50265E9}", 3): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Storage.SystemCritical", "{4D1EBEE8-0803-4774-9842-B77DB50265E9}", 4): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Document.ByteCount", "{D5CDD502-2E9C-101B-9397-08002B2CF9AE}", 4): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Document.CharacterCount", "{F29F85E0-4FF9-1068-AB91-08002B27B3D9}", 16): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Document.ClientID", "{276D7BB0-5B34-4FB0-AA4B-158ED12A1809}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Document.Contributor", "{F334115E-DA1B-4509-9B3D-119504DC7ABB}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Document.DateCreated", "{F29F85E0-4FF9-1068-AB91-08002B27B3D9}", 12): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Document.DatePrinted", "{F29F85E0-4FF9-1068-AB91-08002B27B3D9}", 11): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Document.DateSaved", "{F29F85E0-4FF9-1068-AB91-08002B27B3D9}", 13): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Document.Division", "{1E005EE6-BF27-428B-B01C-79676ACD2870}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Document.DocumentID", "{E08805C8-E395-40DF-80D2-54F0D6C43154}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Document.HiddenSlideCount", "{D5CDD502-2E9C-101B-9397-08002B2CF9AE}", 9): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Document.LastAuthor", "{F29F85E0-4FF9-1068-AB91-08002B27B3D9}", 8): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Document.LineCount", "{D5CDD502-2E9C-101B-9397-08002B2CF9AE}", 5): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Document.Manager", "{D5CDD502-2E9C-101B-9397-08002B2CF9AE}", 14): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Document.MultimediaClipCount", "{D5CDD502-2E9C-101B-9397-08002B2CF9AE}", 10): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Document.NoteCount", "{D5CDD502-2E9C-101B-9397-08002B2CF9AE}", 8): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Document.PageCount", "{F29F85E0-4FF9-1068-AB91-08002B27B3D9}", 14): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Document.ParagraphCount", "{D5CDD502-2E9C-101B-9397-08002B2CF9AE}", 6): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Document.PresentationFormat", "{D5CDD502-2E9C-101B-9397-08002B2CF9AE}", 3): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Document.RevisionNumber", "{F29F85E0-4FF9-1068-AB91-08002B27B3D9}", 9): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Document.Security", "{F29F85E0-4FF9-1068-AB91-08002B27B3D9}", 19): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Document.SlideCount", "{D5CDD502-2E9C-101B-9397-08002B2CF9AE}", 7): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Document.Template", "{F29F85E0-4FF9-1068-AB91-08002B27B3D9}", 7): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Document.TotalEditingTime", "{F29F85E0-4FF9-1068-AB91-08002B27B3D9}", 10): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Document.Version", "{D5CDD502-2E9C-101B-9397-08002B2CF9AE}", 29): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Document.WordCount", "{F29F85E0-4FF9-1068-AB91-08002B27B3D9}", 15): i = i + 1
    m_PKeys(i) = NamedPKEY("System.DRM.DatePlayExpires", "{AEAC19E4-89AE-4508-B9B7-BB867ABEE2ED}", 6): i = i + 1
    m_PKeys(i) = NamedPKEY("System.DRM.DatePlayStarts", "{AEAC19E4-89AE-4508-B9B7-BB867ABEE2ED}", 5): i = i + 1
    m_PKeys(i) = NamedPKEY("System.DRM.Description", "{AEAC19E4-89AE-4508-B9B7-BB867ABEE2ED}", 3): i = i + 1
    m_PKeys(i) = NamedPKEY("System.DRM.IsDisabled", "{AEAC19E4-89AE-4508-B9B7-BB867ABEE2ED}", 7): i = i + 1
    m_PKeys(i) = NamedPKEY("System.DRM.IsProtected", "{AEAC19E4-89AE-4508-B9B7-BB867ABEE2ED}", 2): i = i + 1
    m_PKeys(i) = NamedPKEY("System.DRM.PlayCount", "{AEAC19E4-89AE-4508-B9B7-BB867ABEE2ED}", 4): i = i + 1
    m_PKeys(i) = NamedPKEY("System.GPS.Altitude", "{827EDB4F-5B73-44A7-891D-FDFFABEA35CA}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.GPS.AltitudeDenominator", "{78342DCB-E358-4145-AE9A-6BFE4E0F9F51}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.GPS.AltitudeNumerator", "{2DAD1EB7-816D-40D3-9EC3-C9773BE2AADE}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.GPS.AltitudeRef", "{46AC629D-75EA-4515-867F-6DC4321C5844}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.GPS.AreaInformation", "{972E333E-AC7E-49F1-8ADF-A70D07A9BCAB}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.GPS.Date", "{3602C812-0F3B-45F0-85AD-603468D69423}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.GPS.DestBearing", "{C66D4B3C-E888-47CC-B99F-9DCA3EE34DEA}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.GPS.DestBearingDenominator", "{7ABCF4F8-7C3F-4988-AC91-8D2C2E97ECA5}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.GPS.DestBearingNumerator", "{BA3B1DA9-86EE-4B5D-A2A4-A271A429F0CF}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.GPS.DestBearingRef", "{9AB84393-2A0F-4B75-BB22-7279786977CB}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.GPS.DestDistance", "{A93EAE04-6804-4F24-AC81-09B266452118}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.GPS.DestDistanceDenominator", "{9BC2C99B-AC71-4127-9D1C-2596D0D7DCB7}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.GPS.DestDistanceNumerator", "{2BDA47DA-08C6-4FE1-80BC-A72FC517C5D0}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.GPS.DestDistanceRef", "{ED4DF2D3-8695-450B-856F-F5C1C53ACB66}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.GPS.DestLatitude", "{9D1D7CC5-5C39-451C-86B3-928E2D18CC47}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.GPS.DestLatitudeDenominator", "{3A372292-7FCA-49A7-99D5-E47BB2D4E7AB}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.GPS.DestLatitudeNumerator", "{ECF4B6F6-D5A6-433C-BB92-4076650FC890}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.GPS.DestLatitudeRef", "{CEA820B9-CE61-4885-A128-005D9087C192}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.GPS.DestLongitude", "{47A96261-CB4C-4807-8AD3-40B9D9DBC6BC}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.GPS.DestLongitudeDenominator", "{425D69E5-48AD-4900-8D80-6EB6B8D0AC86}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.GPS.DestLongitudeNumerator", "{A3250282-FB6D-48D5-9A89-DBCACE75CCCF}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.GPS.DestLongitudeRef", "{182C1EA6-7C1C-4083-AB4B-AC6C9F4ED128}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.GPS.Differential", "{AAF4EE25-BD3B-4DD7-BFC4-47F77BB00F6D}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.GPS.DOP", "{0CF8FB02-1837-42F1-A697-A7017AA289B9}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.GPS.DOPDenominator", "{A0BE94C5-50BA-487B-BD35-0654BE8881ED}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.GPS.DOPNumerator", "{47166B16-364F-4AA0-9F31-E2AB3DF449C3}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.GPS.ImgDirection", "{16473C91-D017-4ED9-BA4D-B6BAA55DBCF8}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.GPS.ImgDirectionDenominator", "{10B24595-41A2-4E20-93C2-5761C1395F32}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.GPS.ImgDirectionNumerator", "{DC5877C7-225F-45F7-BAC7-E81334B6130A}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.GPS.ImgDirectionRef", "{A4AAA5B7-1AD0-445F-811A-0F8F6E67F6B5}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.GPS.Latitude", "{8727CFFF-4868-4EC6-AD5B-81B98521D1AB}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.GPS.LatitudeDecimal", "{0F55CDE2-4F49-450D-92C1-DCD16301B1B7}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.GPS.LatitudeDenominator", "{16E634EE-2BFF-497B-BD8A-4341AD39EEB9}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.GPS.LatitudeNumerator", "{7DDAAAD1-CCC8-41AE-B750-B2CB8031AEA2}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.GPS.LatitudeRef", "{029C0252-5B86-46C7-ACA0-2769FFC8E3D4}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.GPS.Longitude", "{C4C4DBB2-B593-466B-BBDA-D03D27D5E43A}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.GPS.LongitudeDecimal", "{4679C1B5-844D-4590-BAF5-F322231F1B81}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.GPS.LongitudeDenominator", "{BE6E176C-4534-4D2C-ACE5-31DEDAC1606B}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.GPS.LongitudeNumerator", "{02B0F689-A914-4E45-821D-1DDA452ED2C4}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.GPS.LongitudeRef", "{33DCF22B-28D5-464C-8035-1EE9EFD25278}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.GPS.MapDatum", "{2CA2DAE6-EDDC-407D-BEF1-773942ABFA95}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.GPS.MeasureMode", "{A015ED5D-AAEA-4D58-8A86-3C586920EA0B}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.GPS.ProcessingMethod", "{59D49E61-840F-4AA9-A939-E2099B7F6399}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.GPS.Satellites", "{467EE575-1F25-4557-AD4E-B8B58B0D9C15}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.GPS.Speed", "{DA5D0862-6E76-4E1B-BABD-70021BD25494}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.GPS.SpeedDenominator", "{7D122D5A-AE5E-4335-8841-D71E7CE72F53}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.GPS.SpeedNumerator", "{ACC9CE3D-C213-4942-8B48-6D0820F21C6D}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.GPS.SpeedRef", "{ECF7F4C9-544F-4D6D-9D98-8AD79ADAF453}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.GPS.Status", "{125491F4-818F-46B2-91B5-D537753617B2}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.GPS.Track", "{76C09943-7C33-49E3-9E7E-CDBA872CFADA}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.GPS.TrackDenominator", "{C8D1920C-01F6-40C0-AC86-2F3A4AD00770}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.GPS.TrackNumerator", "{702926F4-44A6-43E1-AE71-45627116893B}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.GPS.TrackRef", "{35DBE6FE-44C3-4400-AAAE-D2C799C407E8}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.GPS.VersionID", "{22704DA4-C6B2-4A99-8E56-F16DF8C92599}", 100): i = i + 1
    ' end part2
    Init3 i
End Sub

Private Sub Init3(ByVal i As Long)
    m_PKeys(i) = NamedPKEY("System.History.VisitCount", "{5CBF2787-48CF-4208-B90E-EE5E5D420294}", 7): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Image.BitDepth", "{6444048F-4C8B-11D1-8B70-080036B11A03}", 7): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Image.ColorSpace", "{14B81DA1-0135-4D31-96D9-6CBFC9671A99}", 40961): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Image.CompressedBitsPerPixel", "{364B6FA9-37AB-482A-BE2B-AE02F60D4318}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Image.CompressedBitsPerPixelDenominator", "{1F8844E1-24AD-4508-9DFD-5326A415CE02}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Image.CompressedBitsPerPixelNumerator", "{D21A7148-D32C-4624-8900-277210F79C0F}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Image.Compression", "{14B81DA1-0135-4D31-96D9-6CBFC9671A99}", 259): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Image.CompressionText", "{3F08E66F-2F44-4BB9-A682-AC35D2562322}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Image.Dimensions", "{6444048F-4C8B-11D1-8B70-080036B11A03}", 13): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Image.HorizontalResolution", "{6444048F-4C8B-11D1-8B70-080036B11A03}", 5): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Image.HorizontalSize", "{6444048F-4C8B-11D1-8B70-080036B11A03}", 3): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Image.ImageID", "{10DABE05-32AA-4C29-BF1A-63E2D220587F}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Image.ResolutionUnit", "{19B51FA6-1F92-4A5C-AB48-7DF0ABD67444}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Image.VerticalResolution", "{6444048F-4C8B-11D1-8B70-080036B11A03}", 6): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Image.VerticalSize", "{6444048F-4C8B-11D1-8B70-080036B11A03}", 4): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Journal.Contacts", "{DEA7C82C-1D89-4A66-9427-A4E3DEBABCB1}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Journal.EntryType", "{95BEB1FC-326D-4644-B396-CD3ED90E6DDF}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.LayoutPattern.ContentViewModeForBrowse", "{C9944A21-A406-48FE-8225-AEC7E24C211B}", 500): i = i + 1
    m_PKeys(i) = NamedPKEY("System.LayoutPattern.ContentViewModeForSearch", "{C9944A21-A406-48FE-8225-AEC7E24C211B}", 501): i = i + 1
    m_PKeys(i) = NamedPKEY("System.History.SelectionCount", "{1CE0D6BC-536C-4600-B0DD-7E0C66B350D5}", 8): i = i + 1
    m_PKeys(i) = NamedPKEY("System.History.TargetUrlHostName", "{1CE0D6BC-536C-4600-B0DD-7E0C66B350D5}", 9): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Link.Arguments", "{436F2667-14E2-4FEB-B30A-146C53B5B674}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Link.Comment", "{B9B4B3FC-2B51-4A42-B5D8-324146AFCF25}", 5): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Link.DateVisited", "{5CBF2787-48CF-4208-B90E-EE5E5D420294}", 23): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Link.Description", "{5CBF2787-48CF-4208-B90E-EE5E5D420294}", 21): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Link.FeedItemLocalId", "{8A2F99F9-3C37-465D-A8D7-69777A246D0C}", 2): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Link.Status", "{B9B4B3FC-2B51-4A42-B5D8-324146AFCF25}", 3): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Link.TargetExtension", "{7A7D76F4-B630-4BD7-95FF-37CC51A975C9}", 2): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Link.TargetParsingPath", "{B9B4B3FC-2B51-4A42-B5D8-324146AFCF25}", 2): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Link.TargetSFGAOFlags", "{B9B4B3FC-2B51-4A42-B5D8-324146AFCF25}", 8): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Link.TargetUrlHostName", "{8A2F99F9-3C37-465D-A8D7-69777A246D0C}", 5): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Link.TargetUrlPath", "{8A2F99F9-3C37-465D-A8D7-69777A246D0C}", 6): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Media.AuthorUrl", "{64440492-4C8B-11D1-8B70-080036B11A03}", 32): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Media.AverageLevel", "{09EDD5B6-B301-43C5-9990-D00302EFFD46}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Media.ClassPrimaryID", "{64440492-4C8B-11D1-8B70-080036B11A03}", 13): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Media.ClassSecondaryID", "{64440492-4C8B-11D1-8B70-080036B11A03}", 14): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Media.CollectionGroupID", "{64440492-4C8B-11D1-8B70-080036B11A03}", 24): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Media.CollectionID", "{64440492-4C8B-11D1-8B70-080036B11A03}", 25): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Media.ContentDistributor", "{64440492-4C8B-11D1-8B70-080036B11A03}", 18): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Media.ContentID", "{64440492-4C8B-11D1-8B70-080036B11A03}", 26): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Media.CreatorApplication", "{64440492-4C8B-11D1-8B70-080036B11A03}", 27): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Media.CreatorApplicationVersion", "{64440492-4C8B-11D1-8B70-080036B11A03}", 28): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Media.DateEncoded", "{2E4B640D-5019-46D8-8881-55414CC5CAA0}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Media.DateReleased", "{DE41CC29-6971-4290-B472-F59F2E2F31E2}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Media.DlnaProfileID", "{CFA31B45-525D-4998-BB44-3F7D81542FA4}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Media.Duration", "{64440490-4C8B-11D1-8B70-080036B11A03}", 3): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Media.DVDID", "{64440492-4C8B-11D1-8B70-080036B11A03}", 15): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Media.EncodedBy", "{64440492-4C8B-11D1-8B70-080036B11A03}", 36): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Media.EncodingSettings", "{64440492-4C8B-11D1-8B70-080036B11A03}", 37): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Media.EpisodeNumber", "{64440492-4C8B-11D1-8B70-080036B11A03}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Media.FrameCount", "{6444048F-4C8B-11D1-8B70-080036B11A03}", 12): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Media.MCDI", "{64440492-4C8B-11D1-8B70-080036B11A03}", 16): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Media.MetadataContentProvider", "{64440492-4C8B-11D1-8B70-080036B11A03}", 17): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Media.Producer", "{64440492-4C8B-11D1-8B70-080036B11A03}", 22): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Media.PromotionUrl", "{64440492-4C8B-11D1-8B70-080036B11A03}", 33): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Media.ProtectionType", "{64440492-4C8B-11D1-8B70-080036B11A03}", 38): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Media.ProviderRating", "{64440492-4C8B-11D1-8B70-080036B11A03}", 39): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Media.ProviderStyle", "{64440492-4C8B-11D1-8B70-080036B11A03}", 40): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Media.Publisher", "{64440492-4C8B-11D1-8B70-080036B11A03}", 30): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Media.SeasonNumber", "{64440492-4C8B-11D1-8B70-080036B11A03}", 101): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Media.SeriesName", "{64440492-4C8B-11D1-8B70-080036B11A03}", 42): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Media.SubscriptionContentId", "{9AEBAE7A-9644-487D-A92C-657585ED751A}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Media.SubTitle", "{56A3372E-CE9C-11D2-9F0E-006097C686F6}", 38): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Media.ThumbnailLargePath", "{64440492-4C8B-11D1-8B70-080036B11A03}", 47): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Media.ThumbnailLargeUri", "{64440492-4C8B-11D1-8B70-080036B11A03}", 48): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Media.ThumbnailSmallPath", "{64440492-4C8B-11D1-8B70-080036B11A03}", 49): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Media.ThumbnailSmallUri", "{64440492-4C8B-11D1-8B70-080036B11A03}", 50): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Media.UniqueFileIdentifier", "{64440492-4C8B-11D1-8B70-080036B11A03}", 35): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Media.UserNoAutoInfo", "{64440492-4C8B-11D1-8B70-080036B11A03}", 41): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Media.UserWebUrl", "{64440492-4C8B-11D1-8B70-080036B11A03}", 34): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Media.Writer", "{64440492-4C8B-11D1-8B70-080036B11A03}", 23): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Media.Year", "{56A3372E-CE9C-11D2-9F0E-006097C686F6}", 5): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Message.AttachmentContents", "{3143BF7C-80A8-4854-8880-E2E40189BDD0}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Message.AttachmentNames", "{E3E0584C-B788-4A5A-BB20-7F5A44C9ACDD}", 21): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Message.BccAddress", "{E3E0584C-B788-4A5A-BB20-7F5A44C9ACDD}", 2): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Message.BccName", "{E3E0584C-B788-4A5A-BB20-7F5A44C9ACDD}", 3): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Message.CcAddress", "{E3E0584C-B788-4A5A-BB20-7F5A44C9ACDD}", 4): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Message.CcName", "{E3E0584C-B788-4A5A-BB20-7F5A44C9ACDD}", 5): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Message.ConversationID", "{DC8F80BD-AF1E-4289-85B6-3DFC1B493992}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Message.ConversationIndex", "{DC8F80BD-AF1E-4289-85B6-3DFC1B493992}", 101): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Message.DateReceived", "{E3E0584C-B788-4A5A-BB20-7F5A44C9ACDD}", 20): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Message.DateSent", "{E3E0584C-B788-4A5A-BB20-7F5A44C9ACDD}", 19): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Message.Flags", "{A82D9EE7-CA67-4312-965E-226BCEA85023}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Message.FromAddress", "{E3E0584C-B788-4A5A-BB20-7F5A44C9ACDD}", 13): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Message.FromName", "{E3E0584C-B788-4A5A-BB20-7F5A44C9ACDD}", 14): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Message.HasAttachments", "{9C1FCF74-2D97-41BA-B4AE-CB2E3661A6E4}", 8): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Message.IsFwdOrReply", "{9A9BC088-4F6D-469E-9919-E705412040F9}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Message.MessageClass", "{CD9ED458-08CE-418F-A70E-F912C7BB9C5C}", 103): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Message.Participants", "{1A9BA605-8E7C-4D11-AD7D-A50ADA18BA1B}", 2): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Message.ProofInProgress", "{9098F33C-9A7D-48A8-8DE5-2E1227A64E91}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Message.SenderAddress", "{0BE1C8E7-1981-4676-AE14-FDD78F05A6E7}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Message.SenderName", "{0DA41CFA-D224-4A18-AE2F-596158DB4B3A}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Message.Store", "{E3E0584C-B788-4A5A-BB20-7F5A44C9ACDD}", 15): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Message.ToAddress", "{E3E0584C-B788-4A5A-BB20-7F5A44C9ACDD}", 16): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Message.ToDoFlags", "{1F856A9F-6900-4ABA-9505-2D5F1B4D66CB}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Message.ToDoTitle", "{BCCC8A3C-8CEF-42E5-9B1C-C69079398BC7}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Message.ToName", "{E3E0584C-B788-4A5A-BB20-7F5A44C9ACDD}", 17): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Music.AlbumArtist", "{56A3372E-CE9C-11D2-9F0E-006097C686F6}", 13): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Music.AlbumArtistSortOverride", "{F1FDB4AF-F78C-466C-BB05-56E92DB0B8EC}", 103): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Music.AlbumID", "{56A3372E-CE9C-11D2-9F0E-006097C686F6}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Music.AlbumTitle", "{56A3372E-CE9C-11D2-9F0E-006097C686F6}", 4): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Music.AlbumTitleSortOverride", "{13EB7FFC-EC89-4346-B19D-CCC6F1784223}", 101): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Music.Artist", "{56A3372E-CE9C-11D2-9F0E-006097C686F6}", 2): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Music.ArtistSortOverride", "{DEEB2DB5-0696-4CE0-94FE-A01F77A45FB5}", 102): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Music.BeatsPerMinute", "{56A3372E-CE9C-11D2-9F0E-006097C686F6}", 35): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Music.Composer", "{64440492-4C8B-11D1-8B70-080036B11A03}", 19): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Music.ComposerSortOverride", "{00BC20A3-BD48-4085-872C-A88D77F5097E}", 105): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Music.Conductor", "{56A3372E-CE9C-11D2-9F0E-006097C686F6}", 36): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Music.ContentGroupDescription", "{56A3372E-CE9C-11D2-9F0E-006097C686F6}", 33): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Music.DiscNumber", "{6AFE7437-9BCD-49C7-80FE-4A5C65FA5874}", 104): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Music.DisplayArtist", "{FD122953-FA93-4EF7-92C3-04C946B2F7C8}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Music.Genre", "{56A3372E-CE9C-11D2-9F0E-006097C686F6}", 11): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Music.InitialKey", "{56A3372E-CE9C-11D2-9F0E-006097C686F6}", 34): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Music.IsCompilation", "{C449D5CB-9EA4-4809-82E8-AF9D59DED6D1}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Music.Lyrics", "{56A3372E-CE9C-11D2-9F0E-006097C686F6}", 12): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Music.Mood", "{56A3372E-CE9C-11D2-9F0E-006097C686F6}", 39): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Music.PartOfSet", "{56A3372E-CE9C-11D2-9F0E-006097C686F6}", 37): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Music.Period", "{64440492-4C8B-11D1-8B70-080036B11A03}", 31): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Music.SynchronizedLyrics", "{6B223B6A-162E-4AA9-B39F-05D678FC6D77}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Music.TrackNumber", "{56A3372E-CE9C-11D2-9F0E-006097C686F6}", 7): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Note.Color", "{4776CAFA-BCE4-4CB1-A23E-265E76D8EB11}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Note.ColorText", "{46B4E8DE-CDB2-440D-885C-1658EB65B914}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Photo.Aperture", "{14B81DA1-0135-4D31-96D9-6CBFC9671A99}", 37378): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Photo.ApertureDenominator", "{E1A9A38B-6685-46BD-875E-570DC7AD7320}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Photo.ApertureNumerator", "{0337ECEC-39FB-4581-A0BD-4C4CC51E9914}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Photo.Brightness", "{1A701BF6-478C-4361-83AB-3701BB053C58}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Photo.BrightnessDenominator", "{6EBE6946-2321-440A-90F0-C043EFD32476}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Photo.BrightnessNumerator", "{9E7D118F-B314-45A0-8CFB-D654B917C9E9}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Photo.CameraManufacturer", "{14B81DA1-0135-4D31-96D9-6CBFC9671A99}", 271): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Photo.CameraModel", "{14B81DA1-0135-4D31-96D9-6CBFC9671A99}", 272): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Photo.CameraSerialNumber", "{14B81DA1-0135-4D31-96D9-6CBFC9671A99}", 273): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Photo.Contrast", "{2A785BA9-8D23-4DED-82E6-60A350C86A10}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Photo.ContrastText", "{59DDE9F2-5253-40EA-9A8B-479E96C6249A}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Photo.DateTaken", "{14B81DA1-0135-4D31-96D9-6CBFC9671A99}", 36867): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Photo.DigitalZoom", "{F85BF840-A925-4BC2-B0C4-8E36B598679E}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Photo.DigitalZoomDenominator", "{745BAF0E-E5C1-4CFB-8A1B-D031A0A52393}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Photo.DigitalZoomNumerator", "{16CBB924-6500-473B-A5BE-F1599BCBE413}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Photo.Event", "{14B81DA1-0135-4D31-96D9-6CBFC9671A99}", 18248): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Photo.EXIFVersion", "{D35F743A-EB2E-47F2-A286-844132CB1427}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Photo.ExposureBias", "{14B81DA1-0135-4D31-96D9-6CBFC9671A99}", 37380): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Photo.ExposureBiasDenominator", "{AB205E50-04B7-461C-A18C-2F233836E627}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Photo.ExposureBiasNumerator", "{738BF284-1D87-420B-92CF-5834BF6EF9ED}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Photo.ExposureIndex", "{967B5AF8-995A-46ED-9E11-35B3C5B9782D}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Photo.ExposureIndexDenominator", "{93112F89-C28B-492F-8A9D-4BE2062CEE8A}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Photo.ExposureIndexNumerator", "{CDEDCF30-8919-44DF-8F4C-4EB2FFDB8D89}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Photo.ExposureProgram", "{14B81DA1-0135-4D31-96D9-6CBFC9671A99}", 34850): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Photo.ExposureProgramText", "{FEC690B7-5F30-4646-AE47-4CAAFBA884A3}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Photo.ExposureTime", "{14B81DA1-0135-4D31-96D9-6CBFC9671A99}", 33434): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Photo.ExposureTimeDenominator", "{55E98597-AD16-42E0-B624-21599A199838}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Photo.ExposureTimeNumerator", "{257E44E2-9031-4323-AC38-85C552871B2E}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Photo.Flash", "{14B81DA1-0135-4D31-96D9-6CBFC9671A99}", 37385): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Photo.FlashEnergy", "{14B81DA1-0135-4D31-96D9-6CBFC9671A99}", 41483): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Photo.FlashEnergyDenominator", "{D7B61C70-6323-49CD-A5FC-C84277162C97}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Photo.FlashEnergyNumerator", "{FCAD3D3D-0858-400F-AAA3-2F66CCE2A6BC}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Photo.FlashManufacturer", "{AABAF6C9-E0C5-4719-8585-57B103E584FE}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Photo.FlashModel", "{FE83BB35-4D1A-42E2-916B-06F3E1AF719E}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Photo.FlashText", "{6B8B68F6-200B-47EA-8D25-D8050F57339F}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Photo.FNumber", "{14B81DA1-0135-4D31-96D9-6CBFC9671A99}", 33437): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Photo.FNumberDenominator", "{E92A2496-223B-4463-A4E3-30EABBA79D80}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Photo.FNumberNumerator", "{1B97738A-FDFC-462F-9D93-1957E08BE90C}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Photo.FocalLength", "{14B81DA1-0135-4D31-96D9-6CBFC9671A99}", 37386): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Photo.FocalLengthDenominator", "{305BC615-DCA1-44A5-9FD4-10C0BA79412E}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Photo.FocalLengthInFilm", "{A0E74609-B84D-4F49-B860-462BD9971F98}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Photo.FocalLengthNumerator", "{776B6B3B-1E3D-4B0C-9A0E-8FBAF2A8492A}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Photo.FocalPlaneXResolution", "{CFC08D97-C6F7-4484-89DD-EBEF4356FE76}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Photo.FocalPlaneXResolutionDenominator", "{0933F3F5-4786-4F46-A8E8-D64DD37FA521}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Photo.FocalPlaneXResolutionNumerator", "{DCCB10AF-B4E2-4B88-95F9-031B4D5AB490}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Photo.FocalPlaneYResolution", "{4FFFE4D0-914F-4AC4-8D6F-C9C61DE169B1}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Photo.FocalPlaneYResolutionDenominator", "{1D6179A6-A876-4031-B013-3347B2B64DC8}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Photo.FocalPlaneYResolutionNumerator", "{A2E541C5-4440-4BA8-867E-75CFC06828CD}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Photo.GainControl", "{FA304789-00C7-4D80-904A-1E4DCC7265AA}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Photo.GainControlDenominator", "{42864DFD-9DA4-4F77-BDED-4AAD7B256735}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Photo.GainControlNumerator", "{8E8ECF7C-B7B8-4EB8-A63F-0EE715C96F9E}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Photo.GainControlText", "{C06238B2-0BF9-4279-A723-25856715CB9D}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Photo.ISOSpeed", "{14B81DA1-0135-4D31-96D9-6CBFC9671A99}", 34855): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Photo.LensManufacturer", "{E6DDCAF7-29C5-4F0A-9A68-D19412EC7090}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Photo.LensModel", "{E1277516-2B5F-4869-89B1-2E585BD38B7A}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Photo.LightSource", "{14B81DA1-0135-4D31-96D9-6CBFC9671A99}", 37384): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Photo.MakerNote", "{FA303353-B659-4052-85E9-BCAC79549B84}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Photo.MakerNoteOffset", "{813F4124-34E6-4D17-AB3E-6B1F3C2247A1}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Photo.MaxAperture", "{08F6D7C2-E3F2-44FC-AF1E-5AA5C81A2D3E}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Photo.MaxApertureDenominator", "{C77724D4-601F-46C5-9B89-C53F93BCEB77}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Photo.MaxApertureNumerator", "{C107E191-A459-44C5-9AE6-B952AD4B906D}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Photo.MeteringMode", "{14B81DA1-0135-4D31-96D9-6CBFC9671A99}", 37383): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Photo.MeteringModeText", "{F628FD8C-7BA8-465A-A65B-C5AA79263A9E}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Photo.Orientation", "{14B81DA1-0135-4D31-96D9-6CBFC9671A99}", 274): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Photo.OrientationText", "{A9EA193C-C511-498A-A06B-58E2776DCC28}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Photo.PeopleNames", "{E8309B6E-084C-49B4-B1FC-90A80331B638}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Photo.PhotometricInterpretation", "{341796F1-1DF9-4B1C-A564-91BDEFA43877}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Photo.PhotometricInterpretationText", "{821437D6-9EAB-4765-A589-3B1CBBD22A61}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Photo.ProgramMode", "{6D217F6D-3F6A-4825-B470-5F03CA2FBE9B}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Photo.ProgramModeText", "{7FE3AA27-2648-42F3-89B0-454E5CB150C3}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Photo.RelatedSoundFile", "{318A6B45-087F-4DC2-B8CC-05359551FC9E}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Photo.Saturation", "{49237325-A95A-4F67-B211-816B2D45D2E0}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Photo.SaturationText", "{61478C08-B600-4A84-BBE4-E99C45F0A072}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Photo.Sharpness", "{FC6976DB-8349-4970-AE97-B3C5316A08F0}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Photo.SharpnessText", "{51EC3F47-DD50-421D-8769-334F50424B1E}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Photo.ShutterSpeed", "{14B81DA1-0135-4D31-96D9-6CBFC9671A99}", 37377): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Photo.ShutterSpeedDenominator", "{E13D8975-81C7-4948-AE3F-37CAE11E8FF7}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Photo.ShutterSpeedNumerator", "{16EA4042-D6F4-4BCA-8349-7C78D30FB333}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Photo.SubjectDistance", "{14B81DA1-0135-4D31-96D9-6CBFC9671A99}", 37382): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Photo.SubjectDistanceDenominator", "{0C840A88-B043-466D-9766-D4B26DA3FA77}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Photo.SubjectDistanceNumerator", "{8AF4961C-F526-43E5-AA81-DB768219178D}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Photo.TagViewAggregate", "{B812F15D-C2D8-4BBF-BACD-79744346113F}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Photo.TranscodedForSync", "{9A8EBB75-6458-4E82-BACB-35C0095B03BB}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Photo.WhiteBalance", "{EE3D3D8A-5381-4CFA-B13B-AAF66B5F4EC9}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Photo.WhiteBalanceText", "{6336B95E-C7A7-426D-86FD-7AE3D39C84B4}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.PropGroup.Advanced", "{900A403B-097B-4B95-8AE2-071FDAEEB118}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.PropGroup.Audio", "{2804D469-788F-48AA-8570-71B9C187E138}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.PropGroup.Calendar", "{9973D2B5-BFD8-438A-BA94-5349B293181A}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.PropGroup.Camera", "{DE00DE32-547E-4981-AD4B-542F2E9007D8}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.PropGroup.Contact", "{DF975FD3-250A-4004-858F-34E29A3E37AA}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.PropGroup.Content", "{D0DAB0BA-368A-4050-A882-6C010FD19A4F}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.PropGroup.Description", "{8969B275-9475-4E00-A887-FF93B8B41E44}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.PropGroup.FileSystem", "{E3A7D2C1-80FC-4B40-8F34-30EA111BDC2E}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.PropGroup.General", "{CC301630-B192-4C22-B372-9F4C6D338E07}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.PropGroup.GPS", "{F3713ADA-90E3-4E11-AAE5-FDC17685B9BE}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.PropGroup.Image", "{E3690A87-0FA8-4A2A-9A9F-FCE8827055AC}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.PropGroup.Media", "{61872CF7-6B5E-4B4B-AC2D-59DA84459248}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.PropGroup.MediaAdvanced", "{8859A284-DE7E-4642-99BA-D431D044B1EC}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.PropGroup.Message", "{7FD7259D-16B4-4135-9F97-7C96ECD2FA9E}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.PropGroup.Music", "{68DD6094-7216-40F1-A029-43FE7127043F}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.PropGroup.Origin", "{2598D2FB-5569-4367-95DF-5CD3A177E1A5}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.PropGroup.PhotoAdvanced", "{0CB2BF5A-9EE7-4A86-8222-F01E07FDADAF}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.PropGroup.RecordedTV", "{E7B33238-6584-4170-A5C0-AC25EFD9DA56}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.PropGroup.Video", "{BEBE0920-7671-4C54-A3EB-49FDDFC191EE}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.InfoTipText", "{C9944A21-A406-48FE-8225-AEC7E24C211B}", 17): i = i + 1
    m_PKeys(i) = NamedPKEY("System.PropList.ConflictPrompt", "{C9944A21-A406-48FE-8225-AEC7E24C211B}", 11): i = i + 1
    m_PKeys(i) = NamedPKEY("System.PropList.ContentViewModeForBrowse", "{C9944A21-A406-48FE-8225-AEC7E24C211B}", 13): i = i + 1
    m_PKeys(i) = NamedPKEY("System.PropList.ContentViewModeForSearch", "{C9944A21-A406-48FE-8225-AEC7E24C211B}", 14): i = i + 1
    m_PKeys(i) = NamedPKEY("System.PropList.ExtendedTileInfo", "{C9944A21-A406-48FE-8225-AEC7E24C211B}", 9): i = i + 1
    m_PKeys(i) = NamedPKEY("System.PropList.FileOperationPrompt", "{C9944A21-A406-48FE-8225-AEC7E24C211B}", 10): i = i + 1
    m_PKeys(i) = NamedPKEY("System.PropList.FullDetails", "{C9944A21-A406-48FE-8225-AEC7E24C211B}", 2): i = i + 1
    m_PKeys(i) = NamedPKEY("System.PropList.InfoTip", "{C9944A21-A406-48FE-8225-AEC7E24C211B}", 4): i = i + 1
    m_PKeys(i) = NamedPKEY("System.PropList.NonPersonal", "{49D1091F-082E-493F-B23F-D2308AA9668C}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.PropList.PreviewDetails", "{C9944A21-A406-48FE-8225-AEC7E24C211B}", 8): i = i + 1
    m_PKeys(i) = NamedPKEY("System.PropList.PreviewTitle", "{C9944A21-A406-48FE-8225-AEC7E24C211B}", 6): i = i + 1
    m_PKeys(i) = NamedPKEY("System.PropList.QuickTip", "{C9944A21-A406-48FE-8225-AEC7E24C211B}", 5): i = i + 1
    m_PKeys(i) = NamedPKEY("System.PropList.TileInfo", "{C9944A21-A406-48FE-8225-AEC7E24C211B}", 3): i = i + 1
    m_PKeys(i) = NamedPKEY("System.PropList.XPDetailsPanel", "{F2275480-F782-4291-BD94-F13693513AEC}", 0): i = i + 1
    m_PKeys(i) = NamedPKEY("System.RecordedTV.ChannelNumber", "{6D748DE2-8D38-4CC3-AC60-F009B057C557}", 7): i = i + 1
    m_PKeys(i) = NamedPKEY("System.RecordedTV.Credits", "{6D748DE2-8D38-4CC3-AC60-F009B057C557}", 4): i = i + 1
    m_PKeys(i) = NamedPKEY("System.RecordedTV.DateContentExpires", "{6D748DE2-8D38-4CC3-AC60-F009B057C557}", 15): i = i + 1
    m_PKeys(i) = NamedPKEY("System.RecordedTV.EpisodeName", "{6D748DE2-8D38-4CC3-AC60-F009B057C557}", 2): i = i + 1
    m_PKeys(i) = NamedPKEY("System.RecordedTV.IsATSCContent", "{6D748DE2-8D38-4CC3-AC60-F009B057C557}", 16): i = i + 1
    m_PKeys(i) = NamedPKEY("System.RecordedTV.IsClosedCaptioningAvailable", "{6D748DE2-8D38-4CC3-AC60-F009B057C557}", 12): i = i + 1
    m_PKeys(i) = NamedPKEY("System.RecordedTV.IsDTVContent", "{6D748DE2-8D38-4CC3-AC60-F009B057C557}", 17): i = i + 1
    m_PKeys(i) = NamedPKEY("System.RecordedTV.IsHDContent", "{6D748DE2-8D38-4CC3-AC60-F009B057C557}", 18): i = i + 1
    m_PKeys(i) = NamedPKEY("System.RecordedTV.IsRepeatBroadcast", "{6D748DE2-8D38-4CC3-AC60-F009B057C557}", 13): i = i + 1
    m_PKeys(i) = NamedPKEY("System.RecordedTV.IsSAP", "{6D748DE2-8D38-4CC3-AC60-F009B057C557}", 14): i = i + 1
    m_PKeys(i) = NamedPKEY("System.RecordedTV.NetworkAffiliation", "{2C53C813-FB63-4E22-A1AB-0B331CA1E273}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.RecordedTV.OriginalBroadcastDate", "{4684FE97-8765-4842-9C13-F006447B178C}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.RecordedTV.ProgramDescription", "{6D748DE2-8D38-4CC3-AC60-F009B057C557}", 3): i = i + 1
    m_PKeys(i) = NamedPKEY("System.RecordedTV.RecordingTime", "{A5477F61-7A82-4ECA-9DDE-98B69B2479B3}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.RecordedTV.StationCallSign", "{6D748DE2-8D38-4CC3-AC60-F009B057C557}", 5): i = i + 1
    m_PKeys(i) = NamedPKEY("System.RecordedTV.StationName", "{1B5439E7-EBA1-4AF8-BDD7-7AF1D4549493}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Search.AutoSummary", "{560C36C0-503A-11CF-BAA1-00004C752A9A}", 2): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Search.ContainerHash", "{BCEEE283-35DF-4D53-826A-F36A3EEFC6BE}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Search.Contents", "{B725F130-47EF-101A-A5F1-02608C9EEBAC}", 19): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Search.EntryID", "{49691C90-7E17-101A-A91C-08002B2ECDA9}", 5): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Search.ExtendedProperties", "{7B03B546-FA4F-4A52-A2FE-03D5311E5865}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Search.GatherTime", "{0B63E350-9CCC-11D0-BCDB-00805FCCCE04}", 8): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Search.HitCount", "{49691C90-7E17-101A-A91C-08002B2ECDA9}", 4): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Search.IsClosedDirectory", "{0B63E343-9CCC-11D0-BCDB-00805FCCCE04}", 23): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Search.IsFullyContained", "{0B63E343-9CCC-11D0-BCDB-00805FCCCE04}", 24): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Search.QueryFocusedSummary", "{560C36C0-503A-11CF-BAA1-00004C752A9A}", 3): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Search.QueryFocusedSummaryWithFallback", "{560C36C0-503A-11CF-BAA1-00004C752A9A}", 4): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Search.QueryPropertyHits", "{49691C90-7E17-101A-A91C-08002B2ECDA9}", 21): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Search.Rank", "{49691C90-7E17-101A-A91C-08002B2ECDA9}", 3): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Search.Store", "{A06992B3-8CAF-4ED7-A547-B259E32AC9FC}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Search.UrlToIndex", "{0B63E343-9CCC-11D0-BCDB-00805FCCCE04}", 2): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Search.UrlToIndexWithModificationTime", "{0B63E343-9CCC-11D0-BCDB-00805FCCCE04}", 12): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Supplemental.AlbumID", "{0C73B141-39D6-4653-A683-CAB291EAF95B}", 2): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Supplemental.ResourceId", "{0C73B141-39D6-4653-A683-CAB291EAF95B}", 3): i = i + 1
    m_PKeys(i) = NamedPKEY("System.DescriptionID", "{28636AA6-953D-11D2-B5D6-00C04FD918D0}", 2): i = i + 1
    m_PKeys(i) = NamedPKEY("System.InternalName", "{0CEF7D53-FA64-11D1-A203-0000F81FEDEE}", 5): i = i + 1
    m_PKeys(i) = NamedPKEY("System.LibraryLocationsCount", "{908696C7-8F87-44F2-80ED-A8C1C6894575}", 2): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Link.TargetSFGAOFlagsStrings", "{D6942081-D53B-443D-AD47-5E059D9CD27A}", 3): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Link.TargetUrl", "{5CBF2787-48CF-4208-B90E-EE5E5D420294}", 2): i = i + 1
    m_PKeys(i) = NamedPKEY("System.NamespaceCLSID", "{28636AA6-953D-11D2-B5D6-00C04FD918D0}", 6): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Shell.SFGAOFlagsStrings", "{D6942081-D53B-443D-AD47-5E059D9CD27A}", 2): i = i + 1
    m_PKeys(i) = NamedPKEY("System.StatusBarSelectedItemCount", "{26DC287C-6E3D-4BD3-B2B0-6A26BA2E346D}", 3): i = i + 1
    m_PKeys(i) = NamedPKEY("System.StatusBarViewItemCount", "{26DC287C-6E3D-4BD3-B2B0-6A26BA2E346D}", 2): i = i + 1
    m_PKeys(i) = NamedPKEY("System.AppUserModel.ExcludeFromShowInNewInstall", "{9F4C2855-9F79-4B39-A8D0-E1D42DE1D5F3}", 8): i = i + 1
    m_PKeys(i) = NamedPKEY("System.AppUserModel.ID", "{9F4C2855-9F79-4B39-A8D0-E1D42DE1D5F3}", 5): i = i + 1
    m_PKeys(i) = NamedPKEY("System.AppUserModel.IsDestListSeparator", "{9F4C2855-9F79-4B39-A8D0-E1D42DE1D5F3}", 6): i = i + 1
    m_PKeys(i) = NamedPKEY("System.AppUserModel.IsDualMode", "{9F4C2855-9F79-4B39-A8D0-E1D42DE1D5F3}", 11): i = i + 1
    m_PKeys(i) = NamedPKEY("System.AppUserModel.PreventPinning", "{9F4C2855-9F79-4B39-A8D0-E1D42DE1D5F3}", 9): i = i + 1
    m_PKeys(i) = NamedPKEY("System.AppUserModel.RelaunchCommand", "{9F4C2855-9F79-4B39-A8D0-E1D42DE1D5F3}", 2): i = i + 1
    m_PKeys(i) = NamedPKEY("System.AppUserModel.RelaunchDisplayNameResource", "{9F4C2855-9F79-4B39-A8D0-E1D42DE1D5F3}", 4): i = i + 1
    m_PKeys(i) = NamedPKEY("System.AppUserModel.RelaunchIconResource", "{9F4C2855-9F79-4B39-A8D0-E1D42DE1D5F3}", 3): i = i + 1
    m_PKeys(i) = NamedPKEY("System.AppUserModel.StartPinOption", "{9F4C2855-9F79-4B39-A8D0-E1D42DE1D5F3}", 12): i = i + 1
    m_PKeys(i) = NamedPKEY("System.AppUserModel.ToastActivatorCLSID", "{9F4C2855-9F79-4B39-A8D0-E1D42DE1D5F3}", 26): i = i + 1
    m_PKeys(i) = NamedPKEY("System.EdgeGesture.DisableTouchWhenFullscreen", "{32CE38B2-2C9A-41B1-9BC5-B3784394AA44}", 2): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Software.DateLastUsed", "{841E4F90-FF59-4D16-8947-E81BBFFAB36D}", 16): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Software.ProductName", "{0CEF7D53-FA64-11D1-A203-0000F81FEDEE}", 7): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Sync.Comments", "{7BD5533E-AF15-44DB-B8C8-BD6624E1D032}", 13): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Sync.ConflictDescription", "{CE50C159-2FB8-41FD-BE68-D3E042E274BC}", 4): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Sync.ConflictFirstLocation", "{CE50C159-2FB8-41FD-BE68-D3E042E274BC}", 6): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Sync.ConflictSecondLocation", "{CE50C159-2FB8-41FD-BE68-D3E042E274BC}", 7): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Sync.HandlerCollectionID", "{7BD5533E-AF15-44DB-B8C8-BD6624E1D032}", 2): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Sync.HandlerID", "{7BD5533E-AF15-44DB-B8C8-BD6624E1D032}", 3): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Sync.HandlerName", "{CE50C159-2FB8-41FD-BE68-D3E042E274BC}", 2): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Sync.HandlerType", "{7BD5533E-AF15-44DB-B8C8-BD6624E1D032}", 8): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Sync.HandlerTypeLabel", "{7BD5533E-AF15-44DB-B8C8-BD6624E1D032}", 9): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Sync.ItemID", "{7BD5533E-AF15-44DB-B8C8-BD6624E1D032}", 6): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Sync.ItemName", "{CE50C159-2FB8-41FD-BE68-D3E042E274BC}", 3): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Sync.ProgressPercentage", "{7BD5533E-AF15-44DB-B8C8-BD6624E1D032}", 23): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Sync.State", "{7BD5533E-AF15-44DB-B8C8-BD6624E1D032}", 24): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Sync.Status", "{7BD5533E-AF15-44DB-B8C8-BD6624E1D032}", 10): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Task.BillingInformation", "{D37D52C6-261C-4303-82B3-08B926AC6F12}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Task.CompletionStatus", "{084D8A0A-E6D5-40DE-BF1F-C8820E7C877C}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Task.Owner", "{08C7CC5F-60F2-4494-AD75-55E3E0B5ADD0}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Video.Compression", "{64440491-4C8B-11D1-8B70-080036B11A03}", 10): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Video.Director", "{64440492-4C8B-11D1-8B70-080036B11A03}", 20): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Video.EncodingBitrate", "{64440491-4C8B-11D1-8B70-080036B11A03}", 8): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Video.FourCC", "{64440491-4C8B-11D1-8B70-080036B11A03}", 44): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Video.FrameHeight", "{64440491-4C8B-11D1-8B70-080036B11A03}", 4): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Video.FrameRate", "{64440491-4C8B-11D1-8B70-080036B11A03}", 6): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Video.FrameWidth", "{64440491-4C8B-11D1-8B70-080036B11A03}", 3): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Video.HorizontalAspectRatio", "{64440491-4C8B-11D1-8B70-080036B11A03}", 42): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Video.IsSpherical", "{64440491-4C8B-11D1-8B70-080036B11A03}", 100): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Video.IsStereo", "{64440491-4C8B-11D1-8B70-080036B11A03}", 98): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Video.Orientation", "{64440491-4C8B-11D1-8B70-080036B11A03}", 99): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Video.SampleSize", "{64440491-4C8B-11D1-8B70-080036B11A03}", 9): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Video.StreamName", "{64440491-4C8B-11D1-8B70-080036B11A03}", 2): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Video.StreamNumber", "{64440491-4C8B-11D1-8B70-080036B11A03}", 11): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Video.TotalBitrate", "{64440491-4C8B-11D1-8B70-080036B11A03}", 43): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Video.TranscodedForSync", "{64440491-4C8B-11D1-8B70-080036B11A03}", 46): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Video.VerticalAspectRatio", "{64440491-4C8B-11D1-8B70-080036B11A03}", 45): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Volume.FileSystem", "{9B174B35-40FF-11D2-A27E-00C04FC30871}", 4): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Volume.IsMappedDrive", "{149C0B69-2C2D-48FC-808F-D318D78C4636}", 2): i = i + 1
    m_PKeys(i) = NamedPKEY("System.Volume.IsRoot", "{9B174B35-40FF-11D2-A27E-00C04FC30871}", 10): i = i + 1
End Sub

