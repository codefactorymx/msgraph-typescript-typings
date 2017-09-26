// Type definitions for the Microsoft Graph API
// Project: https://github.com/microsoftgraph/msgraph-typescript-typings
// Definitions by: Microsoft Graph Team <https://github.com/microsoftgraph>

//
// Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.  See License in the project root for license information.
//


export as namespace microsoftgraphbeta;

export type Status = "active" | "updated" | "deleted" | "ignored"
export type AutomaticRepliesStatus = "disabled" | "alwaysEnabled" | "scheduled"
export type ExternalAudienceScope = "none" | "contactsOnly" | "all"
export type AttendeeType = "required" | "optional" | "resource"
export type FreeBusyStatus = "free" | "tentative" | "busy" | "oof" | "workingElsewhere" | "unknown"
export type PhysicalAddressType = "unknown" | "home" | "business" | "other"
export type LocationType = "default" | "conferenceRoom" | "homeAddress" | "businessAddress" | "geoCoordinates" | "streetAddress" | "hotel" | "restaurant" | "localBusiness" | "postalAddress"
export type LocationUniqueIdType = "unknown" | "locationStore" | "directory" | "private" | "bing"
export type ActivityDomain = "unknown" | "work" | "personal" | "unrestricted"
export type RecipientScopeType = "none" | "internal" | "external" | "externalPartner" | "externalNonPartner"
export type MailTipsType = "automaticReplies" | "mailboxFullStatus" | "customMailTip" | "externalMemberCount" | "totalMemberCount" | "maxMessageSize" | "deliveryRestriction" | "moderationStatus" | "recipientScope" | "recipientSuggestions"
export type TimeZoneStandard = "windows" | "iana"
export type BodyType = "text" | "html"
export type Importance = "low" | "normal" | "high"
export type InferenceClassificationType = "focused" | "other"
export type FollowupFlagStatus = "notFlagged" | "complete" | "flagged"
export type CalendarColor = "lightBlue" | "lightGreen" | "lightOrange" | "lightGray" | "lightYellow" | "lightTeal" | "lightPink" | "lightBrown" | "lightRed" | "maxColor" | "auto"
export type ResponseType = "none" | "organizer" | "tentativelyAccepted" | "accepted" | "declined" | "notResponded"
export type Sensitivity = "normal" | "personal" | "private" | "confidential"
export type RecurrencePatternType = "daily" | "weekly" | "absoluteMonthly" | "relativeMonthly" | "absoluteYearly" | "relativeYearly"
export type DayOfWeek = "sunday" | "monday" | "tuesday" | "wednesday" | "thursday" | "friday" | "saturday"
export type WeekIndex = "first" | "second" | "third" | "fourth" | "last"
export type RecurrenceRangeType = "endDate" | "noEnd" | "numbered"
export type EventType = "singleInstance" | "occurrence" | "exception" | "seriesMaster"
export type PhoneType = "home" | "business" | "mobile" | "other" | "assistant" | "homeFax" | "businessFax" | "otherFax" | "pager" | "radio"
export type WebsiteType = "other" | "home" | "work" | "blog" | "profile"
export type MeetingMessageType = "none" | "meetingRequest" | "meetingCancelled" | "meetingAccepted" | "meetingTentativelyAccepted" | "meetingDeclined"
export type MessageActionFlag = "any" | "call" | "doNotForward" | "followUp" | "fyi" | "forward" | "noResponseNecessary" | "read" | "reply" | "replyToAll" | "review"
export type ReferenceAttachmentProvider = "other" | "oneDriveBusiness" | "oneDriveConsumer" | "dropbox"
export type ReferenceAttachmentPermission = "other" | "view" | "edit" | "anonymousView" | "anonymousEdit" | "organizationView" | "organizationEdit"
export type GroupAccessType = "none" | "private" | "secret" | "public"
export type CategoryColor = "preset0" | "preset1" | "preset2" | "preset3" | "preset4" | "preset5" | "preset6" | "preset7" | "preset8" | "preset9" | "preset10" | "preset11" | "preset12" | "preset13" | "preset14" | "preset15" | "preset16" | "preset17" | "preset18" | "preset19" | "preset20" | "preset21" | "preset22" | "preset23" | "preset24" | "none"
export type TaskStatus = "notStarted" | "inProgress" | "completed" | "waitingOnOthers" | "deferred"
export type PlannerPreviewType = "automatic" | "noPreview" | "checklist" | "description" | "reference"
export type TaskBoardType = "progress" | "assignedTo" | "bucket"
export type PreviewType = "automatic" | "noPreview" | "checklist" | "description" | "reference"
export type OperationStatus = "NotStarted" | "Running" | "Completed" | "Failed"
export type OnenotePatchInsertPosition = "After" | "Before"
export type OnenotePatchActionType = "Replace" | "Append" | "Delete" | "Insert" | "Prepend"
export type OnenoteSourceService = "Unknown" | "OneDrive" | "OneDriveForBusiness" | "OnPremOneDriveForBusiness"
export type OnenoteUserRole = "Owner" | "Contributor" | "Reader" | "None"
export type RiskEventStatus = "active" | "remediated" | "dismissedAsFixed" | "dismissedAsFalsePositive" | "dismissedAsIgnore" | "loginBlocked" | "closedMfaAuto" | "closedMultipleReasons"
export type RiskLevel = "low" | "medium" | "high"
export type UserRiskLevel = "unknown" | "none" | "low" | "medium" | "high"
export type ApprovalState = "pending" | "approved" | "denied" | "aborted" | "canceled"
export type RoleSummaryStatus = "ok" | "bad"
export type SetupStatus = "unknown" | "notRegisteredYet" | "registeredSetupNotStarted" | "registeredSetupInProgress" | "registrationAndSetupCompleted" | "registrationFailed" | "registrationTimedOut" | "disabled"
export type AndroidForWorkBindStatus = "notBound" | "bound" | "boundAndValidated" | "unbinding"
export type AndroidForWorkSyncStatus = "success" | "credentialsNotValid" | "androidForWorkApiError" | "managementServiceError" | "unknownError" | "none"
export type AndroidForWorkEnrollmentTarget = "none" | "all" | "targeted" | "targetedAsEnrollmentRestrictions"
export type AndroidForWorkAppConfigurationSchemaItemDataType = "bool" | "integer" | "string" | "choice" | "multiselect" | "bundle" | "bundleArray" | "hidden"
export type AppInstallIntent = "available" | "notApplicable" | "required" | "uninstall" | "availableWithoutEnrollment"
export type ResultantAppState = "installed" | "failed" | "notInstalled" | "uninstallFailed" | "unknown" | "notApplicable"
export type OfficeProductId = "o365ProPlusRetail" | "o365BusinessRetail" | "visioProRetail" | "projectProRetail"
export type OfficeUpdateChannel = "none" | "current" | "deferred" | "firstReleaseCurrent" | "firstReleaseDeferred"
export type WindowsArchitecture = "none" | "x86" | "x64" | "arm" | "neutral"
export type ManagedAppAvailability = "global" | "lineOfBusiness"
export type MobileAppContentFileUploadState = "success" | "transientError" | "error" | "unknown" | "azureStorageUriRequestSuccess" | "azureStorageUriRequestPending" | "azureStorageUriRequestFailed" | "azureStorageUriRequestTimedOut" | "azureStorageUriRenewalSuccess" | "azureStorageUriRenewalPending" | "azureStorageUriRenewalFailed" | "azureStorageUriRenewalTimedOut" | "commitFileSuccess" | "commitFilePending" | "commitFileFailed" | "commitFileTimedOut"
export type WindowsDeviceType = "none" | "desktop" | "mobile"
export type VolumePurchaseProgramTokenAccountType = "business" | "education"
export type VolumePurchaseProgramTokenState = "unknown" | "valid" | "expired" | "invalid"
export type VolumePurchaseProgramTokenSyncStatus = "none" | "inProgress" | "completed" | "failed"
export type WindowsStoreForBusinessLicenseType = "offline" | "online"
export type CertificateStatus = "notProvisioned" | "provisioned"
export type ComplianceStatus = "unknown" | "notApplicable" | "compliant" | "remediated" | "nonCompliant" | "error" | "conflict"
export type AndroidPermissionActionType = "prompt" | "autoGrant" | "autoDeny"
export type MdmAppConfigKeyType = "stringType" | "integerType" | "realType" | "booleanType" | "tokenType"
export type ITunesPairingMode = "disallow" | "allow" | "requiresCertificate"
export type ImportedDeviceIdentityType = "unknown" | "imei" | "serialNumber"
export type EnrollmentState = "unknown" | "enrolled" | "pendingReset" | "failed" | "notContacted"
export type Platform = "unknown" | "ios" | "android" | "windows" | "windowsMobile" | "macOS"
export type DiscoverySource = "unknown" | "adminImport" | "deviceEnrollmentProgram"
export type OwnerType = "unknown" | "company" | "personal"
export type DeviceActionState = "none" | "pending" | "canceled" | "active" | "done" | "failed" | "notSupported"
export type ManagementState = "managed" | "retirePending" | "retireFailed" | "wipePending" | "wipeFailed" | "unhealthy" | "deletePending" | "retireIssued" | "wipeIssued" | "wipeCanceled" | "retireCanceled" | "discovered"
export type ChassisType = "unknown" | "desktop" | "laptop" | "worksWorkstation" | "enterpriseServer" | "phone" | "tablet" | "mobileOther" | "mobileUnknown"
export type DeviceType = "desktop" | "windowsRT" | "winMO6" | "nokia" | "windowsPhone" | "mac" | "winCE" | "winEmbedded" | "iPhone" | "iPad" | "iPod" | "android" | "iSocConsumer" | "unix" | "macMDM" | "holoLens" | "surfaceHub" | "androidForWork" | "windowsBlue" | "windowsPhoneBlue" | "blackberry" | "palm" | "fakeDevice" | "unknown"
export type ComplianceState = "unknown" | "compliant" | "noncompliant" | "conflict" | "error"
export type ManagementAgentType = "eas" | "mdm" | "easMdm" | "intuneClient" | "easIntuneClient" | "configManagerClient" | "configurationManagerClientMdmEas" | "unknown"
export type EnrollmentType = "unknown" | "userEnrollment" | "deviceEnrollment" | "deviceEnrollmentWithUDA" | "azureDomainJoined" | "userEnrollmentWithServiceAccount" | "depDeviceEnrollment" | "depDeviceEnrollmentWithUDA" | "autoEnrollment"
export type LostModeState = "disabled" | "enabled"
export type DeviceRegistrationState = "notRegistered" | "smsidConflict" | "registered" | "revoked" | "keyConflict" | "approvalPending" | "resetCert" | "notRegisteredPendingEnrollment" | "unknown"
export type DeviceManagementExchangeAccessState = "none" | "unknown" | "allowed" | "blocked" | "quarantined"
export type DeviceManagementExchangeAccessStateReason = "none" | "unknown" | "exchangeGlobalRule" | "exchangeIndividualRule" | "exchangeDeviceRule" | "exchangeUpgrade" | "exchangeMailboxPolicy" | "other" | "compliant" | "notCompliant" | "notEnrolled" | "unknownLocation" | "mfaRequired" | "azureADBlockDueToAccessPolicy" | "compromisedPassword" | "deviceNotKnownWithManagedApp"
export type WindowsDeviceHealthState = "clean" | "fullScanPending" | "rebootPending" | "manualStepsPending" | "offlineScanPending" | "critical"
export type WindowsMalwareSeverity = "unknown" | "low" | "moderate" | "high" | "severe"
export type WindowsMalwareCategory = "invalid" | "adware" | "spyware" | "passwordStealer" | "trojanDownloader" | "worm" | "backdoor" | "remoteAccessTrojan" | "trojan" | "emailFlooder" | "keylogger" | "dialer" | "monitoringSoftware" | "browserModifier" | "cookie" | "browserPlugin" | "aolExploit" | "nuker" | "securityDisabler" | "jokeProgram" | "hostileActiveXControl" | "softwareBundler" | "stealthNotifier" | "settingsModifier" | "toolBar" | "remoteControlSoftware" | "trojanFtp" | "potentialUnwantedSoftware" | "icqExploit" | "trojanTelnet" | "exploit" | "filesharingProgram" | "malwareCreationTool" | "remote_Control_Software" | "tool" | "trojanDenialOfService" | "trojanDropper" | "trojanMassMailer" | "trojanMonitoringSoftware" | "trojanProxyServer" | "virus" | "known" | "unknown" | "spp" | "behavior" | "vulnerability" | "policy"
export type WindowsMalwareExecutionState = "unknown" | "blocked" | "allowed" | "running" | "notRunning"
export type WindowsMalwareState = "unknown" | "detected" | "cleaned" | "quarantined" | "removed" | "allowed" | "blocked" | "cleanFailed" | "quarantineFailed" | "removeFailed" | "allowFailed" | "abandoned" | "blockFailed"
export type RemoteAction = "unknown" | "factoryReset" | "removeCompanyData" | "resetPasscode" | "remoteLock" | "enableLostMode" | "disableLostMode" | "locateDevice" | "rebootNow" | "recoverPasscode" | "cleanWindowsDevice" | "logoutSharedAppleDeviceActiveUser"
export type RunAsAccountType = "system" | "user"
export type RunState = "unknown" | "success" | "fail"
export type DeviceManagementSubscriptionState = "pending" | "active" | "warning" | "disabled" | "deleted" | "blocked" | "lockedOut"
export type DeviceManagementSubscriptions = "none" | "intune" | "office365" | "intunePremium" | "intune_EDU" | "intune_SMB"
export type HealthState = "unknown" | "healthy" | "unhealthy"
export type EasAuthenticationMethod = "usernameAndPassword" | "certificate"
export type EmailSyncDuration = "userDefined" | "oneDay" | "threeDays" | "oneWeek" | "twoWeeks" | "oneMonth" | "unlimited"
export type UserEmailSource = "userPrincipalName" | "primarySmtpAddress"
export type SubjectNameFormat = "commonName" | "commonNameIncludingEmail" | "commonNameAsEmail" | "custom"
export type SubjectAlternativeNameType = "emailAddress" | "userPrincipalName" | "customAzureADAttribute"
export type CertificateValidityPeriodScale = "days" | "months" | "years"
export type KeyUsages = "keyEncipherment" | "digitalSignature"
export type KeySize = "size1024" | "size2048"
export type HashAlgorithms = "sha1" | "sha2"
export type DevicePlatformType = "android" | "androidForWork" | "iOS" | "macOS" | "windowsPhone81" | "windows81AndLater" | "windows10AndLater"
export type AndroidUsernameSource = "username" | "userPrincipalName"
export type EmailSyncSchedule = "userDefined" | "asMessagesArrive" | "manual" | "fifteenMinutes" | "thirtyMinutes" | "sixtyMinutes" | "basedOnMyUsage"
export type AndroidWiFiSecurityType = "open" | "wpaEnterprise"
export type WiFiAuthenticationMethod = "certificate" | "usernameAndPassword"
export type AndroidEapType = "eapTls" | "eapTtls" | "peap"
export type NonEapAuthenticationMethodForEapTtlsType = "unencryptedPassword" | "challengeHandshakeAuthenticationProtocol" | "microsoftChap" | "microsoftChapVersionTwo"
export type NonEapAuthenticationMethodForPeap = "none" | "microsoftChapVersionTwo"
export type AndroidForWorkRequiredPasswordType = "deviceDefault" | "lowSecurityBiometric" | "required" | "atLeastNumeric" | "numericComplex" | "atLeastAlphabetic" | "atLeastAlphanumeric" | "alphanumericWithSymbols"
export type AndroidForWorkCrossProfileDataSharingType = "deviceDefault" | "preventAny" | "allowPersonalToWork" | "noRestrictions"
export type AndroidForWorkDefaultAppPermissionPolicyType = "deviceDefault" | "prompt" | "autoGrant" | "autoDeny"
export type AndroidForWorkVpnConnectionType = "ciscoAnyConnect" | "pulseSecure" | "f5EdgeClient" | "dellSonicWallMobileConnect" | "checkPointCapsuleVpn" | "citrix"
export type VpnAuthenticationMethod = "certificate" | "usernameAndPassword"
export type AppsComplianceListType = "none" | "appsInListCompliant" | "appsNotInListCompliant"
export type AppListType = "none" | "appsInListCompliant" | "appsNotInListCompliant"
export type AndroidRequiredPasswordType = "deviceDefault" | "alphabetic" | "alphanumeric" | "alphanumericWithSymbols" | "lowSecurityBiometric" | "numeric" | "numericComplex" | "any"
export type WebBrowserCookieSettings = "browserDefault" | "blockAlways" | "allowCurrentWebSite" | "allowFromWebsitesVisited" | "allowAlways"
export type AndroidVpnConnectionType = "ciscoAnyConnect" | "pulseSecure" | "f5EdgeClient" | "dellSonicWallMobileConnect" | "checkPointCapsuleVpn" | "citrix"
export type AppleSubjectNameFormat = "commonName" | "commonNameAsEmail" | "custom"
export type RatingAustraliaMoviesType = "allAllowed" | "allBlocked" | "general" | "parentalGuidance" | "mature" | "agesAbove15" | "agesAbove18"
export type RatingAustraliaTelevisionType = "allAllowed" | "allBlocked" | "preschoolers" | "children" | "general" | "parentalGuidance" | "mature" | "agesAbove15" | "agesAbove15AdultViolence"
export type RatingCanadaMoviesType = "allAllowed" | "allBlocked" | "general" | "parentalGuidance" | "agesAbove14" | "agesAbove18" | "restricted"
export type RatingCanadaTelevisionType = "allAllowed" | "allBlocked" | "children" | "childrenAbove8" | "general" | "parentalGuidance" | "agesAbove14" | "agesAbove18"
export type RatingFranceMoviesType = "allAllowed" | "allBlocked" | "agesAbove10" | "agesAbove12" | "agesAbove16" | "agesAbove18"
export type RatingFranceTelevisionType = "allAllowed" | "allBlocked" | "agesAbove10" | "agesAbove12" | "agesAbove16" | "agesAbove18"
export type RatingGermanyMoviesType = "allAllowed" | "allBlocked" | "general" | "agesAbove6" | "agesAbove12" | "agesAbove16" | "adults"
export type RatingGermanyTelevisionType = "allAllowed" | "allBlocked" | "general" | "agesAbove6" | "agesAbove12" | "agesAbove16" | "adults"
export type RatingIrelandMoviesType = "allAllowed" | "allBlocked" | "general" | "parentalGuidance" | "agesAbove12" | "agesAbove15" | "agesAbove16" | "adults"
export type RatingIrelandTelevisionType = "allAllowed" | "allBlocked" | "general" | "children" | "youngAdults" | "parentalSupervision" | "mature"
export type RatingJapanMoviesType = "allAllowed" | "allBlocked" | "general" | "parentalGuidance" | "agesAbove15" | "agesAbove18"
export type RatingJapanTelevisionType = "allAllowed" | "allBlocked" | "explicitAllowed"
export type RatingNewZealandMoviesType = "allAllowed" | "allBlocked" | "general" | "parentalGuidance" | "mature" | "agesAbove13" | "agesAbove15" | "agesAbove16" | "agesAbove18" | "restricted" | "agesAbove16Restricted"
export type RatingNewZealandTelevisionType = "allAllowed" | "allBlocked" | "general" | "parentalGuidance" | "adults"
export type RatingUnitedKingdomMoviesType = "allAllowed" | "allBlocked" | "general" | "universalChildren" | "parentalGuidance" | "agesAbove12Video" | "agesAbove12Cinema" | "agesAbove15" | "adults"
export type RatingUnitedKingdomTelevisionType = "allAllowed" | "allBlocked" | "caution"
export type RatingUnitedStatesMoviesType = "allAllowed" | "allBlocked" | "general" | "parentalGuidance" | "parentalGuidance13" | "restricted" | "adults"
export type RatingUnitedStatesTelevisionType = "allAllowed" | "allBlocked" | "childrenAll" | "childrenAbove7" | "general" | "parentalGuidance" | "childrenAbove14" | "adults"
export type RatingAppsType = "allAllowed" | "allBlocked" | "agesAbove4" | "agesAbove9" | "agesAbove12" | "agesAbove17"
export type RequiredPasswordType = "deviceDefault" | "alphanumeric" | "numeric"
export type RatingRegionType = "noRegion" | "australia" | "canada" | "france" | "germany" | "ireland" | "japan" | "newZealand" | "unitedKingdom" | "unitedStates"
export type WiFiSecurityType = "open" | "wpaPersonal" | "wpaEnterprise" | "wep"
export type WiFiProxySetting = "none" | "manual" | "automatic"
export type EapType = "eapTls" | "leap" | "eapSim" | "eapTtls" | "peap" | "eapFast"
export type EapFastConfiguration = "noProtectedAccessCredential" | "useProtectedAccessCredential" | "useProtectedAccessCredentialAndProvision" | "useProtectedAccessCredentialAndProvisionAnonymously"
export type IosNotificationAlertType = "deviceDefault" | "banner" | "modal" | "none"
export type AppleVpnConnectionType = "ciscoAnyConnect" | "pulseSecure" | "f5EdgeClient" | "dellSonicWallMobileConnect" | "checkPointCapsuleVpn" | "customVpn" | "ciscoIPSec" | "citrix"
export type VpnOnDemandRuleConnectionAction = "connect" | "evaluateConnection" | "ignore" | "disconnect"
export type VpnOnDemandRuleConnectionDomainAction = "connectIfNeeded" | "neverConnect"
export type ApplicationGuardBlockFileTransferType = "notConfigured" | "blockImageAndTextFile" | "blockImageFile" | "blockNone"
export type BitLockerEncryptionMethod = "aesCbc128" | "aesCbc256" | "xtsAes128" | "xtsAes256"
export type ConfigurationUsage = "blocked" | "required" | "allowed"
export type BitLockerRecoveryinformationType = "passwordAndKey" | "passwordOnly"
export type EdgeCookiePolicy = "userDefined" | "allow" | "blockThirdParty" | "blockAll"
export type DefenderThreatAction = "deviceDefault" | "clean" | "quarantine" | "remove" | "allow" | "userDefined" | "block"
export type WeeklySchedule = "userDefined" | "everyday" | "sunday" | "monday" | "tuesday" | "wednesday" | "thursday" | "friday" | "saturday"
export type DefenderMonitorFileActivity = "userDefined" | "disable" | "monitorAllFiles" | "monitorIncomingFilesOnly" | "monitorOutgoingFilesOnly"
export type DefenderPotentiallyUnwantedAppAction = "deviceDefault" | "block" | "audit"
export type DefenderPromptForSampleSubmission = "userDefined" | "alwaysPrompt" | "promptBeforeSendingPersonalData" | "neverSendData" | "sendAllDataWithoutPrompting"
export type DefenderScanType = "userDefined" | "disabled" | "quick" | "full"
export type StateManagementSetting = "notConfigured" | "blocked" | "allowed"
export type WindowsPrivacyDataAccessLevel = "notConfigured" | "forceAllow" | "forceDeny" | "userInControl"
export type WindowsPrivacyDataCategory = "notConfigured" | "accountInfo" | "appsRunInBackground" | "calendar" | "callHistory" | "camera" | "contacts" | "diagnosticsInfo" | "email" | "location" | "messaging" | "microphone" | "motion" | "notifications" | "phone" | "radios" | "tasks" | "syncWithDevices" | "trustedDevices"
export type WindowsStartMenuAppListVisibilityType = "userDefined" | "collapse" | "remove" | "disableSettingsApp"
export type WindowsStartMenuModeType = "userDefined" | "fullScreen" | "nonFullScreen"
export type VisibilitySetting = "notConfigured" | "hide" | "show"
export type WindowsSpotlightEnablementSettings = "notConfigured" | "disabled" | "enabled"
export type AutomaticUpdateMode = "userDefined" | "notifyDownload" | "autoInstallAtMaintenanceTime" | "autoInstallAndRebootAtMaintenanceTime" | "autoInstallAndRebootAtScheduledTime" | "autoInstallAndRebootWithoutEndUserControl"
export type DiagnosticDataSubmissionMode = "userDefined" | "none" | "basic" | "enhanced" | "full"
export type SafeSearchFilterType = "userDefined" | "strict" | "moderate"
export type EdgeSearchEngineType = "default" | "bing"
export type PrereleaseFeatures = "userDefined" | "settingsOnly" | "settingsAndExperimentations" | "notAllowed"
export type SharedPCAccountDeletionPolicyType = "immediate" | "diskSpaceThreshold" | "diskSpaceThresholdOrInactiveThreshold"
export type SharedPCAllowedAccountType = "guest" | "domain"
export type KeyStorageProviderOption = "useTpmKspOtherwiseUseSoftwareKsp" | "useTpmKspOtherwiseFail" | "usePassportForWorkKspOtherwiseFail" | "useSoftwareKsp"
export type CertificateDestinationStore = "computerCertStoreRoot" | "computerCertStoreIntermediate" | "userCertStoreIntermediate"
export type WindowsDeliveryOptimizationMode = "userDefined" | "httpOnly" | "httpWithPeeringNat" | "httpWithPeeringPrivateGroup" | "httpWithInternetPeering" | "simpleDownload" | "bypassMode"
export type WindowsUpdateRestartMode = "userDefined" | "batteryLevelCheckEnabled" | "batteryLevelCheckDisabled"
export type WindowsUpdateType = "userDefined" | "all" | "businessReadyOnly"
export type WindowsUpdateInsiderBuildControl = "userDefined" | "allowed" | "notAllowed"
export type Windows10VpnConnectionType = "pulseSecure" | "f5EdgeClient" | "dellSonicWallMobileConnect" | "checkPointCapsuleVpn" | "automatic" | "ikEv2" | "l2tp" | "pptp"
export type Windows10VpnAuthenticationMethod = "certificate" | "usernameAndPassword" | "customEapXml"
export type Windows10AppType = "desktop" | "universal"
export type VpnTrafficRuleAppType = "none" | "desktop" | "universal"
export type VpnTrafficRuleRoutingPolicyType = "none" | "splitTunnel" | "forceTunnel"
export type WindowsVpnConnectionType = "pulseSecure" | "f5EdgeClient" | "dellSonicWallMobileConnect" | "checkPointCapsuleVpn"
export type InternetSiteSecurityLevel = "userDefined" | "medium" | "mediumHigh" | "high"
export type SiteSecurityLevel = "userDefined" | "low" | "mediumLow" | "medium" | "mediumHigh" | "high"
export type UpdateClassification = "userDefined" | "recommendedAndImportant" | "important" | "none"
export type WindowsUserAccountControlSettings = "userDefined" | "alwaysNotify" | "notifyOnAppChanges" | "notifyOnAppChangesWithoutDimming" | "neverNotify"
export type MiracastChannel = "userDefined" | "one" | "two" | "three" | "four" | "five" | "six" | "seven" | "eight" | "nine" | "ten" | "eleven" | "thirtySix" | "forty" | "fortyFour" | "fortyEight" | "oneHundredFortyNine" | "oneHundredFiftyThree" | "oneHundredFiftySeven" | "oneHundredSixtyOne" | "oneHundredSixtyFive"
export type WelcomeScreenMeetingInformation = "userDefined" | "showOrganizerAndTimeOnly" | "showOrganizerAndTimeAndSubject"
export type EditionUpgradeLicenseType = "productKey" | "licenseFile"
export type Windows10EditionType = "windows10Enterprise" | "windows10EnterpriseN" | "windows10Education" | "windows10EducationN" | "windows10MobileEnterprise" | "windows10HolographicEnterprise" | "windows10Professional" | "windows10ProfessionalN" | "windows10ProfessionalEducation" | "windows10ProfessionalEducationN"
export type DeviceComplianceActionType = "noAction" | "notification" | "block" | "retire" | "wipe" | "removeResourceAccessProfiles"
export type NotificationTemplateBrandingOptions = "none" | "includeCompanyLogo" | "includeCompanyName" | "includeContactInformation"
export type DeviceThreatProtectionLevel = "unavailable" | "secured" | "low" | "medium" | "high" | "notSet"
export type CloudPkiProvider = "unKnown" | "symantec"
export type SyncStatus = "unKnown" | "succeeded" | "failed"
export type PolicyPlatformType = "android" | "androidForWork" | "iOS" | "macOS" | "windowsPhone81" | "windows81AndLater" | "windows10AndLater" | "all"
export type IosUpdatesInstallStatus = "success" | "available" | "idle" | "downloading" | "downloadFailed" | "downloadRequiresComputer" | "downloadInsufficientSpace" | "downloadInsufficientPower" | "downloadInsufficientNetwork" | "installing" | "installInsufficientSpace" | "installInsufficientPower" | "installPhoneCallInProgress" | "installFailed" | "notSupportedOperation"
export type DeviceManagementExchangeConnectorSyncType = "fullSync" | "deltaSync"
export type MdmAuthority = "unknown" | "intune" | "sccm" | "office365"
export type WindowsHelloForBusinessPinUsage = "allowed" | "required" | "disallowed"
export type WindowsHelloForBusinessConfiguration = "disabled" | "enabled" | "notConfigured"
export type DeviceManagementExchangeConnectorStatus = "connectionPending" | "connected" | "disconnected" | "none"
export type DeviceManagementExchangeConnectorType = "onPremises" | "hosted" | "serviceToService" | "dedicated"
export type DeviceManagementExchangeAccessLevel = "none" | "allow" | "block" | "quarantine"
export type ExchangeAccessRuleType = "family" | "model"
export type MobileThreatPartnerTenantState = "unavailable" | "available" | "enabled" | "unresponsive"
export type ManagedAppDataTransferLevel = "allApps" | "managedApps" | "none"
export type ManagedAppClipboardSharingLevel = "allApps" | "managedAppsWithPasteIn" | "managedApps" | "blocked"
export type ManagedAppPinCharacterSet = "any" | "numeric" | "alphanumeric" | "alphanumericAndSymbol"
export type ManagedAppDataStorageLocation = "oneDriveForBusiness" | "sharePoint" | "localStorage"
export type ManagedAppDataEncryptionType = "useDeviceSettings" | "afterDeviceRestart" | "whenDeviceLockedExceptOpenFiles" | "whenDeviceLocked"
export type WindowsInformationProtectionEnforcementLevel = "noProtection" | "encryptAndAuditOnly" | "encryptAuditAndPrompt" | "encryptAuditAndBlock"
export type WindowsInformationProtectionPinCharacterRequirements = "notAllow" | "requireAtLeastOne" | "allow"
export type ManagedAppFlaggedReason = "none" | "rootedDevice"
export type InstallIntent = "notApplicable" | "available" | "required" | "uninstall" | "availableWithoutEnrollment"
export type InstallState = "notApplicable" | "installed" | "failed" | "notInstalled" | "uninstallFailed" | "unknown"
export type RemoteAssistanceOnboardingStatus = "notOnboarded" | "onboarding" | "onboarded"
export type ApplicationType = "universal" | "desktop"
export type ViewType = "detail" | "activations" | "users" | "services" | "activity" | "groups" | "storage" | "apps" | "versions" | "mailbox" | "quota" | "files" | "account" | "sites" | "minutes" | "distribution" | "pages"
export type PeriodType = "d7" | "d30" | "d90" | "d180"

export interface Entity {

	    id?: string

}

export interface Extension extends Entity {

}

export interface DirectoryObject extends Entity {

	    deletedDateTime?: string

}

export interface User extends DirectoryObject {

	    accountEnabled?: boolean

	    assignedLicenses?: AssignedLicense[]

	    assignedPlans?: AssignedPlan[]

	    businessPhones?: string[]

	    city?: string

	    companyName?: string

	    country?: string

	    department?: string

	    deviceKeys?: DeviceKey[]

	    displayName?: string

	    employeeId?: string

	    givenName?: string

	    imAddresses?: string[]

	    jobTitle?: string

	    mail?: string

	    mailNickname?: string

	    mobilePhone?: string

	    onPremisesImmutableId?: string

	    onPremisesLastSyncDateTime?: string

	    onPremisesProvisioningErrors?: OnPremisesProvisioningError[]

	    onPremisesSecurityIdentifier?: string

	    onPremisesSyncEnabled?: boolean

	    onPremisesDomainName?: string

	    onPremisesSamAccountName?: string

	    onPremisesUserPrincipalName?: string

	    passwordPolicies?: string

	    passwordProfile?: PasswordProfile

	    officeLocation?: string

	    postalCode?: string

	    preferredDataLocation?: string

	    preferredLanguage?: string

	    provisionedPlans?: ProvisionedPlan[]

	    proxyAddresses?: string[]

	    refreshTokensValidFromDateTime?: string

	    showInAddressList?: boolean

	    state?: string

	    streetAddress?: string

	    surname?: string

	    usageLocation?: string

	    userPrincipalName?: string

	    userType?: string

	    mailboxSettings?: MailboxSettings

	    aboutMe?: string

	    birthday?: string

	    hireDate?: string

	    interests?: string[]

	    mySite?: string

	    pastProjects?: string[]

	    preferredName?: string

	    responsibilities?: string[]

	    schools?: string[]

	    skills?: string[]

	    identityUserRisk?: IdentityUserRisk

	    deviceEnrollmentLimit?: number

	    extensions?: Extension[]

	    ownedDevices?: DirectoryObject[]

	    registeredDevices?: DirectoryObject[]

	    manager?: DirectoryObject

	    directReports?: DirectoryObject[]

	    memberOf?: DirectoryObject[]

	    createdObjects?: DirectoryObject[]

	    ownedObjects?: DirectoryObject[]

	    scopedRoleMemberOf?: ScopedRoleMembership[]

	    licenseDetails?: LicenseDetails[]

	    activities?: Activity[]

	    outlook?: OutlookUser

	    messages?: Message[]

	    joinedGroups?: Group[]

	    mailFolders?: MailFolder[]

	    calendar?: Calendar

	    calendars?: Calendar[]

	    calendarGroups?: CalendarGroup[]

	    calendarView?: Event[]

	    events?: Event[]

	    people?: Person[]

	    contacts?: Contact[]

	    contactFolders?: ContactFolder[]

	    inferenceClassification?: InferenceClassification

	    photo?: ProfilePhoto

	    photos?: ProfilePhoto[]

	    drive?: Drive

	    drives?: Drive[]

	    insights?: OfficeGraphInsights

	    tasks?: Task[]

	    plans?: Plan[]

	    planner?: PlannerUser

	    onenote?: Onenote

	    managedDevices?: ManagedDevice[]

	    managedAppRegistrations?: ManagedAppRegistration[]

	    devices?: Device[]

	    joinedTeams?: Group[]

}

export interface ScopedRoleMembership extends Entity {

	    roleId?: string

	    administrativeUnitId?: string

	    roleMemberInfo?: Identity

}

export interface LicenseDetails extends Entity {

	    servicePlans?: ServicePlanInfo[]

	    skuId?: string

	    skuPartNumber?: string

}

export interface Activity extends Entity {

	    visualElements?: VisualInfo

	    activitySourceHost?: string

	    activationUrl?: string

	    appActivityId?: string

	    appDisplayName?: string

	    contentUrl?: string

	    createdDateTime?: string

	    expirationDateTime?: string

	    fallbackUrl?: string

	    lastModifiedDateTime?: string

	    userTimezone?: string

	    contentInfo?: any

	    historyItems?: HistoryItem[]

}

export interface OutlookUser extends Entity {

	    masterCategories?: OutlookCategory[]

	    taskGroups?: OutlookTaskGroup[]

	    taskFolders?: OutlookTaskFolder[]

	    tasks?: OutlookTask[]

}

export interface OutlookItem extends Entity {

	    createdDateTime?: string

	    lastModifiedDateTime?: string

	    changeKey?: string

	    categories?: string[]

}

export interface Message extends OutlookItem {

	    receivedDateTime?: string

	    sentDateTime?: string

	    hasAttachments?: boolean

	    internetMessageId?: string

	    internetMessageHeaders?: InternetMessageHeader[]

	    subject?: string

	    body?: ItemBody

	    bodyPreview?: string

	    importance?: Importance

	    parentFolderId?: string

	    sender?: Recipient

	    from?: Recipient

	    toRecipients?: Recipient[]

	    ccRecipients?: Recipient[]

	    bccRecipients?: Recipient[]

	    replyTo?: Recipient[]

	    conversationId?: string

	    conversationIndex?: number

	    uniqueBody?: ItemBody

	    isDeliveryReceiptRequested?: boolean

	    isReadReceiptRequested?: boolean

	    isRead?: boolean

	    isDraft?: boolean

	    webLink?: string

	    mentionsPreview?: MentionsPreview

	    inferenceClassification?: InferenceClassificationType

	    unsubscribeData?: string[]

	    unsubscribeEnabled?: boolean

	    flag?: FollowupFlag

	    attachments?: Attachment[]

	    extensions?: Extension[]

	    singleValueExtendedProperties?: SingleValueLegacyExtendedProperty[]

	    multiValueExtendedProperties?: MultiValueLegacyExtendedProperty[]

	    mentions?: Mention[]

}

export interface Group extends DirectoryObject {

	    classification?: string

	    createdDateTime?: string

	    description?: string

	    displayName?: string

	    groupTypes?: string[]

	    mail?: string

	    mailEnabled?: boolean

	    mailNickname?: string

	    membershipRule?: string

	    membershipRuleProcessingState?: string

	    onPremisesLastSyncDateTime?: string

	    onPremisesProvisioningErrors?: OnPremisesProvisioningError[]

	    onPremisesSecurityIdentifier?: string

	    onPremisesSyncEnabled?: boolean

	    preferredLanguage?: string

	    proxyAddresses?: string[]

	    renewedDateTime?: string

	    resourceBehaviorOptions?: string[]

	    resourceProvisioningOptions?: string[]

	    securityEnabled?: boolean

	    theme?: string

	    visibility?: string

	    accessType?: GroupAccessType

	    allowExternalSenders?: boolean

	    autoSubscribeNewMembers?: boolean

	    isFavorite?: boolean

	    isSubscribedByMail?: boolean

	    unseenCount?: number

	    unseenConversationsCount?: number

	    unseenMessagesCount?: number

	    extensions?: Extension[]

	    members?: DirectoryObject[]

	    memberOf?: DirectoryObject[]

	    createdOnBehalfOf?: DirectoryObject

	    owners?: DirectoryObject[]

	    settings?: DirectorySetting[]

	    endpoints?: Endpoint[]

	    threads?: ConversationThread[]

	    calendar?: Calendar

	    calendarView?: Event[]

	    events?: Event[]

	    conversations?: Conversation[]

	    photo?: ProfilePhoto

	    photos?: ProfilePhoto[]

	    acceptedSenders?: DirectoryObject[]

	    rejectedSenders?: DirectoryObject[]

	    drive?: Drive

	    drives?: Drive[]

	    sites?: Site[]

	    plans?: Plan[]

	    planner?: PlannerGroup

	    onenote?: Onenote

	    channels?: Channel[]

	    groupLifecyclePolicies?: GroupLifecyclePolicy[]

}

export interface MailFolder extends Entity {

	    displayName?: string

	    parentFolderId?: string

	    childFolderCount?: number

	    unreadItemCount?: number

	    totalItemCount?: number

	    wellKnownName?: string

	    messages?: Message[]

	    messageRules?: MessageRule[]

	    childFolders?: MailFolder[]

	    userConfigurations?: UserConfiguration[]

	    singleValueExtendedProperties?: SingleValueLegacyExtendedProperty[]

	    multiValueExtendedProperties?: MultiValueLegacyExtendedProperty[]

}

export interface Calendar extends Entity {

	    name?: string

	    color?: CalendarColor

	    hexColor?: string

	    isDefaultCalendar?: boolean

	    changeKey?: string

	    canShare?: boolean

	    canViewPrivateItems?: boolean

	    isShared?: boolean

	    isSharedWithMe?: boolean

	    canEdit?: boolean

	    owner?: EmailAddress

	    events?: Event[]

	    calendarView?: Event[]

	    singleValueExtendedProperties?: SingleValueLegacyExtendedProperty[]

	    multiValueExtendedProperties?: MultiValueLegacyExtendedProperty[]

}

export interface CalendarGroup extends Entity {

	    name?: string

	    classId?: string

	    changeKey?: string

	    calendars?: Calendar[]

}

export interface Event extends OutlookItem {

	    originalStartTimeZone?: string

	    originalEndTimeZone?: string

	    responseStatus?: ResponseStatus

	    iCalUId?: string

	    reminderMinutesBeforeStart?: number

	    isReminderOn?: boolean

	    hasAttachments?: boolean

	    subject?: string

	    body?: ItemBody

	    bodyPreview?: string

	    importance?: Importance

	    sensitivity?: Sensitivity

	    start?: DateTimeTimeZone

	    originalStart?: string

	    end?: DateTimeTimeZone

	    location?: Location

	    locations?: Location[]

	    isAllDay?: boolean

	    isCancelled?: boolean

	    isOrganizer?: boolean

	    recurrence?: PatternedRecurrence

	    responseRequested?: boolean

	    seriesMasterId?: string

	    showAs?: FreeBusyStatus

	    type?: EventType

	    attendees?: Attendee[]

	    organizer?: Recipient

	    webLink?: string

	    onlineMeetingUrl?: string

	    creationOptions?: EventCreationOptions

	    calendar?: Calendar

	    instances?: Event[]

	    extensions?: Extension[]

	    attachments?: Attachment[]

	    singleValueExtendedProperties?: SingleValueLegacyExtendedProperty[]

	    multiValueExtendedProperties?: MultiValueLegacyExtendedProperty[]

}

export interface Person extends Entity {

	    displayName?: string

	    givenName?: string

	    surname?: string

	    birthday?: string

	    personNotes?: string

	    isFavorite?: boolean

	    emailAddresses?: RankedEmailAddress[]

	    phones?: Phone[]

	    postalAddresses?: Location[]

	    websites?: Website[]

	    title?: string

	    companyName?: string

	    yomiCompany?: string

	    department?: string

	    officeLocation?: string

	    profession?: string

	    sources?: PersonDataSource[]

	    mailboxType?: string

	    personType?: string

	    userPrincipalName?: string

}

export interface Contact extends OutlookItem {

	    parentFolderId?: string

	    birthday?: string

	    fileAs?: string

	    displayName?: string

	    givenName?: string

	    initials?: string

	    middleName?: string

	    nickName?: string

	    surname?: string

	    title?: string

	    yomiGivenName?: string

	    yomiSurname?: string

	    yomiCompanyName?: string

	    generation?: string

	    emailAddresses?: EmailAddress[]

	    websites?: Website[]

	    imAddresses?: string[]

	    jobTitle?: string

	    companyName?: string

	    department?: string

	    officeLocation?: string

	    profession?: string

	    assistantName?: string

	    manager?: string

	    phones?: Phone[]

	    postalAddresses?: PhysicalAddress[]

	    spouseName?: string

	    personalNotes?: string

	    children?: string[]

	    weddingAnniversary?: string

	    gender?: string

	    isFavorite?: boolean

	    flag?: FollowupFlag

	    extensions?: Extension[]

	    singleValueExtendedProperties?: SingleValueLegacyExtendedProperty[]

	    multiValueExtendedProperties?: MultiValueLegacyExtendedProperty[]

	    photo?: ProfilePhoto

}

export interface ContactFolder extends Entity {

	    parentFolderId?: string

	    displayName?: string

	    wellKnownName?: string

	    contacts?: Contact[]

	    childFolders?: ContactFolder[]

	    singleValueExtendedProperties?: SingleValueLegacyExtendedProperty[]

	    multiValueExtendedProperties?: MultiValueLegacyExtendedProperty[]

}

export interface InferenceClassification extends Entity {

	    overrides?: InferenceClassificationOverride[]

}

export interface ProfilePhoto extends Entity {

	    height?: number

	    width?: number

}

export interface BaseItem extends Entity {

	    createdBy?: IdentitySet

	    createdDateTime?: string

	    description?: string

	    eTag?: string

	    lastModifiedBy?: IdentitySet

	    lastModifiedDateTime?: string

	    name?: string

	    parentReference?: ItemReference

	    webUrl?: string

	    createdByUser?: User

	    lastModifiedByUser?: User

}

export interface Drive extends BaseItem {

	    driveType?: string

	    owner?: IdentitySet

	    quota?: Quota

	    sharePointIds?: SharepointIds

	    system?: SystemFacet

	    activities?: ItemActivity[]

	    items?: DriveItem[]

	    list?: List

	    root?: DriveItem

	    special?: DriveItem[]

}

export interface OfficeGraphInsights extends Entity {

	    trending?: Trending[]

	    shared?: SharedInsight[]

	    used?: UsedInsight[]

}

export interface Task extends Entity {

	    createdBy?: string

	    assignedTo?: string

	    planId?: string

	    bucketId?: string

	    title?: string

	    orderHint?: string

	    assigneePriority?: string

	    percentComplete?: number

	    startDateTime?: string

	    assignedDateTime?: string

	    createdDateTime?: string

	    assignedBy?: string

	    dueDateTime?: string

	    hasDescription?: boolean

	    previewType?: PreviewType

	    completedDateTime?: string

	    appliedCategories?: AppliedCategoriesCollection

	    conversationThreadId?: string

	    details?: TaskDetails

	    assignedToTaskBoardFormat?: TaskBoardTaskFormat

	    progressTaskBoardFormat?: TaskBoardTaskFormat

	    bucketTaskBoardFormat?: TaskBoardTaskFormat

}

export interface Plan extends Entity {

	    createdBy?: string

	    createdDateTime?: string

	    owner?: string

	    title?: string

	    isVisibleInPlannerWebClient?: boolean

	    tasks?: Task[]

	    buckets?: Bucket[]

	    details?: PlanDetails

	    assignedToTaskBoard?: PlanTaskBoard

	    progressTaskBoard?: PlanTaskBoard

	    bucketTaskBoard?: PlanTaskBoard

}

export interface PlannerUser extends Entity {

	    tasks?: PlannerTask[]

	    plans?: PlannerPlan[]

}

export interface Onenote extends Entity {

	    notebooks?: Notebook[]

	    sections?: OnenoteSection[]

	    sectionGroups?: SectionGroup[]

	    pages?: OnenotePage[]

	    resources?: OnenoteResource[]

	    operations?: OnenoteOperation[]

}

export interface ManagedDevice extends Entity {

	    userId?: string

	    deviceName?: string

	    hardwareInformation?: HardwareInformation

	    ownerType?: OwnerType

	    deviceActionResults?: DeviceActionResult[]

	    managementState?: ManagementState

	    enrolledDateTime?: string

	    lastSyncDateTime?: string

	    chassisType?: ChassisType

	    operatingSystem?: string

	    deviceType?: DeviceType

	    complianceState?: ComplianceState

	    jailBroken?: string

	    managementAgent?: ManagementAgentType

	    osVersion?: string

	    easActivated?: boolean

	    easDeviceId?: string

	    easActivationDateTime?: string

	    aadRegistered?: boolean

	    enrollmentType?: EnrollmentType

	    lostModeState?: LostModeState

	    activationLockBypassCode?: string

	    emailAddress?: string

	    azureActiveDirectoryDeviceId?: string

	    deviceRegistrationState?: DeviceRegistrationState

	    deviceCategoryDisplayName?: string

	    isSupervised?: boolean

	    exchangeLastSuccessfulSyncDateTime?: string

	    exchangeAccessState?: DeviceManagementExchangeAccessState

	    exchangeAccessStateReason?: DeviceManagementExchangeAccessStateReason

	    remoteAssistanceSessionUrl?: string

	    isEncrypted?: boolean

	    userPrincipalName?: string

	    model?: string

	    manufacturer?: string

	    imei?: string

	    complianceGracePeriodExpirationDateTime?: string

	    serialNumber?: string

	    phoneNumber?: string

	    androidSecurityPatchLevel?: string

	    userDisplayName?: string

	    configurationManagerClientEnabledFeatures?: ConfigurationManagerClientEnabledFeatures

	    deviceConfigurationStates?: DeviceConfigurationState[]

	    detectedApps?: DetectedApp[]

	    deviceCategory?: DeviceCategory

	    deviceCompliancePolicyStates?: DeviceCompliancePolicyState[]

}

export interface ManagedAppRegistration extends Entity {

	    createdDateTime?: string

	    lastSyncDateTime?: string

	    applicationVersion?: string

	    managementSdkVersion?: string

	    platformVersion?: string

	    deviceType?: string

	    deviceTag?: string

	    deviceName?: string

	    flaggedReasons?: ManagedAppFlaggedReason[]

	    userId?: string

	    appIdentifier?: MobileAppIdentifier

	    version?: string

	    appliedPolicies?: ManagedAppPolicy[]

	    intendedPolicies?: ManagedAppPolicy[]

	    operations?: ManagedAppOperation[]

}

export interface Device extends DirectoryObject {

	    accountEnabled?: boolean

	    alternativeSecurityIds?: AlternativeSecurityId[]

	    approximateLastSignInDateTime?: string

	    deviceId?: string

	    deviceMetadata?: string

	    deviceVersion?: number

	    displayName?: string

	    isCompliant?: boolean

	    isManaged?: boolean

	    onPremisesLastSyncDateTime?: string

	    onPremisesSyncEnabled?: boolean

	    operatingSystem?: string

	    operatingSystemVersion?: string

	    physicalIds?: string[]

	    trustType?: string

	    Name?: string

	    Manufacturer?: string

	    Model?: string

	    Kind?: string

	    Status?: string

	    Platform?: string

	    extensions?: Extension[]

	    registeredOwners?: DirectoryObject[]

	    registeredUsers?: DirectoryObject[]

	    commands?: Command[]

}

export interface DirectorySetting extends Entity {

	    displayName?: string

	    templateId?: string

	    values?: SettingValue[]

}

export interface Endpoint extends DirectoryObject {

	    capability?: string

	    providerId?: string

	    providerName?: string

	    uri?: string

	    providerResourceId?: string

}

export interface ConversationThread extends Entity {

	    toRecipients?: Recipient[]

	    topic?: string

	    hasAttachments?: boolean

	    lastDeliveredDateTime?: string

	    uniqueSenders?: string[]

	    ccRecipients?: Recipient[]

	    preview?: string

	    isLocked?: boolean

	    posts?: Post[]

}

export interface Conversation extends Entity {

	    topic?: string

	    hasAttachments?: boolean

	    lastDeliveredDateTime?: string

	    uniqueSenders?: string[]

	    preview?: string

	    threads?: ConversationThread[]

}

export interface Site extends BaseItem {

	    displayName?: string

	    root?: Root

	    sharepointIds?: SharepointIds

	    siteCollection?: SiteCollection

	    columns?: ColumnDefinition[]

	    contentTypes?: ContentType[]

	    drive?: Drive

	    drives?: Drive[]

	    items?: BaseItem[]

	    lists?: List[]

	    sites?: Site[]

	    onenote?: Onenote

}

export interface PlannerGroup extends Entity {

	    plans?: PlannerPlan[]

}

export interface Channel extends Entity {

	    displayName?: string

	    description?: string

	    chatThreads?: ChatThread[]

}

export interface GroupLifecyclePolicy extends Entity {

	    groupLifetimeInDays?: number

	    managedGroupTypes?: string

	    alternateNotificationEmails?: string

}

export interface Command extends Entity {

	    Status?: string

	    Type?: string

	    AppServiceName?: string

	    PackageFamilyName?: string

	    Error?: string

	    Payload?: PayloadRequest

	    PermissionTicket?: string

	    PostBackUri?: string

	    responsepayload?: PayloadResponse

}

export interface AdministrativeUnit extends DirectoryObject {

	    displayName?: string

	    description?: string

	    visibility?: string

	    extensions?: Extension[]

	    members?: DirectoryObject[]

	    scopedRoleMembers?: ScopedRoleMembership[]

}

export interface Organization extends DirectoryObject {

	    assignedPlans?: AssignedPlan[]

	    businessPhones?: string[]

	    city?: string

	    country?: string

	    countryLetterCode?: string

	    displayName?: string

	    isMultipleDataLocationsForServicesEnabled?: boolean

	    marketingNotificationEmails?: string[]

	    onPremisesLastSyncDateTime?: string

	    onPremisesSyncEnabled?: boolean

	    postalCode?: string

	    preferredLanguage?: string

	    provisionedPlans?: ProvisionedPlan[]

	    securityComplianceNotificationMails?: string[]

	    securityComplianceNotificationPhones?: string[]

	    state?: string

	    street?: string

	    technicalNotificationMails?: string[]

	    verifiedDomains?: VerifiedDomain[]

	    mobileDeviceManagementAuthority?: MdmAuthority

	    defaultDeviceEnrollmentRestrictions?: DefaultDeviceEnrollmentRestrictions

	    defaultDeviceEnrollmentWindowsHelloForBusinessSettings?: DefaultDeviceEnrollmentWindowsHelloForBusinessSettings

	    defaultDeviceEnrollmentLimit?: number

	    certificateConnectorSetting?: CertificateConnectorSetting

	    extensions?: Extension[]

	    depOnboardingSettings?: DepOnboardingSetting[]

	    appleVolumePurchaseProgramTokens?: AppleVolumePurchaseProgramToken[]

}

export interface DepOnboardingSetting extends Entity {

	    appleIdentifier?: string

	    tokenExpirationDateTime?: string

	    lastModifiedDateTime?: string

	    lastSuccessfulSyncDateTime?: string

	    lastSyncTriggeredDateTime?: string

	    shareTokenWithSchoolDataSyncService?: boolean

	    lastSyncErrorCode?: number

}

export interface AppleVolumePurchaseProgramToken extends Entity {

	    organizationName?: string

	    volumePurchaseProgramTokenAccountType?: VolumePurchaseProgramTokenAccountType

	    appleId?: string

	    expirationDateTime?: string

	    lastSyncDateTime?: string

	    token?: string

	    lastModifiedDateTime?: string

	    state?: VolumePurchaseProgramTokenState

	    lastSyncStatus?: VolumePurchaseProgramTokenSyncStatus

	    automaticallyUpdateApps?: boolean

	    countryOrRegion?: string

}

export interface SchemaExtension extends Entity {

	    description?: string

	    targetTypes?: string[]

	    properties?: ExtensionSchemaProperty[]

	    status?: string

	    owner?: string

}

export interface Directory extends Entity {

	    deletedItems?: DirectoryObject[]

}

export interface ExtensionProperty extends DirectoryObject {

	    appDisplayName?: string

	    name?: string

	    dataType?: string

	    isSyncedFromOnPremises?: boolean

	    targetObjects?: string[]

}

export interface AllowedDataLocation extends Entity {

	    appId?: string

	    location?: string

	    isDefault?: boolean

	    domain?: string

}

export interface Application extends DirectoryObject {

	    api?: Api

	    allowPublicClient?: boolean

	    applicationAliases?: string[]

	    appRoles?: AppRole[]

	    createdDateTime?: string

	    installedClients?: InstalledClient

	    displayName?: string

	    info?: InformationalUrl

	    keyCredentials?: KeyCredential[]

	    logo?: any

	    orgRestrictions?: string[]

	    passwordCredentials?: PasswordCredential[]

	    preAuthorizedApplications?: PreAuthorizedApplication[]

	    requiredResourceAccess?: RequiredResourceAccess[]

	    tags?: string[]

	    web?: Web

	    extensionProperties?: ExtensionProperty[]

	    createdOnBehalfOf?: DirectoryObject

	    owners?: DirectoryObject[]

	    policies?: DirectoryObject[]

}

export interface AppRoleAssignment extends Entity {

	    creationTimestamp?: string

	    principalDisplayName?: string

	    principalId?: string

	    principalType?: string

	    resourceDisplayName?: string

	    resourceId?: string

}

export interface OrgContact extends DirectoryObject {

	    businessPhones?: string[]

	    city?: string

	    companyName?: string

	    country?: string

	    department?: string

	    displayName?: string

	    givenName?: string

	    jobTitle?: string

	    mail?: string

	    mailNickname?: string

	    mobilePhone?: string

	    onPremisesSyncEnabled?: boolean

	    onPremisesLastSyncDateTime?: string

	    onPremisesProvisioningErrors?: OnPremisesProvisioningError[]

	    officeLocation?: string

	    postalCode?: string

	    proxyAddresses?: string[]

	    state?: string

	    streetAddress?: string

	    surname?: string

	    manager?: DirectoryObject

	    directReports?: DirectoryObject[]

	    memberOf?: DirectoryObject[]

}

export interface DirectoryRole extends DirectoryObject {

	    description?: string

	    displayName?: string

	    roleTemplateId?: string

	    members?: DirectoryObject[]

	    scopedMembers?: ScopedRoleMembership[]

}

export interface DirectoryRoleTemplate extends DirectoryObject {

	    description?: string

	    displayName?: string

}

export interface DirectorySettingTemplate extends DirectoryObject {

	    displayName?: string

	    description?: string

	    values?: SettingTemplateValue[]

}

export interface Domain extends Entity {

	    authenticationType?: string

	    availabilityStatus?: string

	    isAdminManaged?: boolean

	    isDefault?: boolean

	    isInitial?: boolean

	    isRoot?: boolean

	    isVerified?: boolean

	    supportedServices?: string[]

	    state?: DomainState

	    serviceConfigurationRecords?: DomainDnsRecord[]

	    verificationDnsRecords?: DomainDnsRecord[]

	    domainNameReferences?: DirectoryObject[]

}

export interface DomainDnsRecord extends Entity {

	    isOptional?: boolean

	    label?: string

	    recordType?: string

	    supportedService?: string

	    ttl?: number

}

export interface DomainDnsCnameRecord extends DomainDnsRecord {

	    canonicalName?: string

}

export interface DomainDnsMxRecord extends DomainDnsRecord {

	    mailExchange?: string

	    preference?: number

}

export interface DomainDnsSrvRecord extends DomainDnsRecord {

	    nameTarget?: string

	    port?: number

	    priority?: number

	    protocol?: string

	    service?: string

	    weight?: number

}

export interface DomainDnsTxtRecord extends DomainDnsRecord {

	    text?: string

}

export interface DomainDnsUnavailableRecord extends DomainDnsRecord {

	    description?: string

}

export interface OAuth2PermissionGrant extends Entity {

	    clientId?: string

	    consentType?: string

	    expiryTime?: string

	    principalId?: string

	    resourceId?: string

	    scope?: string

	    startTime?: string

}

export interface Policy extends DirectoryObject {

	    alternativeIdentifier?: string

	    definition?: string[]

	    displayName?: string

	    isOrganizationDefault?: boolean

	    keyCredentials?: KeyCredential[]

	    type?: string

	    appliesTo?: DirectoryObject[]

}

export interface ServicePrincipal extends DirectoryObject {

	    accountEnabled?: boolean

	    addIns?: AddIn[]

	    appDisplayName?: string

	    appId?: string

	    appOwnerOrganizationId?: string

	    appRoleAssignmentRequired?: boolean

	    appRoles?: AppRole[]

	    displayName?: string

	    errorUrl?: string

	    homepage?: string

	    keyCredentials?: KeyCredential[]

	    logoutUrl?: string

	    oauth2Permissions?: OAuth2Permission[]

	    passwordCredentials?: PasswordCredential[]

	    preferredTokenSigningKeyThumbprint?: string

	    publisherName?: string

	    replyUrls?: string[]

	    samlMetadataUrl?: string

	    servicePrincipalNames?: string[]

	    tags?: string[]

	    appRoleAssignedTo?: AppRoleAssignment[]

	    appRoleAssignments?: AppRoleAssignment[]

	    oauth2PermissionGrants?: OAuth2PermissionGrant[]

	    memberOf?: DirectoryObject[]

	    createdObjects?: DirectoryObject[]

	    licenseDetails?: LicenseDetails[]

	    owners?: DirectoryObject[]

	    ownedObjects?: DirectoryObject[]

	    policies?: DirectoryObject[]

}

export interface SubscribedSku extends Entity {

	    capabilityStatus?: string

	    consumedUnits?: number

	    prepaidUnits?: LicenseUnitsDetail

	    servicePlans?: ServicePlanInfo[]

	    skuId?: string

	    skuPartNumber?: string

	    appliesTo?: string

}

export interface Contract extends DirectoryObject {

	    contractType?: string

	    customerId?: string

	    defaultDomainName?: string

	    displayName?: string

}

export interface HistoryItem extends Entity {

	    status?: Status

	    activeDurationSeconds?: number

	    createdDateTime?: string

	    lastActiveDateTime?: string

	    lastModifiedDateTime?: string

	    expirationDateTime?: string

	    startedDateTime?: string

	    userTimezone?: string

}

export interface ColumnDefinition extends Entity {

	    boolean?: BooleanColumn

	    calculated?: CalculatedColumn

	    choice?: ChoiceColumn

	    columnGroup?: string

	    currency?: CurrencyColumn

	    dateTime?: DateTimeColumn

	    defaultValue?: DefaultColumnValue

	    description?: string

	    displayName?: string

	    enforceUniqueValues?: boolean

	    hidden?: boolean

	    indexed?: boolean

	    lookup?: LookupColumn

	    name?: string

	    number?: NumberColumn

	    personOrGroup?: PersonOrGroupColumn

	    readOnly?: boolean

	    required?: boolean

	    text?: TextColumn

}

export interface ContentType extends Entity {

	    description?: string

	    group?: string

	    hidden?: boolean

	    inheritedFrom?: ItemReference

	    name?: string

	    order?: ContentTypeOrder

	    parentId?: string

	    readOnly?: boolean

	    sealed?: boolean

	    columnLinks?: ColumnLink[]

}

export interface List extends BaseItem {

	    displayName?: string

	    list?: ListInfo

	    sharepointIds?: SharepointIds

	    system?: SystemFacet

	    activities?: ItemActivity[]

	    columns?: ColumnDefinition[]

	    contentTypes?: ContentType[]

	    drive?: Drive

	    items?: ListItem[]

}

export interface ItemActivity extends Entity {

	    action?: ItemActionSet

	    actor?: IdentitySet

	    times?: ItemActivityTimeSet

	    driveItem?: DriveItem

	    listItem?: ListItem

}

export interface ListItem extends BaseItem {

	    contentType?: ContentTypeInfo

	    sharepointIds?: SharepointIds

	    activities?: ItemActivity[]

	    driveItem?: DriveItem

	    fields?: FieldValueSet

	    versions?: ListItemVersion[]

}

export interface DriveItem extends BaseItem {

	    audio?: Audio

	    content?: any

	    cTag?: string

	    deleted?: Deleted

	    file?: File

	    fileSystemInfo?: FileSystemInfo

	    folder?: Folder

	    image?: Image

	    location?: GeoCoordinates

	    package?: Package

	    photo?: Photo

	    publication?: PublicationFacet

	    remoteItem?: RemoteItem

	    root?: Root

	    searchResult?: SearchResult

	    shared?: Shared

	    sharepointIds?: SharepointIds

	    size?: number

	    specialFolder?: SpecialFolder

	    video?: Video

	    webDavUrl?: string

	    workbook?: Workbook

	    activities?: ItemActivity[]

	    children?: DriveItem[]

	    listItem?: ListItem

	    permissions?: Permission[]

	    thumbnails?: ThumbnailSet[]

	    versions?: DriveItemVersion[]

}

export interface Workbook extends Entity {

	    application?: WorkbookApplication

	    names?: WorkbookNamedItem[]

	    tables?: WorkbookTable[]

	    worksheets?: WorkbookWorksheet[]

	    functions?: WorkbookFunctions

}

export interface Permission extends Entity {

	    expirationDateTime?: string

	    grantedTo?: IdentitySet

	    grantedToIdentities?: IdentitySet[]

	    inheritedFrom?: ItemReference

	    invitation?: SharingInvitation

	    link?: SharingLink

	    roles?: string[]

	    shareId?: string

}

export interface ThumbnailSet extends Entity {

	    large?: Thumbnail

	    medium?: Thumbnail

	    small?: Thumbnail

	    source?: Thumbnail

}

export interface BaseItemVersion extends Entity {

	    lastModifiedBy?: IdentitySet

	    lastModifiedDateTime?: string

	    publication?: PublicationFacet

}

export interface DriveItemVersion extends BaseItemVersion {

	    content?: any

	    size?: number

}

export interface WorkbookApplication extends Entity {

	    calculationMode?: string

}

export interface WorkbookNamedItem extends Entity {

	    comment?: string

	    name?: string

	    scope?: string

	    type?: string

	    value?: any

	    visible?: boolean

	    worksheet?: WorkbookWorksheet

}

export interface WorkbookTable extends Entity {

	    highlightFirstColumn?: boolean

	    highlightLastColumn?: boolean

	    name?: string

	    showBandedColumns?: boolean

	    showBandedRows?: boolean

	    showFilterButton?: boolean

	    showHeaders?: boolean

	    showTotals?: boolean

	    style?: string

	    columns?: WorkbookTableColumn[]

	    rows?: WorkbookTableRow[]

	    sort?: WorkbookTableSort

	    worksheet?: WorkbookWorksheet

}

export interface WorkbookWorksheet extends Entity {

	    name?: string

	    position?: number

	    visibility?: string

	    charts?: WorkbookChart[]

	    names?: WorkbookNamedItem[]

	    pivotTables?: WorkbookPivotTable[]

	    protection?: WorkbookWorksheetProtection

	    tables?: WorkbookTable[]

}

export interface WorkbookFunctions extends Entity {

}

export interface WorkbookChart extends Entity {

	    height?: number

	    left?: number

	    name?: string

	    top?: number

	    width?: number

	    axes?: WorkbookChartAxes

	    dataLabels?: WorkbookChartDataLabels

	    format?: WorkbookChartAreaFormat

	    legend?: WorkbookChartLegend

	    series?: WorkbookChartSeries[]

	    title?: WorkbookChartTitle

	    worksheet?: WorkbookWorksheet

}

export interface WorkbookChartAxes extends Entity {

	    categoryAxis?: WorkbookChartAxis

	    seriesAxis?: WorkbookChartAxis

	    valueAxis?: WorkbookChartAxis

}

export interface WorkbookChartDataLabels extends Entity {

	    position?: string

	    separator?: string

	    showBubbleSize?: boolean

	    showCategoryName?: boolean

	    showLegendKey?: boolean

	    showPercentage?: boolean

	    showSeriesName?: boolean

	    showValue?: boolean

	    format?: WorkbookChartDataLabelFormat

}

export interface WorkbookChartAreaFormat extends Entity {

	    fill?: WorkbookChartFill

	    font?: WorkbookChartFont

}

export interface WorkbookChartLegend extends Entity {

	    overlay?: boolean

	    position?: string

	    visible?: boolean

	    format?: WorkbookChartLegendFormat

}

export interface WorkbookChartSeries extends Entity {

	    name?: string

	    format?: WorkbookChartSeriesFormat

	    points?: WorkbookChartPoint[]

}

export interface WorkbookChartTitle extends Entity {

	    overlay?: boolean

	    text?: string

	    visible?: boolean

	    format?: WorkbookChartTitleFormat

}

export interface WorkbookChartFill extends Entity {

}

export interface WorkbookChartFont extends Entity {

	    bold?: boolean

	    color?: string

	    italic?: boolean

	    name?: string

	    size?: number

	    underline?: string

}

export interface WorkbookChartAxis extends Entity {

	    majorUnit?: any

	    maximum?: any

	    minimum?: any

	    minorUnit?: any

	    format?: WorkbookChartAxisFormat

	    majorGridlines?: WorkbookChartGridlines

	    minorGridlines?: WorkbookChartGridlines

	    title?: WorkbookChartAxisTitle

}

export interface WorkbookChartAxisFormat extends Entity {

	    font?: WorkbookChartFont

	    line?: WorkbookChartLineFormat

}

export interface WorkbookChartGridlines extends Entity {

	    visible?: boolean

	    format?: WorkbookChartGridlinesFormat

}

export interface WorkbookChartAxisTitle extends Entity {

	    text?: string

	    visible?: boolean

	    format?: WorkbookChartAxisTitleFormat

}

export interface WorkbookChartLineFormat extends Entity {

	    color?: string

}

export interface WorkbookChartAxisTitleFormat extends Entity {

	    font?: WorkbookChartFont

}

export interface WorkbookChartDataLabelFormat extends Entity {

	    fill?: WorkbookChartFill

	    font?: WorkbookChartFont

}

export interface WorkbookChartGridlinesFormat extends Entity {

	    line?: WorkbookChartLineFormat

}

export interface WorkbookChartLegendFormat extends Entity {

	    fill?: WorkbookChartFill

	    font?: WorkbookChartFont

}

export interface WorkbookChartPoint extends Entity {

	    value?: any

	    format?: WorkbookChartPointFormat

}

export interface WorkbookChartPointFormat extends Entity {

	    fill?: WorkbookChartFill

}

export interface WorkbookChartSeriesFormat extends Entity {

	    fill?: WorkbookChartFill

	    line?: WorkbookChartLineFormat

}

export interface WorkbookChartTitleFormat extends Entity {

	    fill?: WorkbookChartFill

	    font?: WorkbookChartFont

}

export interface WorkbookFilter extends Entity {

	    criteria?: WorkbookFilterCriteria

}

export interface WorkbookFormatProtection extends Entity {

	    formulaHidden?: boolean

	    locked?: boolean

}

export interface WorkbookFunctionResult extends Entity {

	    error?: string

	    value?: any

}

export interface WorkbookPivotTable extends Entity {

	    name?: string

	    worksheet?: WorkbookWorksheet

}

export interface WorkbookRange extends Entity {

	    address?: string

	    addressLocal?: string

	    cellCount?: number

	    columnCount?: number

	    columnHidden?: boolean

	    columnIndex?: number

	    formulas?: any

	    formulasLocal?: any

	    formulasR1C1?: any

	    hidden?: boolean

	    numberFormat?: any

	    rowCount?: number

	    rowHidden?: boolean

	    rowIndex?: number

	    text?: any

	    valueTypes?: any

	    values?: any

	    format?: WorkbookRangeFormat

	    sort?: WorkbookRangeSort

	    worksheet?: WorkbookWorksheet

}

export interface WorkbookRangeFormat extends Entity {

	    columnWidth?: number

	    horizontalAlignment?: string

	    rowHeight?: number

	    verticalAlignment?: string

	    wrapText?: boolean

	    borders?: WorkbookRangeBorder[]

	    fill?: WorkbookRangeFill

	    font?: WorkbookRangeFont

	    protection?: WorkbookFormatProtection

}

export interface WorkbookRangeSort extends Entity {

}

export interface WorkbookRangeBorder extends Entity {

	    color?: string

	    sideIndex?: string

	    style?: string

	    weight?: string

}

export interface WorkbookRangeFill extends Entity {

	    color?: string

}

export interface WorkbookRangeFont extends Entity {

	    bold?: boolean

	    color?: string

	    italic?: boolean

	    name?: string

	    size?: number

	    underline?: string

}

export interface WorkbookRangeView extends Entity {

	    cellAddresses?: any

	    columnCount?: number

	    formulas?: any

	    formulasLocal?: any

	    formulasR1C1?: any

	    index?: number

	    numberFormat?: any

	    rowCount?: number

	    text?: any

	    valueTypes?: any

	    values?: any

	    rows?: WorkbookRangeView[]

}

export interface WorkbookTableColumn extends Entity {

	    index?: number

	    name?: string

	    values?: any

	    filter?: WorkbookFilter

}

export interface WorkbookTableRow extends Entity {

	    index?: number

	    values?: any

}

export interface WorkbookTableSort extends Entity {

	    fields?: WorkbookSortField[]

	    matchCase?: boolean

	    method?: string

}

export interface WorkbookWorksheetProtection extends Entity {

	    options?: WorkbookWorksheetProtectionOptions

	    protected?: boolean

}

export interface Attachment extends Entity {

	    lastModifiedDateTime?: string

	    name?: string

	    contentType?: string

	    size?: number

	    isInline?: boolean

}

export interface OutlookCategory extends Entity {

	    displayName?: string

	    color?: CategoryColor

}

export interface OutlookTaskGroup extends Entity {

	    changeKey?: string

	    isDefaultGroup?: boolean

	    name?: string

	    groupKey?: string

	    taskFolders?: OutlookTaskFolder[]

}

export interface OutlookTaskFolder extends Entity {

	    changeKey?: string

	    name?: string

	    isDefaultFolder?: boolean

	    parentGroupKey?: string

	    tasks?: OutlookTask[]

	    singleValueExtendedProperties?: SingleValueLegacyExtendedProperty[]

	    multiValueExtendedProperties?: MultiValueLegacyExtendedProperty[]

}

export interface OutlookTask extends OutlookItem {

	    assignedTo?: string

	    body?: ItemBody

	    completedDateTime?: DateTimeTimeZone

	    dueDateTime?: DateTimeTimeZone

	    hasAttachments?: boolean

	    importance?: Importance

	    isReminderOn?: boolean

	    owner?: string

	    parentFolderId?: string

	    recurrence?: PatternedRecurrence

	    reminderDateTime?: DateTimeTimeZone

	    sensitivity?: Sensitivity

	    startDateTime?: DateTimeTimeZone

	    status?: TaskStatus

	    subject?: string

	    singleValueExtendedProperties?: SingleValueLegacyExtendedProperty[]

	    multiValueExtendedProperties?: MultiValueLegacyExtendedProperty[]

	    attachments?: Attachment[]

}

export interface MessageRule extends Entity {

	    displayName?: string

	    sequence?: number

	    conditions?: MessageRulePredicates

	    actions?: MessageRuleActions

	    exceptions?: MessageRulePredicates

	    isEnabled?: boolean

	    hasError?: boolean

	    isReadOnly?: boolean

}

export interface UserConfiguration extends Entity {

	    binaryData?: number

}

export interface SingleValueLegacyExtendedProperty extends Entity {

	    value?: string

}

export interface MultiValueLegacyExtendedProperty extends Entity {

	    value?: string[]

}

export interface Mention extends Entity {

	    mentioned?: EmailAddress

	    mentionText?: string

	    clientReference?: string

	    createdBy?: EmailAddress

	    createdDateTime?: string

	    serverCreatedDateTime?: string

	    deepLink?: string

	    application?: string

}

export interface FileAttachment extends Attachment {

	    contentId?: string

	    contentLocation?: string

	    contentBytes?: number

}

export interface ItemAttachment extends Attachment {

	    item?: OutlookItem

}

export interface EventMessage extends Message {

	    meetingMessageType?: MeetingMessageType

	    startDateTime?: DateTimeTimeZone

	    endDateTime?: DateTimeTimeZone

	    location?: Location

	    type?: EventType

	    recurrence?: PatternedRecurrence

	    isOutOfDate?: boolean

	    isAllDay?: boolean

	    isDelegated?: boolean

	    event?: Event

}

export interface EventMessageRequest extends EventMessage {

	    previousLocation?: Location

	    previousStartDateTime?: DateTimeTimeZone

	    previousEndDateTime?: DateTimeTimeZone

	    responseRequested?: boolean

}

export interface ReferenceAttachment extends Attachment {

	    sourceUrl?: string

	    providerType?: ReferenceAttachmentProvider

	    thumbnailUrl?: string

	    previewUrl?: string

	    permission?: ReferenceAttachmentPermission

	    isFolder?: boolean

}

export interface OpenTypeExtension extends Extension {

	    extensionName?: string

}

export interface Post extends OutlookItem {

	    body?: ItemBody

	    receivedDateTime?: string

	    hasAttachments?: boolean

	    from?: Recipient

	    sender?: Recipient

	    conversationThreadId?: string

	    newParticipants?: Recipient[]

	    conversationId?: string

	    extensions?: Extension[]

	    inReplyTo?: Post

	    attachments?: Attachment[]

	    singleValueExtendedProperties?: SingleValueLegacyExtendedProperty[]

	    multiValueExtendedProperties?: MultiValueLegacyExtendedProperty[]

	    mentions?: Mention[]

}

export interface InferenceClassificationOverride extends Entity {

	    classifyAs?: InferenceClassificationType

	    senderEmailAddress?: EmailAddress

}

export interface ColumnLink extends Entity {

	    name?: string

}

export interface FieldValueSet extends Entity {

}

export interface ListItemVersion extends BaseItemVersion {

	    fields?: FieldValueSet

}

export interface SharedDriveItem extends BaseItem {

	    owner?: IdentitySet

	    driveItem?: DriveItem

	    items?: DriveItem[]

	    list?: List

	    listItem?: ListItem

	    root?: DriveItem

	    site?: Site

}

export interface Trending extends Entity {

	    weight?: number

	    resourceVisualization?: ResourceVisualization

	    resourceReference?: ResourceReference

	    lastModifiedDateTime?: string

	    resource?: Entity

}

export interface SharedInsight extends Entity {

	    lastShared?: SharingDetail

	    sharingHistory?: SharingDetail[]

	    resourceVisualization?: ResourceVisualization

	    resourceReference?: ResourceReference

	    lastSharedMethod?: Entity

	    resource?: Entity

}

export interface UsedInsight extends Entity {

	    lastUsed?: UsageDetails

	    resourceVisualization?: ResourceVisualization

	    resourceReference?: ResourceReference

	    resource?: Entity

}

export interface PlannerTask extends Entity {

	    createdBy?: IdentitySet

	    planId?: string

	    bucketId?: string

	    title?: string

	    orderHint?: string

	    assigneePriority?: string

	    percentComplete?: number

	    startDateTime?: string

	    createdDateTime?: string

	    dueDateTime?: string

	    hasDescription?: boolean

	    previewType?: PlannerPreviewType

	    completedDateTime?: string

	    completedBy?: IdentitySet

	    referenceCount?: number

	    checklistItemCount?: number

	    activeChecklistItemCount?: number

	    appliedCategories?: PlannerAppliedCategories

	    assignments?: PlannerAssignments

	    conversationThreadId?: string

	    details?: PlannerTaskDetails

	    assignedToTaskBoardFormat?: PlannerAssignedToTaskBoardTaskFormat

	    progressTaskBoardFormat?: PlannerProgressTaskBoardTaskFormat

	    bucketTaskBoardFormat?: PlannerBucketTaskBoardTaskFormat

}

export interface PlannerPlan extends Entity {

	    createdBy?: IdentitySet

	    createdDateTime?: string

	    owner?: string

	    title?: string

	    tasks?: PlannerTask[]

	    buckets?: PlannerBucket[]

	    details?: PlannerPlanDetails

}

export interface Planner extends Entity {

	    tasks?: PlannerTask[]

	    plans?: PlannerPlan[]

	    buckets?: PlannerBucket[]

}

export interface PlannerBucket extends Entity {

	    name?: string

	    planId?: string

	    orderHint?: string

	    tasks?: PlannerTask[]

}

export interface PlannerTaskDetails extends Entity {

	    description?: string

	    previewType?: PlannerPreviewType

	    references?: PlannerExternalReferences

	    checklist?: PlannerChecklistItems

}

export interface PlannerAssignedToTaskBoardTaskFormat extends Entity {

	    unassignedOrderHint?: string

	    orderHintsByAssignee?: PlannerOrderHintsByAssignee

}

export interface PlannerProgressTaskBoardTaskFormat extends Entity {

	    orderHint?: string

}

export interface PlannerBucketTaskBoardTaskFormat extends Entity {

	    orderHint?: string

}

export interface PlannerPlanDetails extends Entity {

	    sharedWith?: PlannerUserIds

	    categoryDescriptions?: PlannerCategoryDescriptions

}

export interface TaskDetails extends Entity {

	    description?: string

	    previewType?: PreviewType

	    completedBy?: string

	    references?: ExternalReferenceCollection

	    checklist?: ChecklistItemCollection

}

export interface TaskBoardTaskFormat extends Entity {

	    type?: TaskBoardType

	    orderHint?: string

}

export interface Bucket extends Entity {

	    name?: string

	    planId?: string

	    orderHint?: string

	    tasks?: Task[]

}

export interface PlanDetails extends Entity {

	    sharedWith?: UserIdCollection

	    category0Description?: string

	    category1Description?: string

	    category2Description?: string

	    category3Description?: string

	    category4Description?: string

	    category5Description?: string

}

export interface PlanTaskBoard extends Entity {

	    type?: TaskBoardType

}

export interface OnenoteEntityBaseModel extends Entity {

	    self?: string

}

export interface OnenoteEntitySchemaObjectModel extends OnenoteEntityBaseModel {

	    createdDateTime?: string

}

export interface OnenoteEntityHierarchyModel extends OnenoteEntitySchemaObjectModel {

	    displayName?: string

	    createdBy?: IdentitySet

	    lastModifiedBy?: IdentitySet

	    lastModifiedDateTime?: string

}

export interface Notebook extends OnenoteEntityHierarchyModel {

	    isDefault?: boolean

	    userRole?: OnenoteUserRole

	    isShared?: boolean

	    sectionsUrl?: string

	    sectionGroupsUrl?: string

	    links?: NotebookLinks

	    sections?: OnenoteSection[]

	    sectionGroups?: SectionGroup[]

}

export interface OnenoteSection extends OnenoteEntityHierarchyModel {

	    isDefault?: boolean

	    links?: SectionLinks

	    pagesUrl?: string

	    parentNotebook?: Notebook

	    parentSectionGroup?: SectionGroup

	    pages?: OnenotePage[]

}

export interface SectionGroup extends OnenoteEntityHierarchyModel {

	    sectionsUrl?: string

	    sectionGroupsUrl?: string

	    parentNotebook?: Notebook

	    parentSectionGroup?: SectionGroup

	    sections?: OnenoteSection[]

	    sectionGroups?: SectionGroup[]

}

export interface OnenotePage extends OnenoteEntitySchemaObjectModel {

	    title?: string

	    createdByAppId?: string

	    links?: PageLinks

	    contentUrl?: string

	    content?: any

	    lastModifiedDateTime?: string

	    level?: number

	    order?: number

	    userTags?: string[]

	    parentSection?: OnenoteSection

	    parentNotebook?: Notebook

}

export interface OnenoteResource extends OnenoteEntityBaseModel {

	    content?: any

	    contentUrl?: string

}

export interface Operation extends Entity {

	    status?: OperationStatus

	    createdDateTime?: string

	    lastActionDateTime?: string

}

export interface OnenoteOperation extends Operation {

	    resourceLocation?: string

	    resourceId?: string

	    error?: OnenoteOperationError

	    percentComplete?: string

}

export interface Subscription extends Entity {

	    resource?: string

	    changeType?: string

	    clientState?: string

	    notificationUrl?: string

	    expirationDateTime?: string

}

export interface IdentityRiskEvent extends Entity {

	    userDisplayName?: string

	    userPrincipalName?: string

	    riskEventDateTime?: string

	    riskEventType?: string

	    riskLevel?: RiskLevel

	    riskEventStatus?: RiskEventStatus

	    closedDateTime?: string

	    createdDateTime?: string

	    userId?: string

	    impactedUser?: User

}

export interface LocatedRiskEvent extends IdentityRiskEvent {

	    location?: SignInLocation

	    ipAddress?: string

}

export interface ImpossibleTravelRiskEvent extends LocatedRiskEvent {

	    userAgent?: string

	    deviceInformation?: string

	    isAtypicalLocation?: boolean

	    previousSigninDateTime?: string

	    previousLocation?: SignInLocation

	    previousIpAddress?: string

}

export interface LeakedCredentialsRiskEvent extends IdentityRiskEvent {

}

export interface AnonymousIpRiskEvent extends LocatedRiskEvent {

}

export interface SuspiciousIpRiskEvent extends LocatedRiskEvent {

}

export interface UnfamiliarLocationRiskEvent extends LocatedRiskEvent {

}

export interface MalwareRiskEvent extends LocatedRiskEvent {

	    deviceInformation?: string

	    malwareName?: string

}

export interface PrivilegedRole extends Entity {

	    name?: string

	    settings?: PrivilegedRoleSettings

	    assignments?: PrivilegedRoleAssignment[]

	    summary?: PrivilegedRoleSummary

}

export interface PrivilegedRoleSettings extends Entity {

	    approverIds?: string[]

	    minElevationDuration?: string

	    maxElavationDuration?: string

	    elevationDuration?: string

	    notificationToUserOnElevation?: boolean

	    ticketingInfoOnElevation?: boolean

	    mfaOnElevation?: boolean

	    lastGlobalAdmin?: boolean

	    isMfaOnElevationConfigurable?: boolean

	    approvalOnElevation?: boolean

}

export interface PrivilegedRoleAssignment extends Entity {

	    userId?: string

	    roleId?: string

	    isElevated?: boolean

	    expirationDateTime?: string

	    resultMessage?: string

	    roleInfo?: PrivilegedRole

}

export interface PrivilegedRoleSummary extends Entity {

	    status?: RoleSummaryStatus

	    usersCount?: number

	    managedCount?: number

	    elevatedCount?: number

	    mfaEnabled?: boolean

}

export interface PrivilegedOperationEvent extends Entity {

	    userId?: string

	    userName?: string

	    userMail?: string

	    roleId?: string

	    roleName?: string

	    expirationDateTime?: string

	    creationDateTime?: string

	    requestorId?: string

	    requestorName?: string

	    tenantId?: string

	    requestType?: string

	    additionalInformation?: string

	    referenceKey?: string

	    referenceSystem?: string

}

export interface PrivilegedSignupStatus extends Entity {

	    isRegistered?: boolean

	    status?: SetupStatus

}

export interface PrivilegedApproval extends Entity {

	    userId?: string

	    roleId?: string

	    approvalType?: string

	    approvalState?: ApprovalState

	    approvalDuration?: string

	    requestorReason?: string

	    approverReason?: string

	    startDateTime?: string

	    endDateTime?: string

	    roleInfo?: PrivilegedRole

}

export interface TenantSetupInfo extends Entity {

	    userRolesActions?: string

	    firstTimeSetup?: boolean

	    relevantRolesSettings?: string[]

	    skipSetup?: boolean

	    setupStatus?: SetupStatus

	    defaultRolesSettings?: PrivilegedRoleSettings

}

export interface Invitation extends Entity {

	    invitedUserDisplayName?: string

	    invitedUserType?: string

	    invitedUserEmailAddress?: string

	    invitedUserMessageInfo?: InvitedUserMessageInfo

	    sendInvitationMessage?: boolean

	    inviteRedirectUrl?: string

	    inviteRedeemUrl?: string

	    status?: string

	    invitedUser?: User

}

export interface DeviceManagement extends Entity {

	    subscriptionState?: DeviceManagementSubscriptionState

	    subscriptions?: DeviceManagementSubscriptions

	    settings?: DeviceManagementSettings

	    intuneBrand?: IntuneBrand

	    termsAndConditions?: TermsAndConditions[]

	    androidForWorkSettings?: AndroidForWorkSettings

	    androidForWorkAppConfigurationSchemas?: AndroidForWorkAppConfigurationSchema[]

	    enrollmentProfiles?: EnrollmentProfile[]

	    importedDeviceIdentities?: ImportedDeviceIdentity[]

	    importedAppleDeviceIdentities?: ImportedAppleDeviceIdentity[]

	    remoteActionAudits?: RemoteActionAudit[]

	    applePushNotificationCertificate?: ApplePushNotificationCertificate

	    deviceManagementScripts?: DeviceManagementScript[]

	    managedDeviceOverview?: ManagedDeviceOverview

	    detectedApps?: DetectedApp[]

	    managedDevices?: ManagedDevice[]

	    deviceConfigurations?: DeviceConfiguration[]

	    deviceCompliancePolicies?: DeviceCompliancePolicy[]

	    softwareUpdateStatusSummary?: SoftwareUpdateStatusSummary

	    cloudPkiSubscriptions?: CloudPkiSubscription[]

	    deviceCompliancePolicyDeviceStateSummary?: DeviceCompliancePolicyDeviceStateSummary

	    deviceCompliancePolicySettingStateSummaries?: DeviceCompliancePolicySettingStateSummary[]

	    deviceConfigurationDeviceStateSummaries?: DeviceConfigurationDeviceStateSummary

	    deviceConfigurationUserStateSummaries?: DeviceConfigurationUserStateSummary

	    cartToClassAssociations?: CartToClassAssociation[]

	    iosUpdateStatuses?: IosUpdateDeviceStatus[]

	    deviceCategories?: DeviceCategory[]

	    exchangeConnectors?: DeviceManagementExchangeConnector[]

	    exchangeOnPremisesPolicy?: DeviceManagementExchangeOnPremisesPolicy

	    mobileThreatDefenseConnectors?: MobileThreatDefenseConnector[]

	    notificationMessageTemplates?: NotificationMessageTemplate[]

	    roleDefinitions?: RoleDefinition[]

	    roleAssignments?: RoleAssignment[]

	    resourceOperations?: ResourceOperation[]

	    telecomExpenseManagementPartners?: TelecomExpenseManagementPartner[]

	    remoteAssistancePartners?: RemoteAssistancePartner[]

	    windowsInformationProtectionAppLearningSummaries?: WindowsInformationProtectionAppLearningSummary[]

	    windowsMalwareInformation?: WindowsMalwareInformation[]

}

export interface TermsAndConditions extends Entity {

	    createdDateTime?: string

	    modifiedDateTime?: string

	    displayName?: string

	    description?: string

	    title?: string

	    bodyText?: string

	    acceptanceStatement?: string

	    version?: number

	    groupAssignments?: TermsAndConditionsGroupAssignment[]

	    acceptanceStatuses?: TermsAndConditionsAcceptanceStatus[]

}

export interface AndroidForWorkSettings extends Entity {

	    bindStatus?: AndroidForWorkBindStatus

	    lastAppSyncDateTime?: string

	    lastAppSyncStatus?: AndroidForWorkSyncStatus

	    ownerUserPrincipalName?: string

	    ownerOrganizationName?: string

	    lastModifiedDateTime?: string

	    enrollmentTarget?: AndroidForWorkEnrollmentTarget

	    targetGroupIds?: string[]

}

export interface AndroidForWorkAppConfigurationSchema extends Entity {

	    exampleJson?: AndroidForWorkAppConfigurationExample

	    schemaItems?: AndroidForWorkAppConfigurationSchemaItem[]

}

export interface EnrollmentProfile extends Entity {

	    displayName?: string

	    description?: string

	    requiresUserAuthentication?: boolean

	    configurationEndpointUrl?: string

}

export interface ImportedDeviceIdentity extends Entity {

	    importedDeviceIdentifier?: string

	    importedDeviceIdentityType?: ImportedDeviceIdentityType

	    lastModifiedDateTime?: string

	    createdDateTime?: string

	    lastContactedDateTime?: string

	    description?: string

	    enrollmentState?: EnrollmentState

	    platform?: Platform

}

export interface ImportedAppleDeviceIdentity extends Entity {

	    serialNumber?: string

	    requestedEnrollmentProfileId?: string

	    requestedEnrollmentProfileAssignmentDateTime?: string

	    isSupervised?: boolean

	    discoverySource?: DiscoverySource

	    createdDateTime?: string

	    lastContactedDateTime?: string

	    description?: string

	    enrollmentState?: EnrollmentState

	    platform?: Platform

}

export interface RemoteActionAudit extends Entity {

	    deviceDisplayName?: string

	    userName?: string

	    initiatedByUserPrincipalName?: string

	    action?: RemoteAction

	    requestDateTime?: string

	    deviceOwnerUserPrincipalName?: string

	    deviceIMEI?: string

	    actionState?: DeviceActionState

}

export interface ApplePushNotificationCertificate extends Entity {

	    appleIdentifier?: string

	    topicIdentifier?: string

	    lastModifiedDateTime?: string

	    expirationDateTime?: string

	    certificateUploadStatus?: string

	    certificateUploadFailureReason?: string

	    certificate?: string

}

export interface DeviceManagementScript extends Entity {

	    displayName?: string

	    description?: string

	    runSchedule?: RunSchedule

	    scriptContent?: string

	    createdDateTime?: string

	    lastModifiedDateTime?: string

	    runAsAccount?: RunAsAccountType

	    enforceSignatureCheck?: boolean

	    fileName?: string

	    groupAssignments?: DeviceManagementScriptGroupAssignment[]

	    runSummary?: DeviceManagementScriptRunSummary

	    deviceRunStates?: DeviceManagementScriptDeviceState[]

	    userRunStates?: DeviceManagementScriptUserState[]

}

export interface ManagedDeviceOverview extends Entity {

	    enrolledDeviceCount?: number

	    mdmEnrolledCount?: number

	    dualEnrolledDeviceCount?: number

	    deviceOperatingSystemSummary?: DeviceOperatingSystemSummary

	    deviceExchangeAccessStateSummary?: DeviceExchangeAccessStateSummary

}

export interface DetectedApp extends Entity {

	    displayName?: string

	    version?: string

	    sizeInByte?: number

	    deviceCount?: number

	    managedDevices?: ManagedDevice[]

}

export interface DeviceConfiguration extends Entity {

	    lastModifiedDateTime?: string

	    assignmentStatus?: string

	    assignmentProgress?: string

	    assignmentErrorMessage?: string

	    createdDateTime?: string

	    description?: string

	    displayName?: string

	    version?: number

	    groupAssignments?: DeviceConfigurationGroupAssignment[]

	    deviceStatuses?: DeviceConfigurationDeviceStatus[]

	    userStatuses?: DeviceConfigurationUserStatus[]

	    deviceStatusOverview?: DeviceConfigurationDeviceOverview

	    userStatusOverview?: DeviceConfigurationUserOverview

	    deviceSettingStateSummaries?: SettingStateDeviceSummary[]

}

export interface DeviceCompliancePolicy extends Entity {

	    createdDateTime?: string

	    description?: string

	    lastModifiedDateTime?: string

	    displayName?: string

	    version?: number

	    groupAssignments?: DeviceCompliancePolicyGroupAssignment[]

	    scheduledActionsForRule?: DeviceComplianceScheduledActionForRule[]

	    deviceStatuses?: DeviceComplianceDeviceStatus[]

	    userStatuses?: DeviceComplianceUserStatus[]

	    deviceStatusOverview?: DeviceComplianceDeviceOverview

	    userStatusOverview?: DeviceComplianceUserOverview

	    deviceSettingStateSummaries?: SettingStateDeviceSummary[]

}

export interface SoftwareUpdateStatusSummary extends Entity {

	    displayName?: string

	    compliantDeviceCount?: number

	    nonCompliantDeviceCount?: number

	    remediatedDeviceCount?: number

	    errorDeviceCount?: number

	    unknownDeviceCount?: number

	    conflictDeviceCount?: number

	    notApplicableDeviceCount?: number

	    compliantUserCount?: number

	    nonCompliantUserCount?: number

	    remediatedUserCount?: number

	    errorUserCount?: number

	    unknownUserCount?: number

	    conflictUserCount?: number

	    notApplicableUserCount?: number

}

export interface CloudPkiSubscription extends Entity {

	    cloudPkiProvider?: CloudPkiProvider

	    createdDateTime?: string

	    description?: string

	    lastModifiedDateTime?: string

	    displayName?: string

	    syncStatus?: SyncStatus

	    lastSyncError?: string

	    lastSyncDateTime?: string

	    credentials?: CloudPkiAdministratorCredentials

	    trustedRootCertificate?: number

	    version?: number

}

export interface DeviceCompliancePolicyDeviceStateSummary extends Entity {

	    inGracePeriodCount?: number

	    configManagerCount?: number

	    unknownDeviceCount?: number

	    notApplicableDeviceCount?: number

	    compliantDeviceCount?: number

	    remediatedDeviceCount?: number

	    nonCompliantDeviceCount?: number

	    errorDeviceCount?: number

	    conflictDeviceCount?: number

}

export interface DeviceCompliancePolicySettingStateSummary extends Entity {

	    setting?: string

	    settingName?: string

	    platformType?: PolicyPlatformType

	    unknownDeviceCount?: number

	    notApplicableDeviceCount?: number

	    compliantDeviceCount?: number

	    remediatedDeviceCount?: number

	    nonCompliantDeviceCount?: number

	    errorDeviceCount?: number

	    conflictDeviceCount?: number

	    deviceComplianceSettingStates?: DeviceComplianceSettingState[]

}

export interface DeviceConfigurationDeviceStateSummary extends Entity {

	    unknownDeviceCount?: number

	    notApplicableDeviceCount?: number

	    compliantDeviceCount?: number

	    remediatedDeviceCount?: number

	    nonCompliantDeviceCount?: number

	    errorDeviceCount?: number

	    conflictDeviceCount?: number

}

export interface DeviceConfigurationUserStateSummary extends Entity {

	    unknownDeviceCount?: number

	    notApplicableDeviceCount?: number

	    compliantDeviceCount?: number

	    remediatedDeviceCount?: number

	    nonCompliantDeviceCount?: number

	    errorDeviceCount?: number

	    conflictDeviceCount?: number

}

export interface CartToClassAssociation extends Entity {

	    createdDateTime?: string

	    lastModifiedDateTime?: string

	    version?: number

	    displayName?: string

	    description?: string

	    deviceCartIds?: string[]

	    classroomIds?: string[]

}

export interface IosUpdateDeviceStatus extends Entity {

	    installStatus?: IosUpdatesInstallStatus

	    osVersion?: string

	    deviceId?: string

	    userId?: string

	    deviceDisplayName?: string

	    userName?: string

	    deviceModel?: string

	    platform?: number

	    complianceGracePeriodExpirationDateTime?: string

	    status?: ComplianceStatus

	    lastReportedDateTime?: string

	    userPrincipalName?: string

}

export interface DeviceCategory extends Entity {

	    displayName?: string

	    description?: string

}

export interface DeviceManagementExchangeConnector extends Entity {

	    lastSyncDateTime?: string

	    status?: DeviceManagementExchangeConnectorStatus

	    primarySmtpAddress?: string

	    serverName?: string

	    exchangeConnectorType?: DeviceManagementExchangeConnectorType

	    version?: string

}

export interface DeviceManagementExchangeOnPremisesPolicy extends Entity {

	    notificationContent?: number

	    defaultAccessLevel?: DeviceManagementExchangeAccessLevel

	    accessRules?: DeviceManagementExchangeAccessRule[]

	    knownDeviceClasses?: DeviceManagementExchangeDeviceClass[]

	    conditionalAccessSettings?: OnPremisesConditionalAccessSettings

}

export interface MobileThreatDefenseConnector extends Entity {

	    lastHeartbeatDateTime?: string

	    partnerState?: MobileThreatPartnerTenantState

	    androidEnabled?: boolean

	    androidDeviceBlockedOnMissingPartnerData?: boolean

	    iosDeviceBlockedOnMissingPartnerData?: boolean

	    partnerUnsupportedOsVersionBlocked?: boolean

	    iosEnabled?: boolean

	    partnerUnresponsivenessThresholdInDays?: number

}

export interface NotificationMessageTemplate extends Entity {

	    lastModifiedDateTime?: string

	    displayName?: string

	    defaultLocale?: string

	    brandingOptions?: NotificationTemplateBrandingOptions

	    localizedNotificationMessages?: LocalizedNotificationMessage[]

}

export interface RoleDefinition extends Entity {

	    displayName?: string

	    description?: string

	    permissions?: RolePermission[]

	    isBuiltInRoleDefinition?: boolean

	    roleAssignments?: RoleAssignment[]

}

export interface RoleAssignment extends Entity {

	    displayName?: string

	    description?: string

	    members?: string[]

	    scopeMembers?: string[]

	    roleDefinition?: RoleDefinition

}

export interface ResourceOperation extends Entity {

	    resourceName?: string

	    actionName?: string

	    description?: string

}

export interface TelecomExpenseManagementPartner extends Entity {

	    displayName?: string

	    url?: string

	    appAuthorized?: boolean

	    enabled?: boolean

	    lastConnectionDateTime?: string

}

export interface RemoteAssistancePartner extends Entity {

	    displayName?: string

	    onboardingUrl?: string

	    onboardingStatus?: RemoteAssistanceOnboardingStatus

	    lastConnectionDateTime?: string

}

export interface WindowsInformationProtectionAppLearningSummary extends Entity {

	    applicationName?: string

	    applicationType?: ApplicationType

	    deviceCount?: number

}

export interface WindowsMalwareInformation extends Entity {

	    displayName?: string

	    additionalInformationUrl?: string

	    severity?: WindowsMalwareSeverity

	    category?: WindowsMalwareCategory

	    lastDetectionDateTime?: string

	    windowsDevicesProtectionState?: WindowsProtectionState[]

}

export interface DeviceAppManagement extends Entity {

	    windowsStoreForBusinessLastSuccessfulSyncDateTime?: string

	    isEnabledForWindowsStoreForBusiness?: boolean

	    windowsStoreForBusinessLanguage?: string

	    windowsStoreForBusinessLastCompletedApplicationSyncTime?: string

	    windowsManagementApp?: WindowsManagementApp

	    mobileApps?: MobileApp[]

	    mobileAppCategories?: MobileAppCategory[]

	    enterpriseCodeSigningCertificates?: EnterpriseCodeSigningCertificate[]

	    iosLobAppProvisioningConfigurations?: IosLobAppProvisioningConfiguration[]

	    symantecCodeSigningCertificate?: SymantecCodeSigningCertificate

	    mobileAppConfigurations?: ManagedDeviceMobileAppConfiguration[]

	    sideLoadingKeys?: SideLoadingKey[]

	    managedAppPolicies?: ManagedAppPolicy[]

	    iosManagedAppProtections?: IosManagedAppProtection[]

	    androidManagedAppProtections?: AndroidManagedAppProtection[]

	    defaultManagedAppProtections?: DefaultManagedAppProtection[]

	    targetedManagedAppConfigurations?: TargetedManagedAppConfiguration[]

	    mdmWindowsInformationProtectionPolicies?: MdmWindowsInformationProtectionPolicy[]

	    windowsInformationProtectionPolicies?: WindowsInformationProtectionPolicy[]

	    managedAppRegistrations?: ManagedAppRegistration[]

	    managedAppStatuses?: ManagedAppStatus[]

	    managedEBooks?: ManagedEBook[]

}

export interface WindowsManagementApp extends Entity {

	    availableVersion?: string

	    healthSummary?: WindowsManagementAppHealthSummary

	    healthStates?: WindowsManagementAppHealthState[]

}

export interface MobileApp extends Entity {

	    displayName?: string

	    description?: string

	    publisher?: string

	    largeIcon?: MimeContent

	    createdDateTime?: string

	    lastModifiedDateTime?: string

	    isFeatured?: boolean

	    privacyInformationUrl?: string

	    informationUrl?: string

	    owner?: string

	    developer?: string

	    notes?: string

	    uploadState?: number

	    categories?: MobileAppCategory[]

	    groupAssignments?: MobileAppGroupAssignment[]

	    installSummary?: MobileAppInstallSummary

	    deviceStatuses?: MobileAppInstallStatus[]

	    userStatuses?: UserAppInstallStatus[]

}

export interface MobileAppCategory extends Entity {

	    displayName?: string

	    lastModifiedDateTime?: string

}

export interface EnterpriseCodeSigningCertificate extends Entity {

	    content?: number

	    status?: CertificateStatus

	    subjectName?: string

	    subject?: string

	    issuerName?: string

	    issuer?: string

	    expirationDateTime?: string

	    uploadDateTime?: string

}

export interface IosLobAppProvisioningConfiguration extends Entity {

	    expirationDateTime?: string

	    payloadFileName?: string

	    payload?: number

	    createdDateTime?: string

	    description?: string

	    lastModifiedDateTime?: string

	    displayName?: string

	    version?: number

	    groupAssignments?: MobileAppProvisioningConfigGroupAssignment[]

	    deviceStatuses?: ManagedDeviceMobileAppConfigurationDeviceStatus[]

	    userStatuses?: ManagedDeviceMobileAppConfigurationUserStatus[]

}

export interface SymantecCodeSigningCertificate extends Entity {

	    content?: number

	    status?: CertificateStatus

	    password?: string

	    subjectName?: string

	    subject?: string

	    issuerName?: string

	    issuer?: string

	    expirationDateTime?: string

	    uploadDateTime?: string

}

export interface ManagedDeviceMobileAppConfiguration extends Entity {

	    targetedMobileApps?: string[]

	    createdDateTime?: string

	    description?: string

	    lastModifiedDateTime?: string

	    displayName?: string

	    version?: number

	    groupAssignments?: MdmAppConfigGroupAssignment[]

	    deviceStatuses?: ManagedDeviceMobileAppConfigurationDeviceStatus[]

	    userStatuses?: ManagedDeviceMobileAppConfigurationUserStatus[]

	    deviceStatusSummary?: ManagedDeviceMobileAppConfigurationDeviceSummary

	    userStatusSummary?: ManagedDeviceMobileAppConfigurationUserSummary

}

export interface SideLoadingKey extends Entity {

	    value?: string

	    displayName?: string

	    description?: string

	    totalActivation?: number

	    lastUpdatedDateTime?: string

}

export interface ManagedAppPolicy extends Entity {

	    displayName?: string

	    description?: string

	    createdDateTime?: string

	    lastModifiedDateTime?: string

	    version?: string

}

export interface ManagedAppProtection extends ManagedAppPolicy {

	    periodOfflineBeforeAccessCheck?: string

	    periodOnlineBeforeAccessCheck?: string

	    allowedInboundDataTransferSources?: ManagedAppDataTransferLevel

	    allowedOutboundDataTransferDestinations?: ManagedAppDataTransferLevel

	    organizationalCredentialsRequired?: boolean

	    allowedOutboundClipboardSharingLevel?: ManagedAppClipboardSharingLevel

	    dataBackupBlocked?: boolean

	    deviceComplianceRequired?: boolean

	    managedBrowserToOpenLinksRequired?: boolean

	    saveAsBlocked?: boolean

	    periodOfflineBeforeWipeIsEnforced?: string

	    pinRequired?: boolean

	    maximumPinRetries?: number

	    simplePinBlocked?: boolean

	    minimumPinLength?: number

	    pinCharacterSet?: ManagedAppPinCharacterSet

	    allowedDataStorageLocations?: ManagedAppDataStorageLocation[]

	    contactSyncBlocked?: boolean

	    printBlocked?: boolean

	    fingerprintBlocked?: boolean

	    disableAppPinIfDevicePinIsSet?: boolean

	    minimumRequiredOsVersion?: string

	    minimumWarningOsVersion?: string

	    minimumRequiredAppVersion?: string

	    minimumWarningAppVersion?: string

}

export interface TargetedManagedAppProtection extends ManagedAppProtection {

	    targetedSecurityGroupsCount?: number

	    targetedSecurityGroupIds?: string[]

	    targetedSecurityGroups?: DirectoryObject[]

}

export interface IosManagedAppProtection extends TargetedManagedAppProtection {

	    appDataEncryptionType?: ManagedAppDataEncryptionType

	    minimumRequiredSdkVersion?: string

	    deployedAppCount?: number

	    mobileAppIdentifierDeployments?: MobileAppIdentifierDeployment[]

	    deploymentSummary?: ManagedAppPolicyDeploymentSummary

}

export interface AndroidManagedAppProtection extends TargetedManagedAppProtection {

	    screenCaptureBlocked?: boolean

	    disableAppEncryptionIfDeviceEncryptionIsEnabled?: boolean

	    encryptAppData?: boolean

	    deployedAppCount?: number

	    mobileAppIdentifierDeployments?: MobileAppIdentifierDeployment[]

	    deploymentSummary?: ManagedAppPolicyDeploymentSummary

}

export interface DefaultManagedAppProtection extends ManagedAppProtection {

	    appDataEncryptionType?: ManagedAppDataEncryptionType

	    screenCaptureBlocked?: boolean

	    encryptAppData?: boolean

	    disableAppEncryptionIfDeviceEncryptionIsEnabled?: boolean

	    minimumRequiredSdkVersion?: string

	    customSettings?: KeyValuePair[]

	    deployedAppCount?: number

	    mobileAppIdentifierDeployments?: MobileAppIdentifierDeployment[]

	    deploymentSummary?: ManagedAppPolicyDeploymentSummary

}

export interface ManagedAppConfiguration extends ManagedAppPolicy {

	    customSettings?: KeyValuePair[]

}

export interface TargetedManagedAppConfiguration extends ManagedAppConfiguration {

	    deployedAppCount?: number

	    targetedSecurityGroupIds?: string[]

	    targetedSecurityGroupsCount?: number

	    mobileAppIdentifierDeployments?: MobileAppIdentifierDeployment[]

	    deploymentSummary?: ManagedAppPolicyDeploymentSummary

	    targetedSecurityGroups?: DirectoryObject[]

}

export interface WindowsInformationProtection extends ManagedAppPolicy {

	    enforcementLevel?: WindowsInformationProtectionEnforcementLevel

	    enterpriseDomain?: string

	    enterpriseProtectedDomainNames?: WindowsInformationProtectionResourceCollection[]

	    protectionUnderLockConfigRequired?: boolean

	    dataRecoveryCertificate?: WindowsInformationProtectionDataRecoveryCertificate

	    revokeOnUnenrollDisabled?: boolean

	    rightsManagementServicesTemplateId?: string

	    azureRightsManagementServicesAllowed?: boolean

	    iconsVisible?: boolean

	    allowedApps?: WindowsInformationProtectionApp[]

	    exemptApps?: WindowsInformationProtectionApp[]

	    enterpriseNetworkDomainNames?: WindowsInformationProtectionResourceCollection[]

	    enterpriseCloudResources?: WindowsInformationProtectionCloudResourceCollection[]

	    enterpriseIPRanges?: WindowsInformationProtectionIPRangeCollection[]

	    enterpriseIPRangesAreAuthoritative?: boolean

	    enterpriseProxyServers?: WindowsInformationProtectionResourceCollection[]

	    enterpriseInternalProxyServers?: WindowsInformationProtectionResourceCollection[]

	    enterpriseProxyServersAreAuthoritative?: boolean

	    neutralDomainResources?: WindowsInformationProtectionResourceCollection[]

	    indexingEncryptedStoresOrItemsBlocked?: boolean

	    smbAutoEncryptedFileExtensions?: WindowsInformationProtectionResourceCollection[]

	    targetedSecurityGroupIds?: string[]

	    allowedAppLockerFiles?: WindowsInformationProtectionAppLockerFile[]

	    exemptAppLockerFiles?: WindowsInformationProtectionAppLockerFile[]

}

export interface MdmWindowsInformationProtectionPolicy extends WindowsInformationProtection {

}

export interface WindowsInformationProtectionPolicy extends WindowsInformationProtection {

	    revokeOnMdmHandoffDisabled?: boolean

	    mdmEnrollmentUrl?: string

	    windowsHelloForBusinessBlocked?: boolean

	    pinMinimumLength?: number

	    pinUppercaseLetters?: WindowsInformationProtectionPinCharacterRequirements

	    pinLowercaseLetters?: WindowsInformationProtectionPinCharacterRequirements

	    pinSpecialCharacters?: WindowsInformationProtectionPinCharacterRequirements

	    pinExpirationDays?: number

	    numberOfPastPinsRemembered?: number

	    passwordMaximumAttemptCount?: number

	    minutesOfInactivityBeforeDeviceLock?: number

}

export interface ManagedAppStatus extends Entity {

	    displayName?: string

	    version?: string

}

export interface ManagedEBook extends Entity {

	    displayName?: string

	    description?: string

	    publisher?: string

	    publishedDateTime?: string

	    largeCover?: MimeContent

	    createdDateTime?: string

	    lastModifiedDateTime?: string

	    informationUrl?: string

	    privacyInformationUrl?: string

	    groupAssignments?: EBookGroupAssignment[]

	    installSummary?: EBookInstallSummary

	    deviceStates?: DeviceInstallState[]

	    userStateSummary?: UserInstallStateSummary[]

}

export interface MdmAppConfigGroupAssignment extends Entity {

	    appConfiguration?: string

	    targetGroupId?: string

}

export interface ManagedDeviceMobileAppConfigurationDeviceStatus extends Entity {

	    deviceDisplayName?: string

	    userName?: string

	    deviceModel?: string

	    platform?: number

	    complianceGracePeriodExpirationDateTime?: string

	    status?: ComplianceStatus

	    lastReportedDateTime?: string

	    userPrincipalName?: string

}

export interface ManagedDeviceMobileAppConfigurationUserStatus extends Entity {

	    userDisplayName?: string

	    devicesCount?: number

	    status?: ComplianceStatus

	    lastReportedDateTime?: string

	    userPrincipalName?: string

}

export interface ManagedDeviceMobileAppConfigurationDeviceSummary extends Entity {

	    pendingCount?: number

	    successCount?: number

	    errorCount?: number

	    failedCount?: number

	    lastUpdateDateTime?: string

	    configurationVersion?: number

}

export interface ManagedDeviceMobileAppConfigurationUserSummary extends Entity {

	    pendingCount?: number

	    successCount?: number

	    errorCount?: number

	    failedCount?: number

	    lastUpdateDateTime?: string

	    configurationVersion?: number

}

export interface IosMobileAppConfiguration extends ManagedDeviceMobileAppConfiguration {

	    settingXml?: string

	    settings?: AppConfigurationSettingItem[]

}

export interface MobileAppGroupAssignment extends Entity {

	    targetGroupId?: string

	    vpnConfigurationId?: string

	    installIntent?: AppInstallIntent

	    app?: MobileApp

}

export interface MobileAppInstallSummary extends Entity {

	    installedDeviceCount?: number

	    failedDeviceCount?: number

	    notInstalledDeviceCount?: number

	    installedUserCount?: number

	    failedUserCount?: number

	    notInstalledUserCount?: number

}

export interface MobileAppInstallStatus extends Entity {

	    deviceName?: string

	    deviceId?: string

	    lastSyncDateTime?: string

	    mobileAppInstallStatusValue?: ResultantAppState

	    errorCode?: number

	    osVersion?: string

	    osDescription?: string

	    userName?: string

	    userPrincipalName?: string

	    app?: MobileApp

}

export interface UserAppInstallStatus extends Entity {

	    userName?: string

	    userPrincipalName?: string

	    installedDeviceCount?: number

	    failedDeviceCount?: number

	    notInstalledDeviceCount?: number

	    app?: MobileApp

	    deviceStatuses?: MobileAppInstallStatus[]

}

export interface MobileAppContentFile extends Entity {

	    azureStorageUri?: string

	    isCommitted?: boolean

	    createdDateTime?: string

	    name?: string

	    size?: number

	    sizeEncrypted?: number

	    azureStorageUriExpirationDateTime?: string

	    manifest?: number

	    uploadState?: MobileAppContentFileUploadState

}

export interface MobileAppProvisioningConfigGroupAssignment extends Entity {

	    targetGroupId?: string

}

export interface MobileAppVppGroupAssignment extends MobileAppGroupAssignment {

	    useDeviceLicensing?: boolean

}

export interface OfficeSuiteApp extends MobileApp {

	    autoAcceptEula?: boolean

	    productIds?: OfficeProductId[]

	    excludedApps?: ExcludedApps

	    useSharedComputerActivation?: boolean

	    updateChannel?: OfficeUpdateChannel

	    officePlatformArchitecture?: WindowsArchitecture

	    localesToInstall?: string[]

}

export interface ManagedApp extends MobileApp {

	    appAvailability?: ManagedAppAvailability

	    version?: string

}

export interface ManagedAndroidStoreApp extends ManagedApp {

	    packageId?: string

}

export interface ManagedIOSStoreApp extends ManagedApp {

	    bundleId?: string

}

export interface ManagedMobileLobApp extends ManagedApp {

	    committedContentVersion?: string

	    fileName?: string

	    size?: number

	    identityVersion?: string

	    contentVersions?: MobileAppContent[]

}

export interface MobileAppContent extends Entity {

	    files?: MobileAppContentFile[]

}

export interface ManagedAndroidLobApp extends ManagedMobileLobApp {

	    identityName?: string

	    minimumSupportedOperatingSystem?: AndroidMinimumOperatingSystem

}

export interface ManagedIOSLobApp extends ManagedMobileLobApp {

	    bundleId?: string

	    applicableDeviceType?: IosDeviceType

	    minimumSupportedOperatingSystem?: IosMinimumOperatingSystem

	    expirationDateTime?: string

}

export interface MobileLobApp extends MobileApp {

	    committedContentVersion?: string

	    fileName?: string

	    size?: number

	    identityVersion?: string

	    contentVersions?: MobileAppContent[]

}

export interface WindowsMobileMSI extends MobileLobApp {

	    commandLine?: string

	    productCode?: string

}

export interface WindowsPhone81AppX extends MobileLobApp {

	    applicableArchitectures?: WindowsArchitecture

	    identityName?: string

	    identityPublisherHash?: string

	    identityResourceIdentifier?: string

	    minimumSupportedOperatingSystem?: WindowsMinimumOperatingSystem

	    phoneProductIdentifier?: string

	    phonePublisherId?: string

}

export interface WindowsPhone81AppXBundle extends WindowsPhone81AppX {

	    appXPackageInformationList?: WindowsPackageInformation[]

}

export interface WindowsUniversalAppX extends MobileLobApp {

	    applicableArchitectures?: WindowsArchitecture

	    applicableDeviceTypes?: WindowsDeviceType

	    identityName?: string

	    identityPublisherHash?: string

	    identityResourceIdentifier?: string

	    isBundle?: boolean

	    minimumSupportedOperatingSystem?: WindowsMinimumOperatingSystem

}

export interface WindowsAppX extends MobileLobApp {

	    applicableArchitectures?: WindowsArchitecture

	    identityName?: string

	    identityPublisherHash?: string

	    identityResourceIdentifier?: string

	    isBundle?: boolean

	    minimumSupportedOperatingSystem?: WindowsMinimumOperatingSystem

}

export interface WindowsPhoneXAP extends MobileLobApp {

	    minimumSupportedOperatingSystem?: WindowsMinimumOperatingSystem

	    productIdentifier?: string

}

export interface AndroidLobApp extends MobileLobApp {

	    identityName?: string

	    minimumSupportedOperatingSystem?: AndroidMinimumOperatingSystem

}

export interface IosLobApp extends MobileLobApp {

	    bundleId?: string

	    applicableDeviceType?: IosDeviceType

	    minimumSupportedOperatingSystem?: IosMinimumOperatingSystem

	    expirationDateTime?: string

}

export interface AndroidForWorkApp extends MobileApp {

	    appIdentifier?: string

	    usedLicenseCount?: number

	    totalLicenseCount?: number

	    appStoreUrl?: string

}

export interface WebApp extends MobileApp {

	    appUrl?: string

	    useManagedBrowser?: boolean

}

export interface WindowsPhone81StoreApp extends MobileApp {

	    appStoreUrl?: string

}

export interface WindowsStoreApp extends MobileApp {

	    appStoreUrl?: string

}

export interface AndroidStoreApp extends MobileApp {

	    appIdentifier?: string

	    appStoreUrl?: string

	    minimumSupportedOperatingSystem?: AndroidMinimumOperatingSystem

}

export interface IosVppApp extends MobileApp {

	    usedLicenseCount?: number

	    totalLicenseCount?: number

	    releaseDateTime?: string

	    appStoreUrl?: string

	    licensingType?: VppLicensingType

	    applicableDeviceType?: IosDeviceType

	    vppToken?: AppleVolumePurchaseProgramToken

}

export interface IosStoreApp extends MobileApp {

	    bundleId?: string

	    appStoreUrl?: string

	    applicableDeviceType?: IosDeviceType

	    minimumSupportedOperatingSystem?: IosMinimumOperatingSystem

}

export interface WindowsStoreForBusinessApp extends MobileApp {

	    usedLicenseCount?: number

	    totalLicenseCount?: number

	    productKey?: string

	    licenseType?: WindowsStoreForBusinessLicenseType

	    packageIdentityName?: string

}

export interface AppReportingOverviewStatus extends Entity {

}

export interface AndroidForWorkMobileAppConfiguration extends ManagedDeviceMobileAppConfiguration {

	    packageName?: string

	    payloadJson?: string

	    permissionActions?: AndroidPermissionAction[]

}

export interface TermsAndConditionsGroupAssignment extends Entity {

	    targetGroupId?: string

	    termsAndConditions?: TermsAndConditions

}

export interface TermsAndConditionsAcceptanceStatus extends Entity {

	    userDisplayName?: string

	    acceptedVersion?: number

	    acceptedDateTime?: string

	    termsAndConditions?: TermsAndConditions

}

export interface ImportedDeviceIdentityResult extends ImportedDeviceIdentity {

	    status?: boolean

}

export interface ImportedAppleDeviceIdentityResult extends ImportedAppleDeviceIdentity {

	    status?: boolean

}

export interface DepEnrollmentProfile extends EnrollmentProfile {

	    supervisedModeEnabled?: boolean

	    supportDepartment?: string

	    passCodeDisabled?: boolean

	    isMandatory?: boolean

	    locationDisabled?: boolean

	    supportPhoneNumber?: string

	    iTunesPairingMode?: ITunesPairingMode

	    profileRemovalDisabled?: boolean

	    managementCertificates?: ManagementCertificateWithThumbprint[]

	    restoreBlocked?: boolean

	    restoreFromAndroidDisabled?: boolean

	    appleIdDisabled?: boolean

	    termsAndConditionsDisabled?: boolean

	    touchIdDisabled?: boolean

	    applePayDisabled?: boolean

	    zoomDisabled?: boolean

	    siriDisabled?: boolean

	    diagnosticsDisabled?: boolean

	    macOSRegistrationDisabled?: boolean

	    macOSFileVaultDisabled?: boolean

	    awaitDeviceConfiguredConfirmation?: boolean

	    sharedIPadMaximumUserCount?: number

	    enableSharedIPad?: boolean

	    enableAuthenticationViaCompanyPortal?: boolean

}

export interface DeviceConfigurationState extends Entity {

	    settingStates?: DeviceConfigurationSettingState[]

	    displayName?: string

	    version?: number

	    platformType?: PolicyPlatformType

	    state?: ComplianceStatus

	    settingCount?: number

}

export interface DeviceCompliancePolicyState extends Entity {

	    settingStates?: DeviceCompliancePolicySettingState[]

	    displayName?: string

	    version?: number

	    platformType?: PolicyPlatformType

	    state?: ComplianceStatus

	    settingCount?: number

}

export interface DeviceManagementScriptGroupAssignment extends Entity {

	    targetGroupId?: string

}

export interface DeviceManagementScriptRunSummary extends Entity {

	    successDeviceCount?: number

	    errorDeviceCount?: number

	    successUserCount?: number

	    errorUserCount?: number

}

export interface DeviceManagementScriptDeviceState extends Entity {

	    runState?: RunState

	    resultMessage?: string

	    lastStateUpdateDateTime?: string

	    errorCode?: number

	    errorDescription?: string

	    managedDevice?: ManagedDevice

}

export interface DeviceManagementScriptUserState extends Entity {

	    successDeviceCount?: number

	    errorDeviceCount?: number

	    userPrincipalName?: string

	    deviceRunStates?: DeviceManagementScriptDeviceState[]

}

export interface WindowsManagedDevice extends ManagedDevice {

	    windowsProtectionState?: WindowsProtectionState

}

export interface WindowsProtectionState extends Entity {

	    malwareProtectionEnabled?: boolean

	    deviceState?: WindowsDeviceHealthState

	    realTimeProtectionEnabled?: boolean

	    networkInspectionSystemEnabled?: boolean

	    quickScanOverdue?: boolean

	    fullScanOverdue?: boolean

	    signatureUpdateOverdue?: boolean

	    rebootRequired?: boolean

	    fullScanRequired?: boolean

	    engineVersion?: string

	    signatureVersion?: string

	    antiMalwareVersion?: string

	    lastQuickScanDateTime?: string

	    lastFullScanDateTime?: string

	    lastQuickScanSignatureVersion?: string

	    lastFullScanSignatureVersion?: string

	    lastReportedDateTime?: string

	    detectedMalwareState?: WindowsDeviceMalwareState[]

}

export interface WindowsDeviceMalwareState extends Entity {

	    displayName?: string

	    additionalInformationUrl?: string

	    severity?: WindowsMalwareSeverity

	    catetgory?: WindowsMalwareCategory

	    executionState?: WindowsMalwareExecutionState

	    state?: WindowsMalwareState

	    initialDetectionDateTime?: string

	    lastStateChangeDateTime?: string

	    detectionCount?: number

}

export interface WindowsManagementAppHealthSummary extends Entity {

	    healthyDeviceCount?: number

	    unhealthyDeviceCount?: number

	    unknownDeviceCount?: number

}

export interface WindowsManagementAppHealthState extends Entity {

	    healthState?: HealthState

	    installedVersion?: string

	    lastCheckInDateTime?: string

	    deviceName?: string

	    deviceOSVersion?: string

}

export interface ReportRoot extends Entity {

}

export interface DeviceConfigurationAssignment extends Entity {

	    deviceConfiguration?: DeviceConfiguration

}

export interface DeviceConfigurationGroupAssignment extends DeviceConfigurationAssignment {

	    targetGroupId?: string

	    excludeGroup?: boolean

}

export interface DeviceConfigurationDeviceStatus extends Entity {

	    deviceDisplayName?: string

	    userName?: string

	    deviceModel?: string

	    platform?: number

	    complianceGracePeriodExpirationDateTime?: string

	    status?: ComplianceStatus

	    lastReportedDateTime?: string

	    userPrincipalName?: string

}

export interface DeviceConfigurationUserStatus extends Entity {

	    userDisplayName?: string

	    devicesCount?: number

	    status?: ComplianceStatus

	    lastReportedDateTime?: string

	    userPrincipalName?: string

}

export interface DeviceConfigurationDeviceOverview extends Entity {

	    pendingCount?: number

	    notApplicableCount?: number

	    successCount?: number

	    errorCount?: number

	    failedCount?: number

	    lastUpdateDateTime?: string

	    configurationVersion?: number

}

export interface DeviceConfigurationUserOverview extends Entity {

	    pendingCount?: number

	    notApplicableCount?: number

	    successCount?: number

	    errorCount?: number

	    failedCount?: number

	    lastUpdateDateTime?: string

	    configurationVersion?: number

}

export interface SettingStateDeviceSummary extends Entity {

	    settingName?: string

	    instancePath?: string

	    unknownDeviceCount?: number

	    notApplicableDeviceCount?: number

	    compliantDeviceCount?: number

	    remediatedDeviceCount?: number

	    nonCompliantDeviceCount?: number

	    errorDeviceCount?: number

	    conflictDeviceCount?: number

}

export interface DeviceCompliancePolicyAssignment extends Entity {

	    deviceCompliancePolicy?: DeviceCompliancePolicy

}

export interface DeviceCompliancePolicyGroupAssignment extends DeviceCompliancePolicyAssignment {

	    targetGroupId?: string

	    excludeGroup?: boolean

}

export interface DeviceComplianceScheduledActionForRule extends Entity {

	    ruleName?: string

	    scheduledActionConfigurations?: DeviceComplianceActionItem[]

}

export interface DeviceComplianceDeviceStatus extends Entity {

	    deviceDisplayName?: string

	    userName?: string

	    deviceModel?: string

	    platform?: number

	    complianceGracePeriodExpirationDateTime?: string

	    status?: ComplianceStatus

	    lastReportedDateTime?: string

	    userPrincipalName?: string

}

export interface DeviceComplianceUserStatus extends Entity {

	    userDisplayName?: string

	    devicesCount?: number

	    status?: ComplianceStatus

	    lastReportedDateTime?: string

	    userPrincipalName?: string

}

export interface DeviceComplianceDeviceOverview extends Entity {

	    pendingCount?: number

	    notApplicableCount?: number

	    successCount?: number

	    errorCount?: number

	    failedCount?: number

	    lastUpdateDateTime?: string

	    configurationVersion?: number

}

export interface DeviceComplianceUserOverview extends Entity {

	    pendingCount?: number

	    notApplicableCount?: number

	    successCount?: number

	    errorCount?: number

	    failedCount?: number

	    lastUpdateDateTime?: string

	    configurationVersion?: number

}

export interface DeviceComplianceActionItem extends Entity {

	    gracePeriodHours?: number

	    actionType?: DeviceComplianceActionType

	    notificationTemplateId?: string

	    notificationMessageCCList?: string[]

	    notificationMessageTemplate?: NotificationMessageTemplate

}

export interface WindowsPrivacyDataAccessControlItem extends DeviceConfigurationAssignment {

	    accessLevel?: WindowsPrivacyDataAccessLevel

	    dataCategory?: WindowsPrivacyDataCategory

	    appPackageFamilyName?: string

	    appDisplayName?: string

}

export interface AndroidForWorkEasEmailProfileBase extends DeviceConfiguration {

	    authenticationMethod?: EasAuthenticationMethod

	    durationOfEmailToSync?: EmailSyncDuration

	    emailAddressSource?: UserEmailSource

	    hostName?: string

	    requireSsl?: boolean

	    usernameSource?: AndroidUsernameSource

	    identityCertificate?: AndroidForWorkCertificateProfileBase

}

export interface AndroidForWorkCertificateProfileBase extends DeviceConfiguration {

	    renewalThresholdPercentage?: number

	    subjectNameFormat?: SubjectNameFormat

	    subjectAlternativeNameType?: SubjectAlternativeNameType

	    certificateValidityPeriodValue?: number

	    certificateValidityPeriodScale?: CertificateValidityPeriodScale

	    extendedKeyUsages?: ExtendedKeyUsage[]

	    rootCertificate?: AndroidForWorkTrustedRootCertificate

}

export interface AndroidForWorkTrustedRootCertificate extends DeviceConfiguration {

	    trustedRootCertificate?: number

	    certFileName?: string

}

export interface AndroidForWorkPkcsCertificateProfile extends AndroidForWorkCertificateProfileBase {

	    certificationAuthority?: string

	    certificationAuthorityName?: string

	    certificateTemplateName?: string

	    subjectAlternativeNameFormatString?: string

}

export interface AndroidForWorkScepCertificateProfile extends AndroidForWorkCertificateProfileBase {

	    scepServerUrls?: string[]

	    subjectNameFormatString?: string

	    keyUsage?: KeyUsages

	    keySize?: KeySize

	    hashAlgorithm?: HashAlgorithms

	    subjectAlternativeNameFormatString?: string

	    managedDeviceCertificateStates?: ManagedDeviceCertificateState[]

}

export interface ManagedDeviceCertificateState extends Entity {

	    devicePlatform?: DevicePlatformType

	    certificateKeyUsage?: KeyUsages

	    certificateProfileDisplayName?: string

	    deviceDisplayName?: string

	    userDisplayName?: string

	    serverUrl?: string

	    certificateExpirationDateTime?: string

	    lastCertificateStateChangeDateTime?: string

	    certificateIssuer?: string

	    certificateThumbprint?: string

	    certificateSerialNumber?: string

	    certificateKeyLength?: number

	    enhancedKeyUsage?: string

}

export interface AndroidForWorkGmailEasConfiguration extends AndroidForWorkEasEmailProfileBase {

}

export interface AndroidForWorkNineWorkEasConfiguration extends AndroidForWorkEasEmailProfileBase {

	    syncCalendar?: boolean

	    syncContacts?: boolean

	    syncTasks?: boolean

}

export interface AndroidCertificateProfileBase extends DeviceConfiguration {

	    renewalThresholdPercentage?: number

	    subjectNameFormat?: SubjectNameFormat

	    subjectAlternativeNameType?: SubjectAlternativeNameType

	    certificateValidityPeriodValue?: number

	    certificateValidityPeriodScale?: CertificateValidityPeriodScale

	    extendedKeyUsages?: ExtendedKeyUsage[]

	    rootCertificate?: AndroidTrustedRootCertificate

}

export interface AndroidTrustedRootCertificate extends DeviceConfiguration {

	    trustedRootCertificate?: number

	    certFileName?: string

}

export interface AndroidPkcsCertificateProfile extends AndroidCertificateProfileBase {

	    certificationAuthority?: string

	    certificationAuthorityName?: string

	    certificateTemplateName?: string

	    subjectAlternativeNameFormatString?: string

}

export interface AndroidScepCertificateProfile extends AndroidCertificateProfileBase {

	    scepServerUrls?: string[]

	    subjectNameFormatString?: string

	    keyUsage?: KeyUsages

	    keySize?: KeySize

	    hashAlgorithm?: HashAlgorithms

	    subjectAlternativeNameFormatString?: string

	    managedDeviceCertificateStates?: ManagedDeviceCertificateState[]

}

export interface AndroidCustomConfiguration extends DeviceConfiguration {

	    omaSettings?: OmaSetting[]

}

export interface AndroidEasEmailProfileConfiguration extends DeviceConfiguration {

	    accountName?: string

	    authenticationMethod?: EasAuthenticationMethod

	    syncCalendar?: boolean

	    syncContacts?: boolean

	    syncTasks?: boolean

	    syncNotes?: boolean

	    durationOfEmailToSync?: EmailSyncDuration

	    emailAddressSource?: UserEmailSource

	    emailSyncSchedule?: EmailSyncSchedule

	    hostName?: string

	    requireSmime?: boolean

	    requireSsl?: boolean

	    usernameSource?: AndroidUsernameSource

	    identityCertificate?: AndroidCertificateProfileBase

	    smimeSigningCertificate?: AndroidCertificateProfileBase

}

export interface AndroidForWorkCustomConfiguration extends DeviceConfiguration {

	    omaSettings?: OmaSetting[]

}

export interface AndroidForWorkWiFiConfiguration extends DeviceConfiguration {

	    networkName?: string

	    ssid?: string

	    connectAutomatically?: boolean

	    connectWhenNetworkNameIsHidden?: boolean

	    wiFiSecurityType?: AndroidWiFiSecurityType

}

export interface AndroidForWorkEnterpriseWiFiConfiguration extends AndroidForWorkWiFiConfiguration {

	    eapType?: AndroidEapType

	    authenticationMethod?: WiFiAuthenticationMethod

	    nonEapAuthenticationMethodForEapTtls?: NonEapAuthenticationMethodForEapTtlsType

	    nonEapAuthenticationMethodForPeap?: NonEapAuthenticationMethodForPeap

	    enableOuterIdentityPrivacy?: string

	    rootCertificateForServerValidation?: AndroidForWorkTrustedRootCertificate

	    identityCertificateForClientAuthentication?: AndroidForWorkCertificateProfileBase

}

export interface AndroidForWorkGeneralDeviceConfiguration extends DeviceConfiguration {

	    passwordBlockFingerprintUnlock?: boolean

	    passwordBlockTrustAgents?: boolean

	    passwordExpirationDays?: number

	    passwordMinimumLength?: number

	    passwordMinutesOfInactivityBeforeScreenTimeout?: number

	    passwordPreviousPasswordBlockCount?: number

	    passwordSignInFailureCountBeforeFactoryReset?: number

	    passwordRequiredType?: AndroidForWorkRequiredPasswordType

	    workProfileDataSharingType?: AndroidForWorkCrossProfileDataSharingType

	    workProfileBlockNotificationsWhileDeviceLocked?: boolean

	    workProfileDefaultAppPermissionPolicy?: AndroidForWorkDefaultAppPermissionPolicyType

	    workProfilePasswordBlockFingerprintUnlock?: boolean

	    workProfilePasswordBlockTrustAgents?: boolean

	    workProfilePasswordExpirationDays?: number

	    workProfilePasswordMinimumLength?: number

	    workProfilePasswordMinutesOfInactivityBeforeScreenTimeout?: number

	    workProfilePasswordPreviousPasswordBlockCount?: number

	    workProfilePasswordSignInFailureCountBeforeFactoryReset?: number

	    workProfilePasswordRequiredType?: AndroidForWorkRequiredPasswordType

	    workProfileRequirePassword?: boolean

}

export interface AndroidForWorkVpnConfiguration extends DeviceConfiguration {

	    connectionName?: string

	    connectionType?: AndroidForWorkVpnConnectionType

	    role?: string

	    realm?: string

	    servers?: VpnServer[]

	    fingerprint?: string

	    customData?: KeyValue[]

	    enableSplitTunneling?: boolean

	    authenticationMethod?: VpnAuthenticationMethod

	    identityCertificate?: AndroidForWorkCertificateProfileBase

}

export interface AndroidGeneralDeviceConfiguration extends DeviceConfiguration {

	    appsBlockClipboardSharing?: boolean

	    appsBlockCopyPaste?: boolean

	    appsBlockYouTube?: boolean

	    bluetoothBlocked?: boolean

	    cameraBlocked?: boolean

	    cellularBlockDataRoaming?: boolean

	    cellularBlockMessaging?: boolean

	    cellularBlockVoiceRoaming?: boolean

	    cellularBlockWiFiTethering?: boolean

	    compliantAppsList?: AppListItem[]

	    compliantAppListType?: AppListType

	    diagnosticDataBlockSubmission?: boolean

	    locationServicesBlocked?: boolean

	    googleAccountBlockAutoSync?: boolean

	    googlePlayStoreBlocked?: boolean

	    kioskModeBlockSleepButton?: boolean

	    kioskModeBlockVolumeButtons?: boolean

	    kioskModeManagedApps?: AppListItem[]

	    nfcBlocked?: boolean

	    passwordBlockFingerprintUnlock?: boolean

	    passwordBlockTrustAgents?: boolean

	    passwordExpirationDays?: number

	    passwordMinimumLength?: number

	    passwordMinutesOfInactivityBeforeScreenTimeout?: number

	    passwordPreviousPasswordBlockCount?: number

	    passwordSignInFailureCountBeforeFactoryReset?: number

	    passwordRequiredType?: AndroidRequiredPasswordType

	    passwordRequired?: boolean

	    powerOffBlocked?: boolean

	    factoryResetBlocked?: boolean

	    screenCaptureBlocked?: boolean

	    deviceSharingAllowed?: boolean

	    storageBlockGoogleBackup?: boolean

	    storageBlockRemovableStorage?: boolean

	    storageRequireDeviceEncryption?: boolean

	    storageRequireRemovableStorageEncryption?: boolean

	    voiceAssistantBlocked?: boolean

	    voiceDialingBlocked?: boolean

	    webBrowserBlockPopups?: boolean

	    webBrowserBlockAutofill?: boolean

	    webBrowserBlockJavaScript?: boolean

	    webBrowserBlocked?: boolean

	    webBrowserCookieSettings?: WebBrowserCookieSettings

	    wiFiBlocked?: boolean

	    appsInstallAllowList?: AppListItem[]

	    appsLaunchBlockList?: AppListItem[]

	    appsHideList?: AppListItem[]

}

export interface AndroidVpnConfiguration extends DeviceConfiguration {

	    connectionName?: string

	    connectionType?: AndroidVpnConnectionType

	    role?: string

	    realm?: string

	    servers?: VpnServer[]

	    fingerprint?: string

	    customData?: KeyValue[]

	    authenticationMethod?: VpnAuthenticationMethod

	    identityCertificate?: AndroidCertificateProfileBase

}

export interface AndroidWiFiConfiguration extends DeviceConfiguration {

	    networkName?: string

	    ssid?: string

	    connectAutomatically?: boolean

	    connectWhenNetworkNameIsHidden?: boolean

	    wiFiSecurityType?: AndroidWiFiSecurityType

}

export interface AndroidEnterpriseWiFiConfiguration extends AndroidWiFiConfiguration {

	    eapType?: AndroidEapType

	    authenticationMethod?: WiFiAuthenticationMethod

	    innerAuthenticationProtocolForEapTtls?: NonEapAuthenticationMethodForEapTtlsType

	    innerAuthenticationProtocolForPeap?: NonEapAuthenticationMethodForPeap

	    outerIdentityPrivacyTemporaryValue?: string

	    rootCertificateForServerValidation?: AndroidTrustedRootCertificate

	    identityCertificateForClientAuthentication?: AndroidCertificateProfileBase

}

export interface IosCertificateProfileBase extends DeviceConfiguration {

	    renewalThresholdPercentage?: number

	    subjectNameFormat?: AppleSubjectNameFormat

	    subjectAlternativeNameType?: SubjectAlternativeNameType

	    certificateValidityPeriodValue?: number

	    certificateValidityPeriodScale?: CertificateValidityPeriodScale

}

export interface IosPkcsCertificateProfile extends IosCertificateProfileBase {

	    certificationAuthority?: string

	    certificationAuthorityName?: string

	    certificateTemplateName?: string

	    subjectAlternativeNameFormatString?: string

}

export interface IosScepCertificateProfile extends IosCertificateProfileBase {

	    scepServerUrls?: string[]

	    subjectNameFormatString?: string

	    keyUsage?: KeyUsages

	    keySize?: KeySize

	    extendedKeyUsages?: ExtendedKeyUsage[]

	    subjectAlternativeNameFormatString?: string

	    rootCertificate?: IosTrustedRootCertificate

	    managedDeviceCertificateStates?: ManagedDeviceCertificateState[]

}

export interface IosTrustedRootCertificate extends DeviceConfiguration {

	    trustedRootCertificate?: number

	    certFileName?: string

}

export interface IosCustomConfiguration extends DeviceConfiguration {

	    payloadName?: string

	    payloadFileName?: string

	    payload?: number

}

export interface IosEasEmailProfileConfiguration extends DeviceConfiguration {

	    accountName?: string

	    authenticationMethod?: EasAuthenticationMethod

	    blockMovingMessagesToOtherEmailAccounts?: boolean

	    blockSendingEmailFromThirdPartyApps?: boolean

	    blockSyncingRecentlyUsedEmailAddresses?: boolean

	    durationOfEmailToSync?: EmailSyncDuration

	    emailAddressSource?: UserEmailSource

	    hostName?: string

	    requireSmime?: boolean

	    requireSsl?: boolean

	    usernameSource?: UserEmailSource

	    identityCertificate?: IosCertificateProfileBase

	    smimeSigningCertificate?: IosCertificateProfileBase

}

export interface IosEduDeviceConfiguration extends DeviceConfiguration {

	    teacherCertificateSettings?: IosEduCertificateSettings

	    studentCertificateSettings?: IosEduCertificateSettings

	    deviceCertificateSettings?: IosEduCertificateSettings

}

export interface IosEducationDeviceConfiguration extends DeviceConfiguration {

}

export interface IosGeneralDeviceConfiguration extends DeviceConfiguration {

	    accountBlockModification?: boolean

	    activationLockAllowWhenSupervised?: boolean

	    airDropBlocked?: boolean

	    airDropForceUnmanagedDropTarget?: boolean

	    airPlayForcePairingPasswordForOutgoingRequests?: boolean

	    appleWatchBlockPairing?: boolean

	    appleWatchForceWristDetection?: boolean

	    appleNewsBlocked?: boolean

	    appsSingleAppModeList?: AppListItem[]

	    appsVisibilityList?: AppListItem[]

	    appsVisibilityListType?: AppListType

	    appStoreBlockAutomaticDownloads?: boolean

	    appStoreBlocked?: boolean

	    appStoreBlockInAppPurchases?: boolean

	    appStoreBlockUIAppInstallation?: boolean

	    appStoreRequirePassword?: boolean

	    bluetoothBlockModification?: boolean

	    cameraBlocked?: boolean

	    cellularBlockDataRoaming?: boolean

	    cellularBlockGlobalBackgroundFetchWhileRoaming?: boolean

	    cellularBlockPerAppDataModification?: boolean

	    cellularBlockPersonalHotspot?: boolean

	    cellularBlockVoiceRoaming?: boolean

	    certificatesBlockUntrustedTlsCertificates?: boolean

	    classroomAppBlockRemoteScreenObservation?: boolean

	    classroomAppForceUnpromptedScreenObservation?: boolean

	    compliantAppsList?: AppListItem[]

	    compliantAppListType?: AppListType

	    configurationProfileBlockChanges?: boolean

	    definitionLookupBlocked?: boolean

	    deviceBlockEnableRestrictions?: boolean

	    deviceBlockEraseContentAndSettings?: boolean

	    deviceBlockNameModification?: boolean

	    diagnosticDataBlockSubmission?: boolean

	    diagnosticDataBlockSubmissionModification?: boolean

	    documentsBlockManagedDocumentsInUnmanagedApps?: boolean

	    documentsBlockUnmanagedDocumentsInManagedApps?: boolean

	    emailInDomainSuffixes?: string[]

	    enterpriseAppBlockTrust?: boolean

	    enterpriseAppBlockTrustModification?: boolean

	    faceTimeBlocked?: boolean

	    findMyFriendsBlocked?: boolean

	    gamingBlockGameCenterFriends?: boolean

	    gamingBlockMultiplayer?: boolean

	    gameCenterBlocked?: boolean

	    hostPairingBlocked?: boolean

	    iBooksStoreBlocked?: boolean

	    iBooksStoreBlockErotica?: boolean

	    iCloudBlockActivityContinuation?: boolean

	    iCloudBlockBackup?: boolean

	    iCloudBlockDocumentSync?: boolean

	    iCloudBlockManagedAppsSync?: boolean

	    iCloudBlockPhotoLibrary?: boolean

	    iCloudBlockPhotoStreamSync?: boolean

	    iCloudBlockSharedPhotoStream?: boolean

	    iCloudRequireEncryptedBackup?: boolean

	    iTunesBlockExplicitContent?: boolean

	    iTunesBlockMusicService?: boolean

	    iTunesBlockRadio?: boolean

	    keyboardBlockAutoCorrect?: boolean

	    keyboardBlockDictation?: boolean

	    keyboardBlockPredictive?: boolean

	    keyboardBlockShortcuts?: boolean

	    keyboardBlockSpellCheck?: boolean

	    kioskModeAllowAssistiveSpeak?: boolean

	    kioskModeAllowAssistiveTouchSettings?: boolean

	    kioskModeAllowAutoLock?: boolean

	    kioskModeAllowColorInversionSettings?: boolean

	    kioskModeAllowRingerSwitch?: boolean

	    kioskModeAllowScreenRotation?: boolean

	    kioskModeAllowSleepButton?: boolean

	    kioskModeAllowTouchscreen?: boolean

	    kioskModeAllowVoiceOverSettings?: boolean

	    kioskModeAllowVolumeButtons?: boolean

	    kioskModeAllowZoomSettings?: boolean

	    kioskModeAppStoreUrl?: string

	    kioskModeRequireAssistiveTouch?: boolean

	    kioskModeRequireColorInversion?: boolean

	    kioskModeRequireMonoAudio?: boolean

	    kioskModeRequireVoiceOver?: boolean

	    kioskModeRequireZoom?: boolean

	    kioskModeManagedAppId?: string

	    lockScreenBlockControlCenter?: boolean

	    lockScreenBlockNotificationView?: boolean

	    lockScreenBlockPassbook?: boolean

	    lockScreenBlockTodayView?: boolean

	    mediaContentRatingAustralia?: MediaContentRatingAustralia

	    mediaContentRatingCanada?: MediaContentRatingCanada

	    mediaContentRatingFrance?: MediaContentRatingFrance

	    mediaContentRatingGermany?: MediaContentRatingGermany

	    mediaContentRatingIreland?: MediaContentRatingIreland

	    mediaContentRatingJapan?: MediaContentRatingJapan

	    mediaContentRatingNewZealand?: MediaContentRatingNewZealand

	    mediaContentRatingUnitedKingdom?: MediaContentRatingUnitedKingdom

	    mediaContentRatingUnitedStates?: MediaContentRatingUnitedStates

	    networkUsageRules?: IosNetworkUsageRule[]

	    mediaContentRatingApps?: RatingAppsType

	    messagesBlocked?: boolean

	    notificationsBlockSettingsModification?: boolean

	    passcodeBlockFingerprintUnlock?: boolean

	    passcodeBlockFingerprintModification?: boolean

	    passcodeBlockModification?: boolean

	    passcodeBlockSimple?: boolean

	    passcodeExpirationDays?: number

	    passcodeMinimumLength?: number

	    passcodeMinutesOfInactivityBeforeLock?: number

	    passcodeMinutesOfInactivityBeforeScreenTimeout?: number

	    passcodeMinimumCharacterSetCount?: number

	    passcodePreviousPasscodeBlockCount?: number

	    passcodeSignInFailureCountBeforeWipe?: number

	    passcodeRequiredType?: RequiredPasswordType

	    passcodeRequired?: boolean

	    podcastsBlocked?: boolean

	    safariBlockAutofill?: boolean

	    safariBlockJavaScript?: boolean

	    safariBlockPopups?: boolean

	    safariBlocked?: boolean

	    safariCookieSettings?: WebBrowserCookieSettings

	    safariManagedDomains?: string[]

	    safariPasswordAutoFillDomains?: string[]

	    safariRequireFraudWarning?: boolean

	    screenCaptureBlocked?: boolean

	    siriBlocked?: boolean

	    siriBlockedWhenLocked?: boolean

	    siriBlockUserGeneratedContent?: boolean

	    siriRequireProfanityFilter?: boolean

	    spotlightBlockInternetResults?: boolean

	    voiceDialingBlocked?: boolean

	    wallpaperBlockModification?: boolean

	    wiFiConnectOnlyToConfiguredNetworks?: boolean

}

export interface IosUpdateConfiguration extends DeviceConfiguration {

	    isEnabled?: boolean

	    activeHoursStart?: string

	    activeHoursEnd?: string

	    scheduledInstallDays?: DayOfWeek[]

	    utcTimeOffsetInMinutes?: number

}

export interface IosWiFiConfiguration extends DeviceConfiguration {

	    networkName?: string

	    ssid?: string

	    connectAutomatically?: boolean

	    connectWhenNetworkNameIsHidden?: boolean

	    wiFiSecurityType?: WiFiSecurityType

	    proxySettings?: WiFiProxySetting

	    proxyManualAddress?: string

	    proxyManualPort?: number

	    proxyAutomaticConfigurationUrl?: string

	    preSharedKey?: string

}

export interface IosEnterpriseWiFiConfiguration extends IosWiFiConfiguration {

	    eapType?: EapType

	    eapFastConfiguration?: EapFastConfiguration

	    trustedServerCertificateNames?: string[]

	    authenticationMethod?: WiFiAuthenticationMethod

	    innerAuthenticationProtocolForEapTtls?: NonEapAuthenticationMethodForEapTtlsType

	    outerIdentityPrivacyTemporaryValue?: string

	    rootCertificatesForServerValidation?: IosTrustedRootCertificate[]

	    identityCertificateForClientAuthentication?: IosCertificateProfileBase

}

export interface MacOSCertificateProfileBase extends DeviceConfiguration {

	    renewalThresholdPercentage?: number

	    subjectNameFormat?: AppleSubjectNameFormat

	    subjectAlternativeNameType?: SubjectAlternativeNameType

	    certificateValidityPeriodValue?: number

	    certificateValidityPeriodScale?: CertificateValidityPeriodScale

}

export interface MacOSScepCertificateProfile extends MacOSCertificateProfileBase {

	    scepServerUrls?: string[]

	    subjectNameFormatString?: string

	    keyUsage?: KeyUsages

	    keySize?: KeySize

	    hashAlgorithm?: HashAlgorithms

	    extendedKeyUsages?: ExtendedKeyUsage[]

	    subjectAlternativeNameFormatString?: string

	    rootCertificate?: MacOSTrustedRootCertificate

	    managedDeviceCertificateStates?: ManagedDeviceCertificateState[]

}

export interface MacOSTrustedRootCertificate extends DeviceConfiguration {

	    trustedRootCertificate?: number

	    certFileName?: string

}

export interface MacOSCustomConfiguration extends DeviceConfiguration {

	    payloadName?: string

	    payloadFileName?: string

	    payload?: number

}

export interface MacOSGeneralDeviceConfiguration extends DeviceConfiguration {

	    compliantAppsList?: AppListItem[]

	    compliantAppListType?: AppListType

	    emailInDomainSuffixes?: string[]

	    passwordBlockSimple?: boolean

	    passwordExpirationDays?: number

	    passwordMinimumCharacterSetCount?: number

	    passwordMinimumLength?: number

	    passwordMinutesOfInactivityBeforeLock?: number

	    passwordMinutesOfInactivityBeforeScreenTimeout?: number

	    passwordPreviousPasswordBlockCount?: number

	    passwordRequiredType?: RequiredPasswordType

	    passwordRequired?: boolean

}

export interface MacOSWiFiConfiguration extends DeviceConfiguration {

	    networkName?: string

	    ssid?: string

	    connectAutomatically?: boolean

	    connectWhenNetworkNameIsHidden?: boolean

	    wiFiSecurityType?: WiFiSecurityType

	    proxySettings?: WiFiProxySetting

	    proxyManualAddress?: string

	    proxyManualPort?: number

	    proxyAutomaticConfigurationUrl?: string

	    preSharedKey?: string

}

export interface MacOSEnterpriseWiFiConfiguration extends MacOSWiFiConfiguration {

	    eapType?: EapType

	    eapFastConfiguration?: EapFastConfiguration

	    trustedServerCertificateNames?: string[]

	    authenticationMethod?: WiFiAuthenticationMethod

	    innerAuthenticationProtocolForEapTtls?: NonEapAuthenticationMethodForEapTtlsType

	    outerIdentityPrivacyTemporaryValue?: string

	    rootCertificateForServerValidation?: MacOSTrustedRootCertificate

	    identityCertificateForClientAuthentication?: MacOSCertificateProfileBase

}

export interface AppleDeviceFeaturesConfigurationBase extends DeviceConfiguration {

	    airPrintDestinations?: AirPrintDestination[]

}

export interface IosDeviceFeaturesConfiguration extends AppleDeviceFeaturesConfigurationBase {

	    assetTagTemplate?: string

	    contentFilterSettings?: IosWebContentFilterBase

	    lockScreenFootnote?: string

	    homeScreenDockIcons?: IosHomeScreenItem[]

	    homeScreenPages?: IosHomeScreenPage[]

	    notificationSettings?: IosNotificationSettings[]

}

export interface MacOSDeviceFeaturesConfiguration extends AppleDeviceFeaturesConfigurationBase {

}

export interface AppleVpnConfiguration extends DeviceConfiguration {

	    connectionName?: string

	    connectionType?: AppleVpnConnectionType

	    loginGroupOrDomain?: string

	    role?: string

	    realm?: string

	    server?: VpnServer

	    identifier?: string

	    customData?: KeyValue[]

	    enableSplitTunneling?: boolean

	    authenticationMethod?: VpnAuthenticationMethod

	    enablePerApp?: boolean

	    safariDomains?: string[]

	    onDemandRules?: VpnOnDemandRule[]

	    proxyServer?: VpnProxyServer

}

export interface IosVpnConfiguration extends AppleVpnConfiguration {

	    identityCertificate?: IosCertificateProfileBase

}

export interface MacOSVpnConfiguration extends AppleVpnConfiguration {

	    identityCertificate?: MacOSCertificateProfileBase

}

export interface Windows10EndpointProtectionConfiguration extends DeviceConfiguration {

	    smartScreenEnableInShell?: boolean

	    smartScreenBlockOverrideForFiles?: boolean

	    applicationGuardEnabled?: boolean

	    applicationGuardBlockFileTransfer?: ApplicationGuardBlockFileTransferType

	    applicationGuardBlockNonEnterpriseContent?: boolean

	    applicationGuardAllowPersistence?: boolean

	    applicationGuardForceAuditing?: boolean

	    bitLockerEnableStorageCardEncryptionOnMobile?: boolean

	    bitLockerEncryptDevice?: boolean

	    bitLockerSystemDrivePolicy?: BitLockerSystemDrivePolicy

	    bitLockerFixedDrivePolicy?: BitLockerFixedDrivePolicy

	    bitLockerRemovableDrivePolicy?: BitLockerRemovableDrivePolicy

}

export interface Windows10GeneralConfiguration extends DeviceConfiguration {

	    enterpriseCloudPrintDiscoveryEndPoint?: string

	    enterpriseCloudPrintOAuthAuthority?: string

	    enterpriseCloudPrintOAuthClientIdentifier?: string

	    enterpriseCloudPrintResourceIdentifier?: string

	    enterpriseCloudPrintDiscoveryMaxLimit?: number

	    enterpriseCloudPrintMopriaDiscoveryResourceIdentifier?: string

	    searchBlockDiacritics?: boolean

	    searchDisableAutoLanguageDetection?: boolean

	    searchDisableIndexingEncryptedItems?: boolean

	    searchEnableRemoteQueries?: boolean

	    searchDisableUseLocation?: boolean

	    searchDisableIndexerBackoff?: boolean

	    searchDisableIndexingRemovableDrive?: boolean

	    searchEnableAutomaticIndexSizeManangement?: boolean

	    smartScreenEnableAppInstallControl?: boolean

	    personalizationDesktopImageUrl?: string

	    personalizationLockScreenImageUrl?: string

	    bluetoothAllowedServices?: string[]

	    bluetoothBlockAdvertising?: boolean

	    bluetoothBlockDiscoverableMode?: boolean

	    bluetoothBlockPrePairing?: boolean

	    bluetoothDeviceName?: string

	    edgeBlockAutofill?: boolean

	    edgeBlocked?: boolean

	    edgeCookiePolicy?: EdgeCookiePolicy

	    edgeBlockDeveloperTools?: boolean

	    edgeBlockSendingDoNotTrackHeader?: boolean

	    edgeBlockExtensions?: boolean

	    edgeBlockInPrivateBrowsing?: boolean

	    edgeBlockJavaScript?: boolean

	    edgeBlockPasswordManager?: boolean

	    edgeBlockAddressBarDropdown?: boolean

	    edgeBlockCompatibilityList?: boolean

	    edgeClearBrowsingDataOnExit?: boolean

	    edgeAllowStartPagesModification?: boolean

	    edgeDisableFirstRunPage?: boolean

	    edgeBlockLiveTileDataCollection?: boolean

	    edgeSyncFavoritesWithInternetExplorer?: boolean

	    cellularBlockDataWhenRoaming?: boolean

	    cellularBlockVpn?: boolean

	    cellularBlockVpnWhenRoaming?: boolean

	    cellularData?: ConfigurationUsage

	    defenderBlockEndUserAccess?: boolean

	    defenderDaysBeforeDeletingQuarantinedMalware?: number

	    defenderDetectedMalwareActions?: DefenderDetectedMalwareActions

	    defenderSystemScanSchedule?: WeeklySchedule

	    defenderFilesAndFoldersToExclude?: string[]

	    defenderFileExtensionsToExclude?: string[]

	    defenderScanMaxCpu?: number

	    defenderMonitorFileActivity?: DefenderMonitorFileActivity

	    defenderPotentiallyUnwantedAppAction?: DefenderPotentiallyUnwantedAppAction

	    defenderProcessesToExclude?: string[]

	    defenderPromptForSampleSubmission?: DefenderPromptForSampleSubmission

	    defenderRequireBehaviorMonitoring?: boolean

	    defenderRequireCloudProtection?: boolean

	    defenderRequireNetworkInspectionSystem?: boolean

	    defenderRequireRealTimeMonitoring?: boolean

	    defenderScanArchiveFiles?: boolean

	    defenderScanDownloads?: boolean

	    defenderScanNetworkFiles?: boolean

	    defenderScanIncomingMail?: boolean

	    defenderScanMappedNetworkDrivesDuringFullScan?: boolean

	    defenderScanRemovableDrivesDuringFullScan?: boolean

	    defenderScanScriptsLoadedInInternetExplorer?: boolean

	    defenderSignatureUpdateIntervalInHours?: number

	    defenderScanType?: DefenderScanType

	    defenderScheduledScanTime?: string

	    defenderScheduledQuickScanTime?: string

	    lockScreenAllowTimeoutConfiguration?: boolean

	    lockScreenBlockActionCenterNotifications?: boolean

	    lockScreenBlockCortana?: boolean

	    lockScreenBlockToastNotifications?: boolean

	    lockScreenTimeoutInSeconds?: number

	    passwordBlockSimple?: boolean

	    passwordExpirationDays?: number

	    passwordMinimumLength?: number

	    passwordMinutesOfInactivityBeforeScreenTimeout?: number

	    passwordMinimumCharacterSetCount?: number

	    passwordPreviousPasswordBlockCount?: number

	    passwordRequired?: boolean

	    passwordRequireWhenResumeFromIdleState?: boolean

	    passwordRequiredType?: RequiredPasswordType

	    passwordSignInFailureCountBeforeFactoryReset?: number

	    privacyAdvertisingId?: StateManagementSetting

	    privacyAutoAcceptPairingAndConsentPrompts?: boolean

	    privacyBlockInputPersonalization?: boolean

	    startBlockUnpinningAppsFromTaskbar?: boolean

	    startMenuAppListVisibility?: WindowsStartMenuAppListVisibilityType

	    startMenuHideChangeAccountSettings?: boolean

	    startMenuHideFrequentlyUsedApps?: boolean

	    startMenuHideHibernate?: boolean

	    startMenuHideLock?: boolean

	    startMenuHidePowerButton?: boolean

	    startMenuHideRecentJumpLists?: boolean

	    startMenuHideRecentlyAddedApps?: boolean

	    startMenuHideRestartOptions?: boolean

	    startMenuHideShutDown?: boolean

	    startMenuHideSignOut?: boolean

	    startMenuHideSleep?: boolean

	    startMenuHideSwitchAccount?: boolean

	    startMenuHideUserTile?: boolean

	    startMenuLayoutEdgeAssetsXml?: number

	    startMenuLayoutXml?: number

	    startMenuMode?: WindowsStartMenuModeType

	    startMenuPinnedFolderDocuments?: VisibilitySetting

	    startMenuPinnedFolderDownloads?: VisibilitySetting

	    startMenuPinnedFolderFileExplorer?: VisibilitySetting

	    startMenuPinnedFolderHomeGroup?: VisibilitySetting

	    startMenuPinnedFolderMusic?: VisibilitySetting

	    startMenuPinnedFolderNetwork?: VisibilitySetting

	    startMenuPinnedFolderPersonalFolder?: VisibilitySetting

	    startMenuPinnedFolderPictures?: VisibilitySetting

	    startMenuPinnedFolderSettings?: VisibilitySetting

	    startMenuPinnedFolderVideos?: VisibilitySetting

	    settingsBlockSettingsApp?: boolean

	    settingsBlockSystemPage?: boolean

	    settingsBlockDevicesPage?: boolean

	    settingsBlockNetworkInternetPage?: boolean

	    settingsBlockPersonalizationPage?: boolean

	    settingsBlockAccountsPage?: boolean

	    settingsBlockTimeLanguagePage?: boolean

	    settingsBlockEaseOfAccessPage?: boolean

	    settingsBlockPrivacyPage?: boolean

	    settingsBlockUpdateSecurityPage?: boolean

	    settingsBlockAppsPage?: boolean

	    settingsBlockGamingPage?: boolean

	    windowsSpotlightBlockConsumerSpecificFeatures?: boolean

	    windowsSpotlightBlocked?: boolean

	    windowsSpotlightBlockOnActionCenter?: boolean

	    windowsSpotlightBlockTailoredExperiences?: boolean

	    windowsSpotlightBlockThirdPartyNotifications?: boolean

	    windowsSpotlightBlockWelcomeExperience?: boolean

	    windowsSpotlightBlockWindowsTips?: boolean

	    windowsSpotlightConfigureOnLockScreen?: WindowsSpotlightEnablementSettings

	    networkProxyApplySettingsDeviceWide?: boolean

	    networkProxyDisableAutoDetect?: boolean

	    networkProxyAutomaticConfigurationUrl?: string

	    networkProxyServer?: Windows10NetworkProxyServer

	    accountsBlockAddingNonMicrosoftAccountEmail?: boolean

	    antiTheftModeBlocked?: boolean

	    bluetoothBlocked?: boolean

	    cameraBlocked?: boolean

	    connectedDevicesServiceBlocked?: boolean

	    certificatesBlockManualRootCertificateInstallation?: boolean

	    copyPasteBlocked?: boolean

	    cortanaBlocked?: boolean

	    deviceManagementBlockFactoryResetOnMobile?: boolean

	    deviceManagementBlockManualUnenroll?: boolean

	    diagnosticsDataSubmissionMode?: DiagnosticDataSubmissionMode

	    oneDriveDisableFileSync?: boolean

	    safeSearchFilter?: SafeSearchFilterType

	    edgeBlockPopups?: boolean

	    edgeBlockSearchSuggestions?: boolean

	    edgeBlockSendingIntranetTrafficToInternetExplorer?: boolean

	    edgeRequireSmartScreen?: boolean

	    edgeEnterpriseModeSiteListLocation?: string

	    edgeFirstRunUrl?: string

	    edgeSearchEngine?: EdgeSearchEngineBase

	    edgeHomepageUrls?: string[]

	    edgeBlockAccessToAboutFlags?: boolean

	    smartScreenBlockPromptOverride?: boolean

	    smartScreenBlockPromptOverrideForFiles?: boolean

	    webRtcBlockLocalhostIpAddress?: boolean

	    internetSharingBlocked?: boolean

	    settingsBlockAddProvisioningPackage?: boolean

	    settingsBlockRemoveProvisioningPackage?: boolean

	    settingsBlockChangeSystemTime?: boolean

	    settingsBlockEditDeviceName?: boolean

	    settingsBlockChangeRegion?: boolean

	    settingsBlockChangeLanguage?: boolean

	    settingsBlockChangePowerSleep?: boolean

	    locationServicesBlocked?: boolean

	    microsoftAccountBlocked?: boolean

	    microsoftAccountBlockSettingsSync?: boolean

	    nfcBlocked?: boolean

	    resetProtectionModeBlocked?: boolean

	    screenCaptureBlocked?: boolean

	    storageBlockRemovableStorage?: boolean

	    storageRequireMobileDeviceEncryption?: boolean

	    usbBlocked?: boolean

	    voiceRecordingBlocked?: boolean

	    wiFiBlockAutomaticConnectHotspots?: boolean

	    wiFiBlocked?: boolean

	    wiFiBlockManualConfiguration?: boolean

	    wiFiScanInterval?: number

	    wirelessDisplayBlockProjectionToThisDevice?: boolean

	    wirelessDisplayBlockUserInputFromReceiver?: boolean

	    wirelessDisplayRequirePinForPairing?: boolean

	    windowsStoreBlocked?: boolean

	    appsAllowTrustedAppsSideloading?: StateManagementSetting

	    windowsStoreBlockAutoUpdate?: boolean

	    developerUnlockSetting?: StateManagementSetting

	    sharedUserAppDataAllowed?: boolean

	    appsBlockWindowsStoreOriginatedApps?: boolean

	    windowsStoreEnablePrivateStoreOnly?: boolean

	    storageRestrictAppDataToSystemVolume?: boolean

	    storageRestrictAppInstallToSystemVolume?: boolean

	    gameDvrBlocked?: boolean

	    experienceBlockDeviceDiscovery?: boolean

	    experienceBlockErrorDialogWhenNoSIM?: boolean

	    experienceBlockTaskSwitcher?: boolean

	    logonBlockFastUserSwitching?: boolean

	    privacyAccessControls?: WindowsPrivacyDataAccessControlItem[]

}

export interface Windows10CustomConfiguration extends DeviceConfiguration {

	    omaSettings?: OmaSetting[]

}

export interface Windows10EasEmailProfileConfiguration extends DeviceConfiguration {

	    accountName?: string

	    syncCalendar?: boolean

	    syncContacts?: boolean

	    syncTasks?: boolean

	    durationOfEmailToSync?: EmailSyncDuration

	    emailAddressSource?: UserEmailSource

	    emailSyncSchedule?: EmailSyncSchedule

	    hostName?: string

	    requireSsl?: boolean

	    usernameSource?: UserEmailSource

}

export interface Windows10EnterpriseModernAppManagementConfiguration extends DeviceConfiguration {

	    uninstallBuiltInApps?: boolean

}

export interface SharedPCConfiguration extends DeviceConfiguration {

	    accountManagerPolicy?: SharedPCAccountManagerPolicy

	    allowedAccounts?: SharedPCAllowedAccountType

	    allowLocalStorage?: boolean

	    disableAccountManager?: boolean

	    disableEduPolicies?: boolean

	    disablePowerPolicies?: boolean

	    disableSignInOnResume?: boolean

	    enabled?: boolean

	    idleTimeBeforeSleepInSeconds?: number

	    kioskAppDisplayName?: string

	    kioskAppUserModelId?: string

	    maintenanceStartTime?: string

}

export interface Windows10SecureAssessmentConfiguration extends DeviceConfiguration {

	    launchUri?: string

	    configurationAccount?: string

	    allowPrinting?: boolean

	    allowScreenCapture?: boolean

	    allowTextSuggestion?: boolean

}

export interface Windows81WifiImportConfiguration extends DeviceConfiguration {

	    payloadFileName?: string

	    profileName?: string

	    payload?: number

}

export interface WindowsCertificateProfileBase extends DeviceConfiguration {

	    renewalThresholdPercentage?: number

	    keyStorageProvider?: KeyStorageProviderOption

	    subjectNameFormat?: SubjectNameFormat

	    subjectAlternativeNameType?: SubjectAlternativeNameType

	    certificateValidityPeriodValue?: number

	    certificateValidityPeriodScale?: CertificateValidityPeriodScale

}

export interface Windows10CertificateProfileBase extends WindowsCertificateProfileBase {

}

export interface Windows10PkcsCertificateProfile extends Windows10CertificateProfileBase {

	    certificationAuthority?: string

	    certificationAuthorityName?: string

	    certificateTemplateName?: string

	    subjectAlternativeNameFormatString?: string

}

export interface Windows81CertificateProfileBase extends WindowsCertificateProfileBase {

	    extendedKeyUsages?: ExtendedKeyUsage[]

}

export interface Windows81SCEPCertificateProfile extends Windows81CertificateProfileBase {

	    scepServerUrls?: string[]

	    subjectNameFormatString?: string

	    keyUsage?: KeyUsages

	    keySize?: KeySize

	    hashAlgorithm?: HashAlgorithms

	    subjectAlternativeNameFormatString?: string

	    rootCertificate?: Windows81TrustedRootCertificate

	    managedDeviceCertificateStates?: ManagedDeviceCertificateState[]

}

export interface Windows81TrustedRootCertificate extends DeviceConfiguration {

	    trustedRootCertificate?: number

	    certFileName?: string

	    destinationStore?: CertificateDestinationStore

}

export interface WindowsPhone81CustomConfiguration extends DeviceConfiguration {

	    omaSettings?: OmaSetting[]

}

export interface WindowsPhone81TrustedRootCertificate extends DeviceConfiguration {

	    trustedRootCertificate?: number

	    certFileName?: string

}

export interface WindowsPhoneEASEmailProfileConfiguration extends DeviceConfiguration {

	    accountName?: string

	    applyOnlyToWindowsPhone81?: boolean

	    syncCalendar?: boolean

	    syncContacts?: boolean

	    syncTasks?: boolean

	    durationOfEmailToSync?: EmailSyncDuration

	    emailAddressSource?: UserEmailSource

	    emailSyncSchedule?: EmailSyncSchedule

	    hostName?: string

	    requireSsl?: boolean

	    usernameSource?: UserEmailSource

}

export interface WindowsUpdateForBusinessConfiguration extends DeviceConfiguration {

	    deliveryOptimizationMode?: WindowsDeliveryOptimizationMode

	    prereleaseFeatures?: PrereleaseFeatures

	    automaticUpdateMode?: AutomaticUpdateMode

	    microsoftUpdateServiceAllowed?: boolean

	    driversExcluded?: boolean

	    installationSchedule?: WindowsUpdateInstallScheduleType

	    qualityUpdatesDeferralPeriodInDays?: number

	    featureUpdatesDeferralPeriodInDays?: number

	    qualityUpdatesPaused?: boolean

	    featureUpdatesPaused?: boolean

	    qualityUpdatesPauseExpiryDateTime?: string

	    featureUpdatesPauseExpiryDateTime?: string

	    businessReadyUpdatesOnly?: WindowsUpdateType

	    previewBuildSetting?: WindowsUpdateInsiderBuildControl

}

export interface WindowsVpnConfiguration extends DeviceConfiguration {

	    connectionName?: string

	    servers?: VpnServer[]

	    customXml?: number

}

export interface Windows10VpnConfiguration extends WindowsVpnConfiguration {

	    connectionType?: Windows10VpnConnectionType

	    enableSplitTunneling?: boolean

	    authenticationMethod?: Windows10VpnAuthenticationMethod

	    rememberUserCredentials?: boolean

	    enableConditionalAccess?: boolean

	    enableSingleSignOnWithAlternateCertificate?: boolean

	    singleSignOnEku?: ExtendedKeyUsage

	    singleSignOnIssuerHash?: string

	    eapXml?: number

	    proxyServer?: Windows10VpnProxyServer

	    associatedApps?: Windows10AssociatedApps[]

	    onlyAssociatedAppsCanUseConnection?: boolean

	    windowsInformationProtectionDomain?: string

	    trafficRules?: VpnTrafficRule[]

	    routes?: VpnRoute[]

	    dnsRules?: VpnDnsRule[]

	    identityCertificate?: WindowsCertificateProfileBase

}

export interface Windows81VpnConfiguration extends WindowsVpnConfiguration {

	    applyOnlyToWindows81?: boolean

	    connectionType?: WindowsVpnConnectionType

	    loginGroupOrDomain?: string

	    enableSplitTunneling?: boolean

	    proxyServer?: Windows81VpnProxyServer

}

export interface WindowsPhone81VpnConfiguration extends Windows81VpnConfiguration {

	    bypassVpnOnCompanyWifi?: boolean

	    bypassVpnOnHomeWifi?: boolean

	    authenticationMethod?: VpnAuthenticationMethod

	    rememberUserCredentials?: boolean

	    dnsSuffixSearchList?: string[]

	    identityCertificate?: WindowsPhone81CertificateProfileBase

}

export interface WindowsPhone81CertificateProfileBase extends DeviceConfiguration {

	    renewalThresholdPercentage?: number

	    keyStorageProvider?: KeyStorageProviderOption

	    subjectNameFormat?: SubjectNameFormat

	    subjectAlternativeNameType?: SubjectAlternativeNameType

	    certificateValidityPeriodValue?: number

	    certificateValidityPeriodScale?: CertificateValidityPeriodScale

	    extendedKeyUsages?: ExtendedKeyUsage[]

}

export interface WindowsPhone81SCEPCertificateProfile extends WindowsPhone81CertificateProfileBase {

	    scepServerUrls?: string[]

	    subjectNameFormatString?: string

	    keyUsage?: KeyUsages

	    keySize?: KeySize

	    hashAlgorithm?: HashAlgorithms

	    subjectAlternativeNameFormatString?: string

	    rootCertificate?: WindowsPhone81TrustedRootCertificate

	    managedDeviceCertificateStates?: ManagedDeviceCertificateState[]

}

export interface Windows81GeneralConfiguration extends DeviceConfiguration {

	    accountsBlockAddingNonMicrosoftAccountEmail?: boolean

	    applyOnlyToWindows81?: boolean

	    browserBlockAutofill?: boolean

	    browserBlockAutomaticDetectionOfIntranetSites?: boolean

	    browserBlockEnterpriseModeAccess?: boolean

	    browserBlockJavaScript?: boolean

	    browserBlockPlugins?: boolean

	    browserBlockPopups?: boolean

	    browserBlockSendingDoNotTrackHeader?: boolean

	    browserBlockSingleWordEntryOnIntranetSites?: boolean

	    browserRequireSmartScreen?: boolean

	    browserEnterpriseModeSiteListLocation?: string

	    browserInternetSecurityLevel?: InternetSiteSecurityLevel

	    browserIntranetSecurityLevel?: SiteSecurityLevel

	    browserLoggingReportLocation?: string

	    browserRequireHighSecurityForRestrictedSites?: boolean

	    browserRequireFirewall?: boolean

	    browserRequireFraudWarning?: boolean

	    browserTrustedSitesSecurityLevel?: SiteSecurityLevel

	    cellularBlockDataRoaming?: boolean

	    diagnosticsBlockDataSubmission?: boolean

	    passwordBlockPicturePasswordAndPin?: boolean

	    passwordExpirationDays?: number

	    passwordMinimumLength?: number

	    passwordMinutesOfInactivityBeforeScreenTimeout?: number

	    passwordMinimumCharacterSetCount?: number

	    passwordPreviousPasswordBlockCount?: number

	    passwordRequiredType?: RequiredPasswordType

	    passwordSignInFailureCountBeforeFactoryReset?: number

	    storageRequireDeviceEncryption?: boolean

	    minimumAutoInstallClassification?: UpdateClassification

	    updatesRequireAutomaticUpdates?: boolean

	    userAccountControlSettings?: WindowsUserAccountControlSettings

	    workFoldersUrl?: string

}

export interface WindowsPhone81GeneralConfiguration extends DeviceConfiguration {

	    applyOnlyToWindowsPhone81?: boolean

	    appsBlockCopyPaste?: boolean

	    bluetoothBlocked?: boolean

	    cameraBlocked?: boolean

	    cellularBlockWifiTethering?: boolean

	    compliantAppsList?: AppListItem[]

	    compliantAppListType?: AppListType

	    diagnosticDataBlockSubmission?: boolean

	    emailBlockAddingAccounts?: boolean

	    locationServicesBlocked?: boolean

	    microsoftAccountBlocked?: boolean

	    nfcBlocked?: boolean

	    passwordBlockSimple?: boolean

	    passwordExpirationDays?: number

	    passwordMinimumLength?: number

	    passwordMinutesOfInactivityBeforeScreenTimeout?: number

	    passwordMinimumCharacterSetCount?: number

	    passwordPreviousPasswordBlockCount?: number

	    passwordSignInFailureCountBeforeFactoryReset?: number

	    passwordRequiredType?: RequiredPasswordType

	    passwordRequired?: boolean

	    screenCaptureBlocked?: boolean

	    storageBlockRemovableStorage?: boolean

	    storageRequireEncryption?: boolean

	    webBrowserBlocked?: boolean

	    wifiBlocked?: boolean

	    wifiBlockAutomaticConnectHotspots?: boolean

	    wifiBlockHotspotReporting?: boolean

	    windowsStoreBlocked?: boolean

}

export interface Windows10TeamGeneralConfiguration extends DeviceConfiguration {

	    azureOperationalInsightsBlockTelemetry?: boolean

	    azureOperationalInsightsWorkspaceId?: string

	    azureOperationalInsightsWorkspaceKey?: string

	    connectAppBlockAutoLaunch?: boolean

	    deviceAccountBlockExchangeServices?: boolean

	    deviceAccountEmailAddress?: string

	    deviceAccountExchangeServerAddress?: string

	    deviceAccountRequirePasswordRotation?: boolean

	    deviceAccountSessionInitiationProtocolAddress?: string

	    maintenanceWindowBlocked?: boolean

	    maintenanceWindowDurationInHours?: number

	    maintenanceWindowStartTime?: string

	    miracastChannel?: MiracastChannel

	    miracastBlocked?: boolean

	    miracastRequirePin?: boolean

	    settingsBlockMyMeetingsAndFiles?: boolean

	    settingsBlockSessionResume?: boolean

	    settingsBlockSigninSuggestions?: boolean

	    settingsDefaultVolume?: number

	    settingsScreenTimeoutInMinutes?: number

	    settingsSessionTimeoutInMinutes?: number

	    settingsSleepTimeoutInMinutes?: number

	    welcomeScreenBlockAutomaticWakeUp?: boolean

	    welcomeScreenBackgroundImageUrl?: string

	    welcomeScreenMeetingInformation?: WelcomeScreenMeetingInformation

}

export interface EditionUpgradeConfiguration extends DeviceConfiguration {

	    licenseType?: EditionUpgradeLicenseType

	    targetEdition?: Windows10EditionType

	    license?: string

	    productKey?: string

}

export interface WindowsDefenderAdvancedThreatProtectionConfiguration extends DeviceConfiguration {

	    advancedThreatProtectionOnboardingBlob?: string

	    allowSampleSharing?: boolean

	    advancedThreatProtectionOffboardingBlob?: string

}

export interface LocalizedNotificationMessage extends Entity {

	    lastModifiedDateTime?: string

	    locale?: string

	    subject?: string

	    messageTemplate?: string

	    isDefault?: boolean

}

export interface AndroidForWorkCompliancePolicy extends DeviceCompliancePolicy {

	    passwordRequired?: boolean

	    passwordMinimumLength?: number

	    passwordRequiredType?: AndroidRequiredPasswordType

	    passwordMinutesOfInactivityBeforeLock?: number

	    passwordExpirationDays?: number

	    passwordPreviousPasswordBlockCount?: number

	    securityPreventInstallAppsFromUnknownSources?: boolean

	    securityDisableUsbDebugging?: boolean

	    deviceThreatProtectionEnabled?: boolean

	    deviceThreatProtectionRequiredSecurityLevel?: DeviceThreatProtectionLevel

	    securityBlockJailbrokenDevices?: boolean

	    osMinimumVersion?: string

	    osMaximumVersion?: string

	    minAndroidSecurityPatchLevel?: string

	    storageRequireEncryption?: boolean

}

export interface AndroidCompliancePolicy extends DeviceCompliancePolicy {

	    passwordRequired?: boolean

	    passwordMinimumLength?: number

	    passwordRequiredType?: AndroidRequiredPasswordType

	    passwordMinutesOfInactivityBeforeLock?: number

	    passwordExpirationDays?: number

	    passwordPreviousPasswordBlockCount?: number

	    securityPreventInstallAppsFromUnknownSources?: boolean

	    securityDisableUsbDebugging?: boolean

	    requireAppVerify?: boolean

	    deviceThreatProtectionEnabled?: boolean

	    deviceThreatProtectionRequiredSecurityLevel?: DeviceThreatProtectionLevel

	    securityBlockJailbrokenDevices?: boolean

	    osMinimumVersion?: string

	    osMaximumVersion?: string

	    minAndroidSecurityPatchLevel?: string

	    storageRequireEncryption?: boolean

}

export interface AndroidDeviceComplianceLocalActionBase extends Entity {

}

export interface AndroidDeviceComplianceLocalActionLockDevice extends AndroidDeviceComplianceLocalActionBase {

}

export interface IosCompliancePolicy extends DeviceCompliancePolicy {

	    passcodeBlockSimple?: boolean

	    passcodeExpirationDays?: number

	    passcodeMinimumLength?: number

	    passcodeMinutesOfInactivityBeforeLock?: number

	    passcodePreviousPasscodeBlockCount?: number

	    passcodeMinimumCharacterSetCount?: number

	    passcodeRequiredType?: RequiredPasswordType

	    passcodeRequired?: boolean

	    osMinimumVersion?: string

	    osMaximumVersion?: string

	    securityBlockJailbrokenDevices?: boolean

	    deviceThreatProtectionEnabled?: boolean

	    deviceThreatProtectionRequiredSecurityLevel?: DeviceThreatProtectionLevel

	    managedEmailProfileRequired?: boolean

}

export interface MacOSCompliancePolicy extends DeviceCompliancePolicy {

	    passwordRequired?: boolean

	    passwordBlockSimple?: boolean

	    passwordExpirationDays?: number

	    passwordMinimumLength?: number

	    passwordMinutesOfInactivityBeforeLock?: number

	    passwordPreviousPasswordBlockCount?: number

	    passwordMinimumCharacterSetCount?: number

	    passwordRequiredType?: RequiredPasswordType

	    osMinimumVersion?: string

	    osMaximumVersion?: string

	    systemIntegrityProtectionEnabled?: boolean

	    deviceThreatProtectionEnabled?: boolean

	    deviceThreatProtectionRequiredSecurityLevel?: DeviceThreatProtectionLevel

	    storageRequireEncryption?: boolean

}

export interface DefaultDeviceCompliancePolicy extends DeviceCompliancePolicy {

}

export interface Windows10CompliancePolicy extends DeviceCompliancePolicy {

	    passwordRequired?: boolean

	    passwordBlockSimple?: boolean

	    passwordRequiredToUnlockFromIdle?: boolean

	    passwordMinutesOfInactivityBeforeLock?: number

	    passwordExpirationDays?: number

	    passwordMinimumLength?: number

	    passwordMinimumCharacterSetCount?: number

	    passwordRequiredType?: RequiredPasswordType

	    passwordPreviousPasswordBlockCount?: number

	    requireHealthyDeviceReport?: boolean

	    osMinimumVersion?: string

	    osMaximumVersion?: string

	    mobileOsMinimumVersion?: string

	    mobileOsMaximumVersion?: string

	    earlyLaunchAntiMalwareDriverEnabled?: boolean

	    bitLockerEnabled?: boolean

	    secureBootEnabled?: boolean

	    codeIntegrityEnabled?: boolean

	    storageRequireEncryption?: boolean

}

export interface Windows10MobileCompliancePolicy extends DeviceCompliancePolicy {

	    passwordRequired?: boolean

	    passwordBlockSimple?: boolean

	    passwordMinimumLength?: number

	    passwordMinimumCharacterSetCount?: number

	    passwordRequiredType?: RequiredPasswordType

	    passwordPreviousPasswordBlockCount?: number

	    passwordExpirationDays?: number

	    passwordMinutesOfInactivityBeforeLock?: number

	    passwordRequireToUnlockFromIdle?: boolean

	    osMinimumVersion?: string

	    osMaximumVersion?: string

	    earlyLaunchAntiMalwareDriverEnabled?: boolean

	    bitLockerEnabled?: boolean

	    secureBootEnabled?: boolean

	    codeIntegrityEnabled?: boolean

	    storageRequireEncryption?: boolean

}

export interface Windows81CompliancePolicy extends DeviceCompliancePolicy {

	    passwordRequired?: boolean

	    passwordBlockSimple?: boolean

	    passwordExpirationDays?: number

	    passwordMinimumLength?: number

	    passwordMinutesOfInactivityBeforeLock?: number

	    passwordMinimumCharacterSetCount?: number

	    passwordRequiredType?: RequiredPasswordType

	    passwordPreviousPasswordBlockCount?: number

	    osMinimumVersion?: string

	    osMaximumVersion?: string

	    storageRequireEncryption?: boolean

}

export interface WindowsPhone81CompliancePolicy extends DeviceCompliancePolicy {

	    passwordBlockSimple?: boolean

	    passwordExpirationDays?: number

	    passwordMinimumLength?: number

	    passwordMinutesOfInactivityBeforeLock?: number

	    passwordMinimumCharacterSetCount?: number

	    passwordRequiredType?: RequiredPasswordType

	    passwordPreviousPasswordBlockCount?: number

	    passwordRequired?: boolean

	    osMinimumVersion?: string

	    osMaximumVersion?: string

	    storageRequireEncryption?: boolean

}

export interface ComplianceSettingStateSummary extends Entity {

	    setting?: string

	    settingName?: string

	    platformType?: PolicyPlatformType

	    unknownDeviceCount?: number

	    notApplicableDeviceCount?: number

	    compliantDeviceCount?: number

	    remediatedDeviceCount?: number

	    nonCompliantDeviceCount?: number

	    errorDeviceCount?: number

	    conflictDeviceCount?: number

}

export interface DeviceComplianceSettingState extends Entity {

	    platformType?: DeviceType

	    setting?: string

	    settingName?: string

	    deviceId?: string

	    deviceName?: string

	    userId?: string

	    userEmail?: string

	    userName?: string

	    userPrincipalName?: string

	    deviceModel?: string

	    state?: ComplianceStatus

	    complianceGracePeriodExpirationDateTime?: string

}

export interface OnPremisesConditionalAccessSettings extends Entity {

	    enabled?: boolean

	    includedGroups?: string[]

	    excludedGroups?: string[]

	    overrideDefaultRule?: boolean

}

export interface MobileAppIdentifierDeployment extends Entity {

	    mobileAppIdentifier?: MobileAppIdentifier

	    version?: string

}

export interface ManagedAppOperation extends Entity {

	    displayName?: string

	    lastModifiedDateTime?: string

	    state?: string

	    version?: string

}

export interface ManagedAppPolicyDeploymentSummary extends Entity {

	    displayName?: string

	    configurationDeployedUserCount?: number

	    lastRefreshTime?: string

	    configurationDeploymentSummaryPerApp?: ManagedAppPolicyDeploymentSummaryPerApp[]

	    version?: string

}

export interface WindowsInformationProtectionAppLockerFile extends Entity {

	    displayName?: string

	    fileHash?: string

	    file?: number

	    version?: string

}

export interface IosManagedAppRegistration extends ManagedAppRegistration {

}

export interface AndroidManagedAppRegistration extends ManagedAppRegistration {

}

export interface ManagedAppStatusRaw extends ManagedAppStatus {

	    content?: any

}

export interface EBookGroupAssignment extends Entity {

	    targetGroupId?: string

	    installIntent?: InstallIntent

	    eBook?: ManagedEBook

}

export interface EBookInstallSummary extends Entity {

	    installedDeviceCount?: number

	    failedDeviceCount?: number

	    notInstalledDeviceCount?: number

	    installedUserCount?: number

	    failedUserCount?: number

	    notInstalledUserCount?: number

}

export interface DeviceInstallState extends Entity {

	    deviceName?: string

	    deviceId?: string

	    lastSyncDateTime?: string

	    installState?: InstallState

	    errorCode?: string

	    osVersion?: string

	    osDescription?: string

	    userName?: string

}

export interface UserInstallStateSummary extends Entity {

	    userName?: string

	    installedDeviceCount?: number

	    failedDeviceCount?: number

	    notInstalledDeviceCount?: number

	    deviceStates?: DeviceInstallState[]

}

export interface EBookVppGroupAssignment extends EBookGroupAssignment {

	    useDeviceLicensing?: boolean

}

export interface IosVppEBook extends ManagedEBook {

	    vppTokenId?: string

	    appleId?: string

	    vppOrganizationName?: string

	    genres?: string[]

	    language?: string

	    seller?: string

	    totalLicenseCount?: number

	    usedLicenseCount?: number

}

export interface PayloadResponse extends Entity {

}

export interface ChatThread extends Entity {

	    chatMessages?: ChatMessage[]

	    rootMessage?: ChatMessage

}

export interface ChatMessage extends Entity {

	    body?: ItemBody

	    inReplyTo?: ChatMessage

	    replies?: ChatMessage[]

	    from?: User

}
export interface AssignedLicense {

	    disabledPlans?: string[]

	    skuId?: string

}
export interface AssignedPlan {

	    assignedDateTime?: string

	    capabilityStatus?: string

	    service?: string

	    servicePlanId?: string

}
export interface DeviceKey {

	    keyType?: string

	    keyMaterial?: number

	    deviceId?: string

}
export interface OnPremisesProvisioningError {

	    value?: string

	    category?: string

	    propertyCausingError?: string

	    occurredDateTime?: string

}
export interface PasswordProfile {

	    password?: string

	    forceChangePasswordNextSignIn?: boolean

}
export interface ProvisionedPlan {

	    capabilityStatus?: string

	    provisioningStatus?: string

	    service?: string

}
export interface MailboxSettings {

	    automaticRepliesSetting?: AutomaticRepliesSetting

	    archiveFolder?: string

	    timeZone?: string

	    language?: LocaleInfo

}
export interface AutomaticRepliesSetting {

	    status?: AutomaticRepliesStatus

	    externalAudience?: ExternalAudienceScope

	    scheduledStartDateTime?: DateTimeTimeZone

	    scheduledEndDateTime?: DateTimeTimeZone

	    internalReplyMessage?: string

	    externalReplyMessage?: string

}
export interface DateTimeTimeZone {

	    dateTime?: string

	    timeZone?: string

}
export interface LocaleInfo {

	    locale?: string

	    displayName?: string

}
export interface IdentityUserRisk {

	    level?: UserRiskLevel

	    lastChangedDateTime?: string

}
export interface AlternativeSecurityId {

	    type?: number

	    identityProvider?: string

	    key?: number

}
export interface VerifiedDomain {

	    capabilities?: string

	    isDefault?: boolean

	    isInitial?: boolean

	    name?: string

	    type?: string

}
export interface DefaultDeviceEnrollmentRestrictions {

	    iosRestrictions?: DeviceEnrollmentPlatformRestrictions

	    windowsRestrictions?: DeviceEnrollmentPlatformRestrictions

	    windowsMobileRestrictions?: DeviceEnrollmentPlatformRestrictions

	    androidRestrictions?: DeviceEnrollmentPlatformRestrictions

	    androidForWorkRestrictions?: DeviceEnrollmentPlatformRestrictions

	    macRestrictions?: DeviceEnrollmentPlatformRestrictions

}
export interface DeviceEnrollmentPlatformRestrictions {

	    platformBlocked?: boolean

	    personalDeviceEnrollmentBlocked?: boolean

	    osMinimumVersion?: string

	    osMaximumVersion?: string

}
export interface DefaultDeviceEnrollmentWindowsHelloForBusinessSettings {

	    pinMinimumLength?: number

	    pinMaximumLength?: number

	    pinUppercaseLettersUsage?: WindowsHelloForBusinessPinUsage

	    pinLowercaseLettersUsage?: WindowsHelloForBusinessPinUsage

	    pinSpecialCharactersUsage?: WindowsHelloForBusinessPinUsage

	    windowsHelloForBusiness?: WindowsHelloForBusinessConfiguration

	    securityDeviceRequired?: boolean

	    unlockWithBiometricsEnabled?: boolean

	    mobilePinSignInEnabled?: boolean

	    pinPreviousBlockCount?: number

	    pinExpirationInDays?: number

	    enhancedBiometrics?: WindowsHelloForBusinessConfiguration

}
export interface CertificateConnectorSetting {

	    status?: number

	    certExpiryTime?: string

	    enrollmentError?: string

	    lastConnectorConnectionTime?: string

	    connectorVersion?: string

	    lastUploadVersion?: number

}
export interface ExtensionSchemaProperty {

	    name?: string

	    type?: string

}
export interface Api {

	    acceptedAccessTokenVersion?: number

	    publishedPermissionScopes?: PermissionScope[]

}
export interface PermissionScope {

	    adminConsentDescription?: string

	    adminConsentDisplayName?: string

	    id?: string

	    isEnabled?: boolean

	    origin?: string

	    type?: string

	    userConsentDescription?: string

	    userConsentDisplayName?: string

	    value?: string

}
export interface AppRole {

	    allowedMemberTypes?: string[]

	    description?: string

	    displayName?: string

	    id?: string

	    isEnabled?: boolean

	    origin?: string

	    value?: string

}
export interface InstalledClient {

	    redirectUrls?: string[]

}
export interface InformationalUrl {

	    logoUrl?: string

	    marketingUrl?: string

	    privacyStatementUrl?: string

	    supportUrl?: string

	    termsOfServiceUrl?: string

}
export interface KeyCredential {

	    customKeyIdentifier?: number

	    endDateTime?: string

	    keyId?: string

	    startDateTime?: string

	    type?: string

	    usage?: string

	    key?: number

}
export interface PasswordCredential {

	    customKeyIdentifier?: number

	    endDateTime?: string

	    keyId?: string

	    startDateTime?: string

	    secretText?: string

	    hint?: string

}
export interface PreAuthorizedApplication {

	    appId?: string

	    permissionIds?: string[]

}
export interface RequiredResourceAccess {

	    resourceAppId?: string

	    resourceAccess?: ResourceAccess[]

}
export interface ResourceAccess {

	    id?: string

	    type?: string

}
export interface Web {

	    redirectUrls?: string[]

	    logoutUrl?: string

	    oauth2AllowImplicitFlow?: boolean

}
export interface SettingValue {

	    name?: string

	    value?: string

}
export interface SettingTemplateValue {

	    name?: string

	    type?: string

	    defaultValue?: string

	    description?: string

}
export interface DomainState {

	    status?: string

	    operation?: string

	    lastActionDateTime?: string

}
export interface ServicePlanInfo {

	    servicePlanId?: string

	    servicePlanName?: string

	    provisioningStatus?: string

	    appliesTo?: string

}
export interface AddIn {

	    id?: string

	    type?: string

	    properties?: KeyValue[]

}
export interface KeyValue {

	    key?: string

	    value?: string

}
export interface OAuth2Permission {

	    adminConsentDescription?: string

	    adminConsentDisplayName?: string

	    id?: string

	    isEnabled?: boolean

	    origin?: string

	    type?: string

	    userConsentDescription?: string

	    userConsentDisplayName?: string

	    value?: string

}
export interface LicenseUnitsDetail {

	    enabled?: number

	    suspended?: number

	    warning?: number

}
export interface Identity {

	    id?: string

	    displayName?: string

}
export interface ComplexExtensionValue {

}
export interface AllowedDataLocationInfo {

}
export interface ImageInfo {

	    iconUrl?: string

	    alternativeText?: string

	    addImageQuery?: boolean

}
export interface VisualInfo {

	    attribution?: ImageInfo

	    backgroundColor?: string

	    description?: string

	    displayText?: string

	    content?: any

}
export interface Root {

}
export interface SharepointIds {

	    listId?: string

	    listItemId?: string

	    listItemUniqueId?: string

	    siteId?: string

	    siteUrl?: string

	    webId?: string

}
export interface SiteCollection {

	    dataLocationCode?: string

	    hostname?: string

	    root?: Root

}
export interface ListInfo {

	    contentTypesEnabled?: boolean

	    hidden?: boolean

	    template?: string

}
export interface SystemFacet {

}
export interface IdentitySet {

	    application?: Identity

	    device?: Identity

	    user?: Identity

}
export interface Quota {

	    deleted?: number

	    remaining?: number

	    state?: string

	    total?: number

	    used?: number

}
export interface Audio {

	    album?: string

	    albumArtist?: string

	    artist?: string

	    bitrate?: number

	    composers?: string

	    copyright?: string

	    disc?: number

	    discCount?: number

	    duration?: number

	    genre?: string

	    hasDrm?: boolean

	    isVariableBitrate?: boolean

	    title?: string

	    track?: number

	    trackCount?: number

	    year?: number

}
export interface Deleted {

	    state?: string

}
export interface File {

	    hashes?: Hashes

	    mimeType?: string

	    processingMetadata?: boolean

}
export interface Hashes {

	    crc32Hash?: string

	    quickXorHash?: string

	    sha1Hash?: string

}
export interface FileSystemInfo {

	    createdDateTime?: string

	    lastAccessedDateTime?: string

	    lastModifiedDateTime?: string

}
export interface Folder {

	    childCount?: number

	    view?: FolderView

}
export interface FolderView {

	    sortBy?: string

	    sortOrder?: string

	    viewType?: string

}
export interface Image {

	    height?: number

	    width?: number

}
export interface GeoCoordinates {

	    altitude?: number

	    latitude?: number

	    longitude?: number

}
export interface Package {

	    type?: string

}
export interface Photo {

	    cameraMake?: string

	    cameraModel?: string

	    exposureDenominator?: number

	    exposureNumerator?: number

	    fNumber?: number

	    focalLength?: number

	    iso?: number

	    takenDateTime?: string

}
export interface PublicationFacet {

	    level?: string

	    versionId?: string

}
export interface RemoteItem {

	    createdBy?: IdentitySet

	    createdDateTime?: string

	    file?: File

	    fileSystemInfo?: FileSystemInfo

	    folder?: Folder

	    id?: string

	    lastModifiedBy?: IdentitySet

	    lastModifiedDateTime?: string

	    name?: string

	    package?: Package

	    parentReference?: ItemReference

	    shared?: Shared

	    sharepointIds?: SharepointIds

	    size?: number

	    specialFolder?: SpecialFolder

	    webDavUrl?: string

	    webUrl?: string

}
export interface ItemReference {

	    driveId?: string

	    driveType?: string

	    id?: string

	    name?: string

	    path?: string

	    shareId?: string

	    sharepointIds?: SharepointIds

}
export interface Shared {

	    owner?: IdentitySet

	    scope?: string

	    sharedBy?: IdentitySet

	    sharedDateTime?: string

}
export interface SpecialFolder {

	    name?: string

}
export interface SearchResult {

	    onClickTelemetryUrl?: string

}
export interface Video {

	    audioBitsPerSample?: number

	    audioChannels?: number

	    audioFormat?: string

	    audioSamplesPerSecond?: number

	    bitrate?: number

	    duration?: number

	    fourCC?: string

	    frameRate?: number

	    height?: number

	    width?: number

}
export interface WorkbookSessionInfo {

	    id?: string

	    persistChanges?: boolean

}
export interface WorkbookFilterCriteria {

	    color?: string

	    criterion1?: string

	    criterion2?: string

	    dynamicCriteria?: string

	    filterOn?: string

	    icon?: WorkbookIcon

	    operator?: string

	    values?: any

}
export interface WorkbookIcon {

	    index?: number

	    set?: string

}
export interface WorkbookSortField {

	    ascending?: boolean

	    color?: string

	    dataOption?: string

	    icon?: WorkbookIcon

	    key?: number

	    sortOn?: string

}
export interface WorkbookWorksheetProtectionOptions {

	    allowAutoFilter?: boolean

	    allowDeleteColumns?: boolean

	    allowDeleteRows?: boolean

	    allowFormatCells?: boolean

	    allowFormatColumns?: boolean

	    allowFormatRows?: boolean

	    allowInsertColumns?: boolean

	    allowInsertHyperlinks?: boolean

	    allowInsertRows?: boolean

	    allowPivotTables?: boolean

	    allowSort?: boolean

}
export interface WorkbookFilterDatetime {

	    date?: string

	    specificity?: string

}
export interface WorkbookRangeReference {

	    address?: string

}
export interface Recipient {

	    emailAddress?: EmailAddress

}
export interface EmailAddress {

	    name?: string

	    address?: string

}
export interface AttendeeBase extends Recipient {

	    type?: AttendeeType

}
export interface MeetingTimeSuggestionsResult {

	    meetingTimeSuggestions?: MeetingTimeSuggestion[]

	    emptySuggestionsReason?: string

}
export interface MeetingTimeSuggestion {

	    meetingTimeSlot?: TimeSlot

	    confidence?: number

	    organizerAvailability?: FreeBusyStatus

	    attendeeAvailability?: AttendeeAvailability[]

	    locations?: Location[]

	    suggestionReason?: string

}
export interface TimeSlot {

	    start?: DateTimeTimeZone

	    end?: DateTimeTimeZone

}
export interface AttendeeAvailability {

	    attendee?: AttendeeBase

	    availability?: FreeBusyStatus

}
export interface Location {

	    displayName?: string

	    locationEmailAddress?: string

	    address?: PhysicalAddress

	    coordinates?: OutlookGeoCoordinates

	    locationUri?: string

	    locationType?: LocationType

	    uniqueId?: string

	    uniqueIdType?: LocationUniqueIdType

}
export interface PhysicalAddress {

	    type?: PhysicalAddressType

	    postOfficeBox?: string

	    street?: string

	    city?: string

	    state?: string

	    countryOrRegion?: string

	    postalCode?: string

}
export interface OutlookGeoCoordinates {

	    altitude?: number

	    latitude?: number

	    longitude?: number

	    accuracy?: number

	    altitudeAccuracy?: number

}
export interface LocationConstraint {

	    isRequired?: boolean

	    suggestLocation?: boolean

	    locations?: LocationConstraintItem[]

}
export interface LocationConstraintItem extends Location {

	    resolveAvailability?: boolean

}
export interface TimeConstraint {

	    activityDomain?: ActivityDomain

	    timeslots?: TimeSlot[]

}
export interface MeetingTimeCandidatesResult {

	    meetingTimeSlots?: MeetingTimeCandidate[]

	    emptySuggestionsHint?: string

}
export interface MeetingTimeCandidate {

	    meetingTimeSlot?: TimeSlotOLD

	    confidence?: number

	    organizerAvailability?: FreeBusyStatus

	    attendeeAvailability?: AttendeeAvailability[]

	    locations?: Location[]

	    suggestionHint?: string

}
export interface TimeSlotOLD {

	    start?: TimeStamp

	    end?: TimeStamp

}
export interface TimeStamp {

	    date?: string

	    time?: string

	    timeZone?: string

}
export interface MailTips {

	    emailAddress?: EmailAddress

	    automaticReplies?: AutomaticRepliesMailTips

	    mailboxFull?: boolean

	    customMailTip?: string

	    externalMemberCount?: number

	    totalMemberCount?: number

	    deliveryRestricted?: boolean

	    isModerated?: boolean

	    recipientScope?: RecipientScopeType

	    recipientSuggestions?: Recipient[]

	    maxMessageSize?: number

	    error?: MailTipsError

}
export interface AutomaticRepliesMailTips {

	    message?: string

	    messageLanguage?: LocaleInfo

	    scheduledStartTime?: DateTimeTimeZone

	    scheduledEndTime?: DateTimeTimeZone

}
export interface MailTipsError {

	    message?: string

	    code?: string

}
export interface Reminder {

	    eventId?: string

	    eventStartTime?: DateTimeTimeZone

	    eventEndTime?: DateTimeTimeZone

	    changeKey?: string

	    eventSubject?: string

	    eventLocation?: Location

	    eventWebLink?: string

	    reminderFireTime?: DateTimeTimeZone

}
export interface TimeZoneInformation {

	    alias?: string

	    displayName?: string

}
export interface InternetMessageHeader {

	    name?: string

	    value?: string

}
export interface ItemBody {

	    contentType?: BodyType

	    content?: string

}
export interface MentionsPreview {

	    isMentioned?: boolean

}
export interface FollowupFlag {

	    completedDateTime?: DateTimeTimeZone

	    dueDateTime?: DateTimeTimeZone

	    startDateTime?: DateTimeTimeZone

	    flagStatus?: FollowupFlagStatus

}
export interface ResponseStatus {

	    response?: ResponseType

	    time?: string

}
export interface PatternedRecurrence {

	    pattern?: RecurrencePattern

	    range?: RecurrenceRange

}
export interface RecurrencePattern {

	    type?: RecurrencePatternType

	    interval?: number

	    month?: number

	    dayOfMonth?: number

	    daysOfWeek?: DayOfWeek[]

	    firstDayOfWeek?: DayOfWeek

	    index?: WeekIndex

}
export interface RecurrenceRange {

	    type?: RecurrenceRangeType

	    startDate?: string

	    endDate?: string

	    recurrenceTimeZone?: string

	    numberOfOccurrences?: number

}
export interface Attendee extends AttendeeBase {

	    status?: ResponseStatus

}
export interface EventCreationOptions {

	    saveToGroupCalendarOnly?: boolean

}
export interface Phone {

	    type?: PhoneType

	    number?: string

}
export interface Website {

	    type?: WebsiteType

	    address?: string

	    displayName?: string

}
export interface MessageRulePredicates {

	    categories?: string[]

	    subjectContains?: string[]

	    bodyContains?: string[]

	    bodyOrSubjectContains?: string[]

	    senderContains?: string[]

	    recipientContains?: string[]

	    headerContains?: string[]

	    messageActionFlag?: MessageActionFlag

	    importance?: Importance

	    sensitivity?: Sensitivity

	    fromAddresses?: Recipient[]

	    sentToAddresses?: Recipient[]

	    sentToMe?: boolean

	    sentOnlyToMe?: boolean

	    sentCcMe?: boolean

	    sentToOrCcMe?: boolean

	    notSentToMe?: boolean

	    hasAttachments?: boolean

	    isApprovalRequest?: boolean

	    isAutomaticForward?: boolean

	    isAutomaticReply?: boolean

	    isEncrypted?: boolean

	    isMeetingRequest?: boolean

	    isMeetingResponse?: boolean

	    isNonDeliveryReport?: boolean

	    isPermissionControlled?: boolean

	    isReadReceipt?: boolean

	    isSigned?: boolean

	    isVoicemail?: boolean

	    withinSizeRange?: SizeRange

}
export interface SizeRange {

	    minimumSize?: number

	    maximumSize?: number

}
export interface MessageRuleActions {

	    moveToFolder?: string

	    copyToFolder?: string

	    delete?: boolean

	    permanentDelete?: boolean

	    markAsRead?: boolean

	    markImportance?: Importance

	    forwardTo?: Recipient[]

	    forwardAsAttachmentTo?: Recipient[]

	    redirectTo?: Recipient[]

	    assignCategories?: string[]

	    stopProcessingRules?: boolean

}
export interface RankedEmailAddress {

	    address?: string

	    rank?: number

}
export interface PersonDataSource {

	    type?: string

}
export interface BooleanColumn {

}
export interface CalculatedColumn {

	    format?: string

	    formula?: string

	    outputType?: string

}
export interface ChoiceColumn {

	    allowTextEntry?: boolean

	    choices?: string[]

	    displayAs?: string

}
export interface CurrencyColumn {

	    locale?: string

}
export interface DateTimeColumn {

	    displayAs?: string

	    format?: string

}
export interface DefaultColumnValue {

	    formula?: string

	    value?: string

}
export interface LookupColumn {

	    allowMultipleValues?: boolean

	    allowUnlimitedLength?: boolean

	    columnName?: string

	    listId?: string

	    primaryLookupColumnId?: string

}
export interface NumberColumn {

	    decimalPlaces?: string

	    displayAs?: string

	    maximum?: number

	    minimum?: number

}
export interface PersonOrGroupColumn {

	    allowMultipleSelection?: boolean

	    displayAs?: string

	    chooseFromType?: string

}
export interface TextColumn {

	    allowMultipleLines?: boolean

	    appendChangesToExistingText?: boolean

	    linesForEditing?: number

	    maxLength?: number

	    textType?: string

}
export interface ContentTypeOrder {

	    default?: boolean

	    position?: number

}
export interface ItemActionSet {

	    comment?: CommentAction

	    create?: CreateAction

	    delete?: DeleteAction

	    edit?: EditAction

	    mention?: MentionAction

	    move?: MoveAction

	    rename?: RenameAction

	    restore?: RestoreAction

	    share?: ShareAction

	    version?: VersionAction

}
export interface CommentAction {

	    isReply?: boolean

	    parentAuthor?: IdentitySet

	    participants?: IdentitySet[]

}
export interface CreateAction {

}
export interface DeleteAction {

	    name?: string

}
export interface EditAction {

}
export interface MentionAction {

	    mentionees?: IdentitySet[]

}
export interface MoveAction {

	    From?: string

	    to?: string

}
export interface RenameAction {

	    oldName?: string

}
export interface RestoreAction {

}
export interface ShareAction {

	    recipients?: IdentitySet[]

}
export interface VersionAction {

	    newVersion?: string

}
export interface ItemActivityTimeSet {

	    observedDateTime?: string

	    recordedDateTime?: string

}
export interface ContentTypeInfo {

	    id?: string

}
export interface SharingInvitation {

	    email?: string

	    invitedBy?: IdentitySet

	    redeemedBy?: string

	    signInRequired?: boolean

}
export interface SharingLink {

	    application?: Identity

	    configuratorUrl?: string

	    scope?: string

	    type?: string

	    webHtml?: string

	    webUrl?: string

}
export interface Thumbnail {

	    content?: any

	    height?: number

	    sourceItemId?: string

	    url?: string

	    width?: number

}
export interface DriveItemUploadableProperties {

	    description?: string

	    fileSystemInfo?: FileSystemInfo

	    name?: string

}
export interface DriveRecipient {

	    alias?: string

	    email?: string

	    objectId?: string

}
export interface FlexSchemaContainer {

}
export interface UploadSession {

	    expirationDateTime?: string

	    nextExpectedRanges?: string[]

	    uploadUrl?: string

}
export interface ResourceVisualization {

	    title?: string

	    type?: string

	    mediaType?: string

	    previewImageUrl?: string

	    previewText?: string

	    containerWebUrl?: string

	    containerDisplayName?: string

	    containerType?: string

}
export interface ResourceReference {

	    webUrl?: string

	    id?: string

	    type?: string

}
export interface SharingDetail {

	    sharedBy?: InsightIdentity

	    sharedDateTime?: string

	    sharingSubject?: string

	    sharingType?: string

	    sharingReference?: ResourceReference

}
export interface InsightIdentity {

	    displayName?: string

	    id?: string

	    address?: string

}
export interface UsageDetails {

	    lastAccessedDateTime?: string

	    lastModifiedDateTime?: string

}
export interface PlannerAppliedCategories {

}
export interface PlannerAssignments {

}
export interface PlannerExternalReference {

	    alias?: string

	    type?: string

	    previewPriority?: string

	    lastModifiedBy?: IdentitySet

	    lastModifiedDateTime?: string

}
export interface PlannerChecklistItem {

	    isChecked?: boolean

	    title?: string

	    orderHint?: string

	    lastModifiedBy?: IdentitySet

	    lastModifiedDateTime?: string

}
export interface PlannerAssignment {

	    assignedBy?: IdentitySet

	    assignedDateTime?: string

	    orderHint?: string

}
export interface PlannerExternalReferences {

}
export interface PlannerChecklistItems {

}
export interface PlannerOrderHintsByAssignee {

}
export interface PlannerUserIds {

}
export interface PlannerCategoryDescriptions {

	    category1?: string

	    category2?: string

	    category3?: string

	    category4?: string

	    category5?: string

	    category6?: string

}
export interface AppliedCategoriesCollection {

}
export interface ExternalReferenceCollection {

}
export interface ChecklistItemCollection {

}
export interface UserIdCollection {

}
export interface ExternalReference {

	    alias?: string

	    type?: string

	    previewPriority?: string

	    lastModifiedBy?: string

	    lastModifiedDateTime?: string

}
export interface ChecklistItem {

	    isChecked?: boolean

	    title?: string

	    orderHint?: string

	    lastModifiedBy?: string

	    lastModifiedDateTime?: string

}
export interface NotebookLinks {

	    oneNoteClientUrl?: ExternalLink

	    oneNoteWebUrl?: ExternalLink

}
export interface ExternalLink {

	    href?: string

}
export interface SectionLinks {

	    oneNoteClientUrl?: ExternalLink

	    oneNoteWebUrl?: ExternalLink

}
export interface PageLinks {

	    oneNoteClientUrl?: ExternalLink

	    oneNoteWebUrl?: ExternalLink

}
export interface OnenoteOperationError {

	    code?: string

	    message?: string

}
export interface Diagnostic {

	    message?: string

	    url?: string

}
export interface OnenotePatchContentCommand {

	    action?: OnenotePatchActionType

	    target?: string

	    content?: string

	    position?: OnenotePatchInsertPosition

}
export interface OnenotePagePreview {

	    previewText?: string

	    links?: OnenotePagePreviewLinks

}
export interface OnenotePagePreviewLinks {

	    previewImageUrl?: ExternalLink

}
export interface RecentNotebook {

	    displayName?: string

	    lastAccessedTime?: string

	    links?: RecentNotebookLinks

	    sourceService?: OnenoteSourceService

}
export interface RecentNotebookLinks {

	    oneNoteClientUrl?: ExternalLink

	    oneNoteWebUrl?: ExternalLink

}
export interface SignInLocation {

	    city?: string

	    state?: string

	    countryOrRegion?: string

	    geoCoordinates?: GeoCoordinates

}
export interface RoleSuccessStatistics {

	    roleId?: string

	    roleName?: string

	    temporarySuccess?: number

	    temporaryFail?: number

	    permanentSuccess?: number

	    permanentFail?: number

	    removeSuccess?: number

	    removeFail?: number

	    unknownFail?: number

}
export interface InvitedUserMessageInfo {

	    ccRecipients?: Recipient[]

	    messageLanguage?: string

	    customizedMessageBody?: string

}
export interface DeviceManagementSettings {

	    windowsCommercialId?: string

	    windowsCommercialIdLastModifiedTime?: string

	    deviceComplianceCheckinThresholdDays?: number

	    isScheduledActionEnabled?: boolean

	    secureByDefault?: boolean

}
export interface IntuneBrand {

	    displayName?: string

	    contactITName?: string

	    contactITPhoneNumber?: string

	    contactITEmailAddress?: string

	    contactITNotes?: string

	    privacyUrl?: string

	    onlineSupportSiteUrl?: string

	    onlineSupportSiteName?: string

	    themeColor?: RgbColor

	    showLogo?: boolean

	    lightBackgroundLogo?: MimeContent

	    darkBackgroundLogo?: MimeContent

	    showNameNextToLogo?: boolean

}
export interface RgbColor {

	    r?: number

	    g?: number

	    b?: number

}
export interface MimeContent {

	    type?: string

	    value?: number

}
export interface AndroidForWorkAppConfigurationExample {

}
export interface AndroidForWorkAppConfigurationSchemaItem {

	    schemaItemKey?: string

	    displayName?: string

	    description?: string

	    defaultBoolValue?: boolean

	    defaultIntValue?: number

	    defaultStringValue?: string

	    defaultStringArrayValue?: string[]

	    dataType?: AndroidForWorkAppConfigurationSchemaItemDataType

	    selections?: KeyValuePair[]

}
export interface KeyValuePair {

	    name?: string

	    value?: string

}
export interface AndroidForWorkAppConfigurationExampleJson extends AndroidForWorkAppConfigurationExample {

	    example?: string

}
export interface AppConfigurationSettingItem {

	    appConfigKey?: string

	    appConfigKeyType?: MdmAppConfigKeyType

	    appConfigKeyValue?: string

}
export interface FileEncryptionInfo {

	    encryptionKey?: number

	    initializationVector?: number

	    mac?: number

	    macKey?: number

	    profileIdentifier?: string

	    fileDigest?: number

	    fileDigestAlgorithm?: string

}
export interface ExcludedApps {

	    access?: boolean

	    excel?: boolean

	    groove?: boolean

	    infoPath?: boolean

	    lync?: boolean

	    oneDrive?: boolean

	    oneNote?: boolean

	    outlook?: boolean

	    powerPoint?: boolean

	    publisher?: boolean

	    sharePointDesigner?: boolean

	    visio?: boolean

	    word?: boolean

}
export interface AndroidMinimumOperatingSystem {

	    v4_0?: boolean

	    v4_0_3?: boolean

	    v4_1?: boolean

	    v4_2?: boolean

	    v4_3?: boolean

	    v4_4?: boolean

	    v5_0?: boolean

	    v5_1?: boolean

}
export interface IosDeviceType {

	    iPad?: boolean

	    iPhoneAndIPod?: boolean

}
export interface IosMinimumOperatingSystem {

	    v8_0?: boolean

	    v9_0?: boolean

	    v10_0?: boolean

}
export interface WindowsMinimumOperatingSystem {

	    v8_0?: boolean

	    v8_1?: boolean

	    v10_0?: boolean

}
export interface WindowsPackageInformation {

	    applicableArchitecture?: WindowsArchitecture

	    displayName?: string

	    identityName?: string

	    identityPublisher?: string

	    identityResourceIdentifier?: string

	    identityVersion?: string

	    minimumSupportedOperatingSystem?: WindowsMinimumOperatingSystem

}
export interface VppLicensingType {

	    supportUserLicensing?: boolean

	    supportDeviceLicensing?: boolean

}
export interface AppInstallationFailure {

	    applicationId?: string

	    appName?: string

	    platformId?: number

	    userFailures?: number

	    deviceFailures?: number

}
export interface AndroidPermissionAction {

	    permission?: string

	    action?: AndroidPermissionActionType

}
export interface ManagementCertificateWithThumbprint {

	    thumbprint?: string

	    certificate?: string

}
export interface HardwareInformation {

	    serialNumber?: string

	    totalStorageSpace?: number

	    freeStorageSpace?: number

	    imei?: string

	    meid?: string

	    manufacturer?: string

	    model?: string

	    phoneNumber?: string

	    subscriberCarrier?: string

	    cellularTechnology?: string

	    wifiMac?: string

	    operatingSystemLanguage?: string

	    isSupervised?: boolean

	    isEncrypted?: boolean

	    isSharedDevice?: boolean

	    sharedDeviceCachedUsers?: SharedAppleDeviceUser[]

}
export interface SharedAppleDeviceUser {

	    userPrincipalName?: string

	    dataToSync?: boolean

	    dataQuota?: number

	    dataUsed?: number

}
export interface DeviceActionResult {

	    actionName?: string

	    actionState?: DeviceActionState

	    startDateTime?: string

	    lastUpdatedDateTime?: string

}
export interface ConfigurationManagerClientEnabledFeatures {

	    inventory?: boolean

	    modernApps?: boolean

	    resourceAccess?: boolean

	    deviceConfiguration?: boolean

	    compliancePolicy?: boolean

	    windowsUpdateForBusiness?: boolean

}
export interface DeviceOperatingSystemSummary {

	    androidCount?: number

	    iosCount?: number

	    macOSCount?: number

	    windowsMobileCount?: number

	    windowsCount?: number

	    unknownCount?: number

}
export interface DeviceExchangeAccessStateSummary {

	    allowedDeviceCount?: number

	    blockedDeviceCount?: number

	    quarantinedDeviceCount?: number

	    unknownDeviceCount?: number

	    unavailableDeviceCount?: number

}
export interface RunSchedule {

}
export interface WindowsDefenderScanActionResult extends DeviceActionResult {

	    scanType?: string

}
export interface DeleteUserFromSharedAppleDeviceActionResult extends DeviceActionResult {

	    userPrincipalName?: string

}
export interface DeviceGeoLocation {

	    lastCollectedDateTimeUtc?: string

	    longitude?: number

	    latitude?: number

	    altitude?: number

	    horizontalAccuracy?: number

	    verticalAccuracy?: number

	    heading?: number

	    speed?: number

}
export interface LocateDeviceActionResult extends DeviceActionResult {

	    deviceLocation?: DeviceGeoLocation

}
export interface ResetPasscodeActionResult extends DeviceActionResult {

	    passcode?: string

}
export interface DailySchedule extends RunSchedule {

	    interval?: number

}
export interface HourlySchedule extends RunSchedule {

	    interval?: number

}
export interface IpRange {

}
export interface IPv6Range extends IpRange {

	    lowerAddress?: string

	    upperAddress?: string

}
export interface IPv4Range extends IpRange {

	    lowerAddress?: string

	    upperAddress?: string

}
export interface Report {

	    content?: any

}
export interface ExtendedKeyUsage {

	    name?: string

	    objectIdentifier?: string

}
export interface OmaSetting {

	    displayName?: string

	    description?: string

	    omaUri?: string

}
export interface OmaSettingInteger extends OmaSetting {

	    value?: number

}
export interface OmaSettingFloatingPoint extends OmaSetting {

	    value?: number

}
export interface OmaSettingString extends OmaSetting {

	    value?: string

}
export interface OmaSettingDateTime extends OmaSetting {

	    value?: string

}
export interface OmaSettingStringXml extends OmaSetting {

	    fileName?: string

	    value?: number

}
export interface OmaSettingBoolean extends OmaSetting {

	    value?: boolean

}
export interface OmaSettingBase64 extends OmaSetting {

	    fileName?: string

	    value?: string

}
export interface VpnServer {

	    description?: string

	    ipAddressOrFqdn?: string

	    address?: string

	    isDefaultServer?: boolean

}
export interface AppListItem {

	    name?: string

	    publisher?: string

	    appStoreUrl?: string

	    appId?: string

}
export interface AppsComplianceListItem {

	    name?: string

	    publisher?: string

	    appStoreUrl?: string

	    appId?: string

}
export interface IosEduCertificateSettings {

	    trustedRootCertificate?: number

	    certFileName?: string

	    certificationAuthority?: string

	    certificationAuthorityName?: string

	    certificateTemplateName?: string

	    renewalThresholdPercentage?: number

	    certificateValidityPeriodValue?: number

	    certificateValidityPeriodScale?: CertificateValidityPeriodScale

}
export interface MediaContentRatingAustralia {

	    movieRating?: RatingAustraliaMoviesType

	    tvRating?: RatingAustraliaTelevisionType

}
export interface MediaContentRatingCanada {

	    movieRating?: RatingCanadaMoviesType

	    tvRating?: RatingCanadaTelevisionType

}
export interface MediaContentRatingFrance {

	    movieRating?: RatingFranceMoviesType

	    tvRating?: RatingFranceTelevisionType

}
export interface MediaContentRatingGermany {

	    movieRating?: RatingGermanyMoviesType

	    tvRating?: RatingGermanyTelevisionType

}
export interface MediaContentRatingIreland {

	    movieRating?: RatingIrelandMoviesType

	    tvRating?: RatingIrelandTelevisionType

}
export interface MediaContentRatingJapan {

	    movieRating?: RatingJapanMoviesType

	    tvRating?: RatingJapanTelevisionType

}
export interface MediaContentRatingNewZealand {

	    movieRating?: RatingNewZealandMoviesType

	    tvRating?: RatingNewZealandTelevisionType

}
export interface MediaContentRatingUnitedKingdom {

	    movieRating?: RatingUnitedKingdomMoviesType

	    tvRating?: RatingUnitedKingdomTelevisionType

}
export interface MediaContentRatingUnitedStates {

	    movieRating?: RatingUnitedStatesMoviesType

	    tvRating?: RatingUnitedStatesTelevisionType

}
export interface IosNetworkUsageRule {

	    managedApps?: AppListItem[]

	    cellularDataBlockWhenRoaming?: boolean

	    cellularDataBlocked?: boolean

}
export interface AirPrintDestination {

	    ipAddress?: string

	    resourcePath?: string

}
export interface IosWebContentFilterBase {

}
export interface IosHomeScreenItem {

	    displayName?: string

}
export interface IosHomeScreenPage {

	    displayName?: string

	    icons?: IosHomeScreenItem[]

}
export interface IosNotificationSettings {

	    bundleID?: string

	    appName?: string

	    publisher?: string

	    enabled?: boolean

	    showInNotificationCenter?: boolean

	    showOnLockScreen?: boolean

	    alertType?: IosNotificationAlertType

	    badgesEnabled?: boolean

	    soundsEnabled?: boolean

}
export interface IosWebContentFilterSpecificWebsitesAccess extends IosWebContentFilterBase {

	    specificWebsitesOnly?: IosBookmark[]

}
export interface IosBookmark {

	    url?: string

	    bookmarkFolder?: string

	    displayName?: string

}
export interface IosWebContentFilterAutoFilter extends IosWebContentFilterBase {

	    allowedUrls?: string[]

	    blockedUrls?: string[]

}
export interface IosHomeScreenFolder extends IosHomeScreenItem {

	    pages?: IosHomeScreenFolderPage[]

}
export interface IosHomeScreenFolderPage {

	    displayName?: string

	    apps?: IosHomeScreenApp[]

}
export interface IosHomeScreenApp extends IosHomeScreenItem {

	    bundleID?: string

}
export interface VpnOnDemandRule {

	    ssids?: string[]

	    dnsSearchDomains?: string[]

	    probeUrl?: string

	    action?: VpnOnDemandRuleConnectionAction

	    domainAction?: VpnOnDemandRuleConnectionDomainAction

	    domains?: string[]

	    probeRequiredUrl?: string

}
export interface VpnProxyServer {

	    automaticConfigurationScriptUrl?: string

	    address?: string

	    port?: number

}
export interface Windows81VpnProxyServer extends VpnProxyServer {

	    automaticallyDetectProxySettings?: boolean

	    bypassProxyServerForLocalAddress?: boolean

}
export interface Windows10VpnProxyServer extends VpnProxyServer {

	    bypassProxyServerForLocalAddress?: boolean

}
export interface BitLockerSystemDrivePolicy {

	    encryptionMethod?: BitLockerEncryptionMethod

	    startupAuthenticationRequired?: boolean

	    startupAuthenticationBlockWithoutTpmChip?: boolean

	    startupAuthenticationTpmUsage?: ConfigurationUsage

	    startupAuthenticationTpmPinUsage?: ConfigurationUsage

	    startupAuthenticationTpmKeyUsage?: ConfigurationUsage

	    startupAuthenticationTpmPinAndKeyUsage?: ConfigurationUsage

	    minimumPinLength?: number

	    recoveryOptions?: BitLockerRecoveryOptions

	    prebootRecoveryEnableMessageAndUrl?: boolean

	    prebootRecoveryMessage?: string

	    prebootRecoveryUrl?: string

}
export interface BitLockerRecoveryOptions {

	    blockDataRecoveryAgent?: boolean

	    recoveryPasswordUsage?: ConfigurationUsage

	    recoveryKeyUsage?: ConfigurationUsage

	    hideRecoveryOptions?: boolean

	    enableRecoveryInformationSaveToStore?: boolean

	    recoveryInformationToStore?: BitLockerRecoveryinformationType

	    enableBitLockerAfterRecoveryInformationToStore?: boolean

}
export interface BitLockerFixedDrivePolicy {

	    encryptionMethod?: BitLockerEncryptionMethod

	    requireEncryptionForWriteAccess?: boolean

	    recoveryOptions?: BitLockerRecoveryOptions

}
export interface BitLockerRemovableDrivePolicy {

	    encryptionMethod?: BitLockerEncryptionMethod

	    requireEncryptionForWriteAccess?: boolean

	    blockCrossOrganizationWriteAccess?: boolean

}
export interface DefenderDetectedMalwareActions {

	    lowSeverity?: DefenderThreatAction

	    moderateSeverity?: DefenderThreatAction

	    highSeverity?: DefenderThreatAction

	    severeSeverity?: DefenderThreatAction

}
export interface Windows10NetworkProxyServer {

	    address?: string

	    exceptions?: string[]

	    useForLocalAddresses?: boolean

}
export interface EdgeSearchEngineBase {

}
export interface EdgeSearchEngineCustom extends EdgeSearchEngineBase {

	    edgeSearchEngineOpenSearchXmlUrl?: string

}
export interface EdgeSearchEngine extends EdgeSearchEngineBase {

	    edgeSearchEngineType?: EdgeSearchEngineType

}
export interface SharedPCAccountManagerPolicy {

	    accountDeletionPolicy?: SharedPCAccountDeletionPolicyType

	    cacheAccountsAboveDiskFreePercentage?: number

	    inactiveThresholdDays?: number

	    removeAccountsBelowDiskFreePercentage?: number

}
export interface WindowsUpdateInstallScheduleType {

}
export interface WindowsUpdateScheduledInstall extends WindowsUpdateInstallScheduleType {

	    scheduledInstallDay?: WeeklySchedule

	    scheduledInstallTime?: string

	    restartMode?: WindowsUpdateRestartMode

}
export interface WindowsUpdateActiveHoursInstall extends WindowsUpdateInstallScheduleType {

	    activeHoursStart?: string

	    activeHoursEnd?: string

}
export interface Windows10AssociatedApps {

	    appType?: Windows10AppType

	    identifier?: string

}
export interface VpnTrafficRule {

	    name?: string

	    protocols?: number

	    localPortRanges?: NumberRange[]

	    remotePortRanges?: NumberRange[]

	    localAddressRanges?: IPv4Range[]

	    remoteAddressRanges?: IPv4Range[]

	    appId?: string

	    appType?: VpnTrafficRuleAppType

	    routingPolicyType?: VpnTrafficRuleRoutingPolicyType

	    claims?: string

}
export interface NumberRange {

	    lowerNumber?: number

	    upperNumber?: number

}
export interface VpnRoute {

	    destinationPrefix?: string

	    prefixSize?: number

}
export interface VpnDnsRule {

	    name?: string

	    servers?: string[]

	    proxyServerUri?: string

}
export interface CloudPkiAdministratorCredentials {

	    adminUserName?: string

	    adminPassword?: string

	    authenticationCertificate?: number

	    authenticationCertificatePassword?: string

}
export interface DeviceConfigurationSettingState {

	    setting?: string

	    settingName?: string

	    instanceDisplayName?: string

	    state?: ComplianceStatus

	    errorCode?: number

	    errorDescription?: string

	    userId?: string

	    userName?: string

	    userEmail?: string

	    userPrincipalName?: string

	    sources?: SettingSource[]

	    currentValue?: string

}
export interface SettingSource {

	    id?: string

	    displayName?: string

}
export interface DeviceCompliancePolicySettingState {

	    setting?: string

	    settingName?: string

	    instanceDisplayName?: string

	    state?: ComplianceStatus

	    errorCode?: number

	    errorDescription?: string

	    userId?: string

	    userName?: string

	    userEmail?: string

	    userPrincipalName?: string

	    sources?: SettingSource[]

	    currentValue?: string

}
export interface DeviceManagementExchangeAccessRule {

	    deviceClass?: DeviceManagementExchangeDeviceClass

	    accessLevel?: DeviceManagementExchangeAccessLevel

}
export interface DeviceManagementExchangeDeviceClass {

	    name?: string

	    type?: ExchangeAccessRuleType

}
export interface MobileAppIdentifier {

}
export interface ManagedAppDiagnosticStatus {

	    validationName?: string

	    state?: string

	    mitigationInstruction?: string

}
export interface WindowsInformationProtectionResourceCollection {

	    displayName?: string

	    resources?: string[]

}
export interface WindowsInformationProtectionDataRecoveryCertificate {

	    subjectName?: string

	    description?: string

	    expirationDateTime?: string

	    certificate?: number

}
export interface WindowsInformationProtectionApp {

	    displayName?: string

	    description?: string

	    publisherName?: string

	    productName?: string

}
export interface WindowsInformationProtectionCloudResourceCollection {

	    displayName?: string

	    resources?: WindowsInformationProtectionCloudResource[]

}
export interface WindowsInformationProtectionCloudResource {

	    ipAddressOrFQDN?: string

	    proxy?: string

}
export interface WindowsInformationProtectionIPRangeCollection {

	    displayName?: string

	    ranges?: IpRange[]

}
export interface AndroidMobileAppIdentifier extends MobileAppIdentifier {

	    packageId?: string

}
export interface IosMobileAppIdentifier extends MobileAppIdentifier {

	    bundleId?: string

}
export interface ManagedAppPolicyDeploymentSummaryPerApp {

	    mobileAppIdentifier?: MobileAppIdentifier

	    configurationAppliedUserCount?: number

}
export interface WindowsInformationProtectionStoreApp extends WindowsInformationProtectionApp {

}
export interface WindowsInformationProtectionDesktopApp extends WindowsInformationProtectionApp {

	    binaryName?: string

	    binaryVersionLow?: string

	    binaryVersionHigh?: string

}
export interface RolePermission {

	    actions?: string[]

}
export interface PayloadRequest {

}
