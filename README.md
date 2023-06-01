01June2023
I created this script to help other admins in the situation where it was not enough to use the "standard" attributes to create a Dynamic Distribution Group in Exchange M365

This script takes the input of the different AzureAD attributes and creates the powershell command in a textbox for you to copy and execute, to create the Dynamic Distribution Group. It's easy to create them with attribute equal, Not equal, or like (Remember the *)
All Attributes that are available is 
AdministrativeUnits, Alias, AllowUMCallsFromNonUsers, Altitude, AltSecurityIdentities, ArchiveDatabaseRaw, ArchiveDomain, ArchiveGuid, ArchiveRelease, AssistantName, AuthenticationPolicy
C, CertificateSubject, City, Co, Company, ConfigurationBitmapRaw, ConfigurationUnit, ConfigurationXMLRaw, ConsumerNetID, CorrelationId, CorrelationIdRaw, CountryCode, CountryOrRegion
CustomAttribute1, CustomAttribute10, CustomAttribute11, CustomAttribute12, CustomAttribute13, CustomAttribute14, CustomAttribute15, CustomAttribute2, CustomAttribute3, CustomAttribute4
CustomAttribute5, CustomAttribute6, CustomAttribute7, CustomAttribute8, CustomAttribute9, Database, DataEncryptionPolicy, Department, DesiredMailboxWorkloadsModified, DirectReports
DisabledArchiveGuid, DisplayName, DistinguishedName, ElcMailboxFlags, EmailAddresses, ExchangeGuid, ExchangeObjectId, ExchangeObjectIdRaw, ExchangeUserAccountControl, ExchangeVersion
ExtensionCustomAttribute1, ExtensionCustomAttribute2, ExtensionCustomAttribute3, ExtensionCustomAttribute4, ExtensionCustomAttribute5, ExternalDirectoryObjectId, ExternalDirectoryObjectIdRaw
Fax, FirstName, GroupType, Guid, HABSeniorityIndex, HomePhone, Id, ImmutableSid, InformationBarrierSegments, Initials, InPlaceHoldsRaw, IsDirSynced, IsInactiveMailbox, IsLinked, IsSecurityPrincipal
IsSoftDeletedByDisable, IsSoftDeletedByRemove, LastName, Latitude, LegacyExchangeDN, LegacyExchangeDNRaw, LinkMetadata, Longitude, MailboxDatabasesRaw, MailboxGuidsRaw, MailboxLocationsRaw, MailboxRelease
Manager, MasterAccountSid, MemberOfGroup, MobilePhone, Name, NetID, Notes, NTSecurityDescriptor, ObjectCategory, ObjectClass, Office, OrganizationalUnitRoot, OriginalNetID, OtherFax, OtherHomePhone
OtherTelephone, Pager, PasswordLastSetRaw, Phone, PhoneticDisplayName, PostalCode, PostOfficeBox, PreviousRecipientTypeDetails, ProtocolSettings, ProvisioningFlags, RawCanonicalName, RawName
RecipientSoftDeletedStatus, RecipientType, RecipientTypeDetails, RecipientTypeDetailsValue, RemotePowerShellEnabled, ReplicationSignature, SamAccountName, ServerLegacyDN, ShadowAlias, ShadowAssistantName
ShadowC, ShadowCity, ShadowCo, ShadowCompany, ShadowCountryCode, ShadowDepartment, ShadowDisplayName, ShadowEmailAddresses, ShadowFax, ShadowFirstName, ShadowHomePhone, ShadowInitials, ShadowLastName
ShadowManager, ShadowMobilePhone, ShadowNotes, ShadowOffice, ShadowOtherFax, ShadowOtherHomePhone, ShadowOtherTelephone, ShadowPager, ShadowPhone, ShadowPostalCode, ShadowStateOrProvince, ShadowStreetAddress
ShadowTelephoneAssistant, ShadowTitle, ShadowWebPage, ShadowWhenSoftDeleted, ShadowWindowsLiveID, Sid, SidHistory, SidRaw, SimpleDisplayName, SKUAssigned, StateOrProvince, StreetAddress
StsRefreshTokensValidFrom, TelephoneAssistant, Title, UMCallingLineIds, UMDtmfMap, UMRecipientDialPlanId, UpgradeRequest, UpgradeStatus, UserAccountControl, UserPrincipalName, UserPrincipalNameRaw
VoiceMailSettings, WebPage, WhenChanged, WhenChangedRaw, WhenChangedUTC, WhenCreated, WhenCreatedRaw, WhenCreatedUTC, WhenIBSegmentChanged, WhenSoftDeleted, WindowsEmailAddress, WindowsLiveID.


You still have to install and import the modules and connect to exchange online, See the script for details. And you also need to look into AzureAd to find the value of the attribute :)



You are welcome to cut, Copy, Paste or change the script to your needs, just remember to give me a shout out :) 

Happy DDLing

/Martin Byskov
