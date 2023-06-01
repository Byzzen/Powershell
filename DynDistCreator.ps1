# This script was created by Martin Byskov to help administrators create Dynamic Distribution Lists in Exchange Online.
# Feel free to cut, copy, paste, or modify this script. Just remember to give Martin Byskov ( https://github.com/Byzzen/) a shout-out!


#Install-module -name microsoft.graph.authentication
#Install-module azureadpreview
#Install-Module -Name ExchangeOnlineManagement
#Import-module -name Microsoft.Graph.Authentication
#Import-module -name AzureADPreview
#Connect-AzureAD
#Connect-MgGraph
#Connect-ExchangeOnline


Add-Type -AssemblyName System.Windows.Forms

# Define the list of attributes, of available attributes to build a dynamic distribution list

$attributeList = "AdministrativeUnits, Alias, AllowUMCallsFromNonUsers, Altitude, AltSecurityIdentities, ArchiveDatabaseRaw, ArchiveDomain, ArchiveGuid, ArchiveRelease, AssistantName, AuthenticationPolicy, C, CertificateSubject, City, Co, Company, ConfigurationBitmapRaw, ConfigurationUnit, ConfigurationXMLRaw, ConsumerNetID, CorrelationId, CorrelationIdRaw, CountryCode, CountryOrRegion, CustomAttribute1, CustomAttribute10, CustomAttribute11, CustomAttribute12, CustomAttribute13, CustomAttribute14, CustomAttribute15, CustomAttribute2, CustomAttribute3, CustomAttribute4, CustomAttribute5, CustomAttribute6, CustomAttribute7, CustomAttribute8, CustomAttribute9, Database, DataEncryptionPolicy, Department, DesiredMailboxWorkloadsModified, DirectReports, DisabledArchiveGuid, DisplayName, DistinguishedName, ElcMailboxFlags, EmailAddresses, ExchangeGuid, ExchangeObjectId, ExchangeObjectIdRaw, ExchangeUserAccountControl, ExchangeVersion, ExtensionCustomAttribute1, ExtensionCustomAttribute2, ExtensionCustomAttribute3, ExtensionCustomAttribute4, ExtensionCustomAttribute5, ExternalDirectoryObjectId, ExternalDirectoryObjectIdRaw, Fax, FirstName, GroupType, Guid, HABSeniorityIndex, HomePhone, Id, ImmutableSid, InformationBarrierSegments, Initials, InPlaceHoldsRaw, IsDirSynced, IsInactiveMailbox, IsLinked, IsSecurityPrincipal, IsSoftDeletedByDisable, IsSoftDeletedByRemove, LastName, Latitude, LegacyExchangeDN, LegacyExchangeDNRaw, LinkMetadata, Longitude, MailboxDatabasesRaw, MailboxGuidsRaw, MailboxLocationsRaw, MailboxRelease, Manager, MasterAccountSid, MemberOfGroup, MobilePhone, Name, NetID, Notes, NTSecurityDescriptor, ObjectCategory, ObjectClass, Office, OrganizationalUnitRoot, OriginalNetID, OtherFax, OtherHomePhone, OtherTelephone, Pager, PasswordLastSetRaw, Phone, PhoneticDisplayName, PostalCode, PostOfficeBox, PreviousRecipientTypeDetails, ProtocolSettings, ProvisioningFlags, RawCanonicalName, RawName, RecipientSoftDeletedStatus, RecipientType, RecipientTypeDetails, RecipientTypeDetailsValue, RemotePowerShellEnabled, ReplicationSignature, SamAccountName, ServerLegacyDN, ShadowAlias, ShadowAssistantName, ShadowC, ShadowCity, ShadowCo, ShadowCompany, ShadowCountryCode, ShadowDepartment, ShadowDisplayName, ShadowEmailAddresses, ShadowFax, ShadowFirstName, ShadowHomePhone, ShadowInitials, ShadowLastName, ShadowManager, ShadowMobilePhone, ShadowNotes, ShadowOffice, ShadowOtherFax, ShadowOtherHomePhone, ShadowOtherTelephone, ShadowPager, ShadowPhone, ShadowPostalCode, ShadowStateOrProvince, ShadowStreetAddress, ShadowTelephoneAssistant, ShadowTitle, ShadowWebPage, ShadowWhenSoftDeleted, ShadowWindowsLiveID, Sid, SidHistory, SidRaw, SimpleDisplayName, SKUAssigned, StateOrProvince, StreetAddress, StsRefreshTokensValidFrom, TelephoneAssistant, Title, UMCallingLineIds, UMDtmfMap, UMRecipientDialPlanId, UpgradeRequest, UpgradeStatus, UserAccountControl, UserPrincipalName, UserPrincipalNameRaw, VoiceMailSettings, WebPage, WhenChanged, WhenChangedRaw, WhenChangedUTC, WhenCreated, WhenCreatedRaw, WhenCreatedUTC, WhenIBSegmentChanged, WhenSoftDeleted, WindowsEmailAddress, WindowsLiveID"

# Define the list of operators (Equal, NotEqual, Like)
$operatorList = "-eq", "-ne", "-like"

# Create the form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Byskov Dynamic Distribution Builder"
$form.Width = 550
$form.Height = 700
$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
$form.MaximizeBox = $false

# Create the label
$label = New-Object System.Windows.Forms.Label
$label.Text = "Available attributes:"
$label.AutoSize = $true
$label.Location = New-Object System.Drawing.Point(10, 10)

# Create the attribute dropdown
$attributeDropdown = New-Object System.Windows.Forms.ComboBox
$attributeDropdown.Location = New-Object System.Drawing.Point(10, 30)
$attributeDropdown.Width = 200

# Add attributes to the dropdown
$attributes = $attributeList -split ", "
$attributeDropdown.Items.AddRange($attributes)

# Create the operator label
$operatorLabel = New-Object System.Windows.Forms.Label
$operatorLabel.Text = "Operator:"
$operatorLabel.AutoSize = $true
$operatorLabel.Location = New-Object System.Drawing.Point(220, 30)

# Create the operator dropdown
$operatorDropdown = New-Object System.Windows.Forms.ComboBox
$operatorDropdown.Location = New-Object System.Drawing.Point(310, 30)
$operatorDropdown.Width = 80

# Add operators to the dropdown
$operatorDropdown.Items.AddRange($operatorList)

# Create the value label
$valueLabel = New-Object System.Windows.Forms.Label
$valueLabel.Text = "Value from AzureAD:"
$valueLabel.AutoSize = $true
$valueLabel.Location = New-Object System.Drawing.Point(10, 60)

# Create the value textbox
$valueTextbox = New-Object System.Windows.Forms.TextBox
$valueTextbox.Location = New-Object System.Drawing.Point(150, 60)
$valueTextbox.Width = 200

# Create the filter button
$filterButton = New-Object System.Windows.Forms.Button
$filterButton.Text = "Add Filter"
$filterButton.Location = New-Object System.Drawing.Point(10, 90)
$filterButton.Width = 100

# Create the filter textbox
$filterTextbox = New-Object System.Windows.Forms.TextBox
$filterTextbox.Multiline = $true
$filterTextbox.ScrollBars = "Vertical"
$filterTextbox.ReadOnly = $true
$filterTextbox.Location = New-Object System.Drawing.Point(10, 130)
$filterTextbox.Width = 480
$filterTextbox.Height = 200

# Create the name label
$nameLabel = New-Object System.Windows.Forms.Label
$nameLabel.Text = "Name:"
$nameLabel.AutoSize = $true
$nameLabel.Location = New-Object System.Drawing.Point(10, 340)

# Create the name textbox
$nameTextbox = New-Object System.Windows.Forms.TextBox
$nameTextbox.Location = New-Object System.Drawing.Point(80, 340)
$nameTextbox.Width = 200

# Create the alias label
$aliasLabel = New-Object System.Windows.Forms.Label
$aliasLabel.Text = "Alias:"
$aliasLabel.AutoSize = $true
$aliasLabel.Location = New-Object System.Drawing.Point(10, 370)

# Create the alias textbox
$aliasTextbox = New-Object System.Windows.Forms.TextBox
$aliasTextbox.Location = New-Object System.Drawing.Point(80, 370)
$aliasTextbox.Width = 200

# Create the output button
$outputButton = New-Object System.Windows.Forms.Button
$outputButton.Text = "Generate Output"
$outputButton.Location = New-Object System.Drawing.Point(10, 400)
$outputButton.Width = 100

# Create the output textbox
$outputTextbox = New-Object System.Windows.Forms.TextBox
$outputTextbox.Multiline = $true
$outputTextbox.ScrollBars = "Vertical"
$outputTextbox.ReadOnly = $true
$outputTextbox.Location = New-Object System.Drawing.Point(10, 440)
$outputTextbox.Width = 480
$outputTextbox.Height = 200

# Define the event handler for the filter button
$filterButton.Add_Click({
    $attribute = $attributeDropdown.SelectedItem.ToString()
    $operator = $operatorDropdown.SelectedItem.ToString()
    $value = $valueTextbox.Text

    $filter = "($attribute $operator '$value')"
    $filterTextbox.AppendText("$filter -and `r`n")
})

# Define the event handler for the output button
$outputButton.Add_Click({
    $filters = $filterTextbox.Text.TrimEnd("-and `r`n")

    $name = $nameTextbox.Text.Trim()
    $alias = $aliasTextbox.Text.Trim()

    $command = "New-DynamicDistributionGroup -Name '$name' -Alias '$alias' -RecipientFilter $filters"
    $outputTextbox.Text = $command
})

# Add controls to the form
$form.Controls.Add($label)
$form.Controls.Add($attributeDropdown)
$form.Controls.Add($operatorLabel)
$form.Controls.Add($operatorDropdown)
$form.Controls.Add($valueLabel)
$form.Controls.Add($valueTextbox)
$form.Controls.Add($filterButton)
$form.Controls.Add($filterTextbox)
$form.Controls.Add($nameLabel)
$form.Controls.Add($nameTextbox)
$form.Controls.Add($aliasLabel)
$form.Controls.Add($aliasTextbox)
$form.Controls.Add($outputButton)
$form.Controls.Add($outputTextbox)

# Show the form
$form.ShowDialog()
