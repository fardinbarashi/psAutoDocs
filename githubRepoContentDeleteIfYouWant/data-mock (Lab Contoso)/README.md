# Autodoc export - data-mock (Lab Contoso)

Generated: 2026-08-02 20:00
Tenant: Lab Contoso  (labcontoso.onmicrosoft.com)

This folder is a single export - one point-in-time pull of the tenant. Everything here (CSV, Excel, maps) is built from the raw JSON in the `json\` sub-folder, so the JSON is the source of truth.

## Folders

- **json\** - Raw data pulled from Microsoft Graph, one file per area (each file is documented below).
- **excelkpi\** - The KPI workbook "KPI <stamp>.xlsx": dashboard, charts and one analysed sheet per area.
- **svg\** - Vector maps of the tenant (open in a web browser). That folder has its own README.

## JSON files (json\)

Each file is an export of one area. The fields listed under each file are the columns every record carries; a dash means the field name is self-explanatory.

### AppRegEnterpriseApps
App registrations joined to their enterprise app (service principal): ownership, credentials, API permissions and sign-in recency.

- `AppName` : Display name of the application
- `AppId` : Application (client) ID
- `ObjectId` : Directory object ID of the app registration
- `ServicePrincipalObjectId` : Object ID of the enterprise app (service principal)
- `OwnerInfo` : Owners, or MissingOwners when none set
- `HasSecrets` : Whether the app has client secrets
- `HasCertificates` : Whether the app has certificates
- `SecretCount` : Number of client secrets
- `CertificateCount` : Number of certificates
- `SecretExpiryDates` : Expiry date(s) of client secrets
- `CertificateExpiryDates` : Expiry date(s) of certificates
- `NearestSecretExpiry` : Soonest secret expiry
- `NearestCertificateExpiry` : Soonest certificate expiry
- `NearestCredentialExpiry` : Soonest of any credential expiry
- `CredentialExpired` : Whether any credential has already expired
- `CredentialExpiresWithin30Days` : Whether a credential expires within 30 days
- `ServicePrincipalExists` : Whether a service principal exists for the app
- `SSOConfigured` : Whether single sign-on is configured
- `SSOType` : Single sign-on mode (SAML, OIDC, Password, ...)
- `PublicClient` : Whether the app is a public (native) client
- `SignInAudience` : Who can sign in (single tenant, multi-tenant, personal)
- `UsesGraphPermissions` : Whether the app holds Microsoft Graph permissions
- `ApiPermissions` : Raw API permission IDs granted
- `ApiPermissionNames` : Readable API permissions granted
- `ApiPermissionCount` : Number of API permissions
- `LastSignInDateTime` : Most recent sign-in seen for the app
- `UsedWithin30Days` : Whether the app was used in the last 30 days
- `LastDelegatedClientSignIn` : Most recent delegated (user) sign-in
- `LastApplicationAuthSignIn` : Most recent app-only sign-in

### ConditionalAccess
Conditional Access policies: state, grant controls (MFA, ...), and targeted / excluded users, groups and roles.

- `RecordType` : Internal record type marker
- `PolicyId` : Conditional Access policy ID
- `DisplayName` : Display name
- `State` : Policy state (enabled, disabled, report-only)
- `IncludeUsersRaw` : Targeted users (raw object IDs)
- `IncludeUsersResolved` : Targeted users (resolved names)
- `ExcludeUsersRaw` : Excluded users (raw object IDs)
- `ExcludeUsersResolved` : Excluded users (resolved names)
- `IncludeGroupsRaw` : Targeted groups (raw object IDs)
- `IncludeGroupsResolved` : Targeted groups (resolved names)
- `ExcludeGroupsRaw` : Excluded groups (raw object IDs)
- `ExcludeGroupsResolved` : Excluded groups (resolved names)
- `IncludeRolesRaw` : Targeted directory roles (raw IDs)
- `IncludeRolesResolved` : Targeted directory roles (resolved names)
- `ExcludeRolesRaw` : Excluded directory roles (raw IDs)
- `ExcludeRolesResolved` : Excluded directory roles (resolved names)
- `IncludeApplicationsRaw` : Targeted apps (raw IDs, or All)
- `IncludeApplicationsResolved` : Targeted apps (resolved names)
- `ExcludeApplicationsRaw` : Excluded apps (raw IDs)
- `ExcludeApplicationsResolved` : Excluded apps (resolved names)
- `IncludeUserActionsRaw` : Targeted user actions (e.g. register security info)
- `IncludeLocationsRaw` : Targeted locations (raw IDs, or All)
- `IncludeLocationsResolved` : Targeted locations (resolved names)
- `ExcludeLocationsRaw` : Excluded locations (raw IDs)
- `ExcludeLocationsResolved` : Excluded locations (resolved names)
- `ClientAppTypes` : Client app types the policy applies to
- `IncludePlatforms` : Targeted device platforms
- `ExcludePlatforms` : Excluded device platforms
- `SignInRiskLevels` : Sign-in risk levels that trigger the policy
- `UserRiskLevels` : User risk levels that trigger the policy
- `ServicePrincipalRiskLevels` : Workload-identity risk levels that trigger the policy
- `DeviceFilterMode` : Device filter mode (include / exclude)
- `DeviceFilterRule` : Device filter rule expression
- `GrantOperator` : How grant controls combine (AND / OR)
- `BuiltInControls` : Required grant controls (MFA, compliant device, ...)
- `CustomAuthenticationFactors` : Custom authentication factors required
- `TermsOfUse` : Terms of use that must be accepted
- `AuthenticationStrengthId` : Authentication strength ID
- `AuthenticationStrengthResolved` : Authentication strength (resolved name)
- `SignInFrequencyType` : Sign-in frequency unit (hours / days)
- `SignInFrequencyValue` : Sign-in frequency value
- `PersistentBrowserMode` : Persistent browser session control
- `AppEnforcedRestrictionsEnabled` : Whether app-enforced restrictions are on
- `CloudAppSecurityType` : Conditional Access App Control (MCAS) mode
- `NamedLocationId` : Named location ID
- `NamedLocationType` : Named location type (IP / country)
- `CountriesAndRegions` : Countries/regions for the named location
- `IncludeUnknownCountriesAndRegions` : Whether unknown countries are included
- `IpRanges` : IP ranges for the named location
- `IsTrusted` : Whether the location is marked trusted
- `PolicyType` : Policy type (builtIn / custom)
- `RequirementsSatisfied` : Requirements the auth strength satisfies
- `AllowedCombinations` : Allowed authentication method combinations
- `CreatedDateTime` : When the object was created
- `ModifiedDateTime` : When the object was last modified

### EnterpriseApps
Enterprise applications (service principals): single sign-on, Microsoft Graph usage and client type.

- `AppName` : Display name of the application
- `AppId` : Application (client) ID
- `ServicePrincipalObjectId` : Object ID of the enterprise app (service principal)
- `ServicePrincipalType` : SP kind: Application, ManagedIdentity, Legacy, SocialIdp
- `AppOrigin` : Home of the app: this tenant vs external/Microsoft
- `HasAppRegistration` : Whether an app registration exists for this SP
- `AccountEnabled` : Whether the account/app is enabled
- `OwnerInfo` : Owners, or MissingOwners when none set
- `SSOConfigured` : Whether single sign-on is configured
- `SSOType` : Single sign-on mode (SAML, OIDC, Password, ...)
- `AppRoleAssignmentRequired` : Whether user assignment is required to sign in
- `Homepage` : App home page URL
- `Tags` : Service principal tags
- `SignInAudience` : Who can sign in (single tenant, multi-tenant, personal)
- `HasSecrets` : Whether the app has client secrets
- `HasCertificates` : Whether the app has certificates
- `SecretCount` : Number of client secrets
- `CertificateCount` : Number of certificates
- `SecretExpiryDates` : Expiry date(s) of client secrets
- `CertificateExpiryDates` : Expiry date(s) of certificates
- `NearestSecretExpiry` : Soonest secret expiry
- `NearestCertificateExpiry` : Soonest certificate expiry
- `NearestCredentialExpiry` : Soonest of any credential expiry
- `CredentialExpired` : Whether any credential has already expired
- `CredentialExpiresWithin30Days` : Whether a credential expires within 30 days
- `LastSignInDateTime` : Most recent sign-in seen for the app
- `UsedWithin30Days` : Whether the app was used in the last 30 days
- `LastDelegatedClientSignIn` : Most recent delegated (user) sign-in
- `LastApplicationAuthSignIn` : Most recent app-only sign-in

### EntraGroups
All groups: type, category, owners, member count, dynamic membership rule and created date.

- `Id` : Directory object ID
- `DisplayName` : Display name
- `Description` : Free-text description
- `Source` : Where the object is mastered (Cloud or on-premises AD)
- `GroupCategory` : Group category (Security, Microsoft 365, ...)
- `CreatedDateTime` : When the object was created
- `GroupTypes` : Underlying group type flags (e.g. Unified, DynamicMembership)
- `MailEnabled` : Whether the group has a mailbox / can receive mail
- `SecurityEnabled` : Whether the group can be used to grant access (security group)
- `MailNickname` : Mail alias (nickname) of the group
- `Visibility` : Group visibility (Public, Private, HiddenMembership)
- `IsAssignableToRole` : Whether the group can hold directory roles
- `DynamicGroupMembershipRule` : Rule that decides membership for a dynamic group
- `DynamicMembershipRuleProcessingState` : Whether the dynamic membership rule is running or paused
- `ResourceProvisioningOptions` : Extra provisioning options (e.g. Team)
- `OnPremisesSyncEnabled` : Whether the object is synced from on-prem AD
- `Owners` : Group owners
- `Licenses` : Licences assigned to the group (group-based licensing)
- `RolesAndAdministrators` : Directory roles assigned to the group
- `AdministrativeUnits` : Administrative units the group belongs to
- `GroupMemberships` : Groups this group is a member of (nested membership)
- `Applications` : Applications assigned to the group
- `MemberCount` : Number of members
- `WelcomeEmailEnabled` : Whether a welcome email is sent to new members

### EntraGroupsBasicInfo
Aggregate group counts for the tenant (totals per type and per source).

- `TotalGroups` : Total number of groups in the tenant
- `M365Groups` : Number of Microsoft 365 groups
- `SecurityGroups` : Number of security groups
- `DynamicGroups` : Number of dynamic-membership groups
- `CloudGroups` : Number of cloud-only groups
- `OnPremisesGroups` : Number of groups synced from on-premises AD

### EntraServiceLimits

- `key` : -
- `area` : -
- `label` : -
- `limit` : -
- `type` : Domain type (Managed, Federated, None)
- `note` : -

### GroupMembers
One row per group membership - the member together with their department, office location and manager.

- `GroupId` : Directory object ID of the group
- `GroupDisplayName` : Display name of the group
- `UserId` : Directory object ID of the member
- `DisplayName` : Display name
- `UserPrincipalName` : User principal name (sign-in name)
- `Department` : Department
- `OfficeLocation` : Office location
- `Manager` : The user's manager
- `ManagerDisplayName` : Display name of the manager

### GroupWelcomeEmail
Per-group welcome-email setting (whether a welcome mail is sent to new members).

- `GroupId` : Directory object ID of the group
- `GroupDisplayName` : Display name of the group
- `GroupType` : Group type (security, Microsoft 365, ...)
- `WelcomeEmailEnabled` : Whether a welcome email is sent to new members
- `MemberCount` : Number of members

### LicensesInformation
Licence SKUs (subscriptions): enabled / consumed / free units and the resulting utilisation.

- `skuId` : License SKU ID
- `skuPartNumber` : License SKU part number
- `enabledUnits` : Number of licence units purchased / enabled
- `consumedUnits` : Licenses assigned
- `freeUnits` : Number of unused (free) licence units
- `capabilityStatus` : Subscription status (Enabled, Warning, Suspended)

### PasswordReset
Self-service password reset (SSPR) configuration for the tenant.

- `SsprEnabledScope` : Who self-service password reset is enabled for
- `SsprSelectedGroupId` : Group scoped for self-service password reset
- `MethodsAvailableToUsers` : Authentication methods available for reset
- `SecurityQuestionsEnabled` : Whether security questions are enabled
- `AdministratorSsprEnabled` : Whether SSPR is enabled for administrators
- `AdministratorMethodsRequired` : Number of methods admins must register for SSPR
- `AdministratorMethodsAvailable` : Authentication methods available to admins for SSPR

### RBAC
Directory role assignments (active and eligible) with the principal that holds each role.

- `RecordType` : Internal record type marker
- `RoleDefinitionName` : Directory role name
- `RoleDefinitionId` : Directory role definition ID
- `PrincipalId` : Object ID of the principal holding the role
- `PrincipalDisplayName` : Assigned principal (user/group/SP) name
- `PrincipalType` : Assigned principal type
- `AssignmentType` : Role assignment type (active / eligible)
- `MemberType` : How the role is held (Direct, Group, Inherited)
- `Status` : Status of the assignment / record
- `StartDateTime` : When the assignment starts (time-bound roles)
- `EndDateTime` : When the assignment ends (blank = permanent)
- `JITActivated` : Whether an eligible role is currently activated (just-in-time)

### TenantInformation
Tenant identity: name, tenant/organisation IDs, primary domain and directory-sync status.

- `TenantId` : Directory (tenant) ID
- `TenantName` : Tenant display name
- `PrimaryDomain` : Primary (default) domain
- `CountryCode` : Tenant country code
- `NotificationLanguage` : Default notification language for the tenant
- `TechnicalContact` : Technical contact address(es)
- `GlobalPrivacyContact` : Privacy contact for the tenant
- `PrivacyStatementUrl` : URL of the tenant privacy statement
- `Users` : Total number of users in the tenant
- `Groups` : Total number of groups in the tenant
- `AppRegistrations` : Total number of app registrations
- `Devices` : Total number of registered devices
- `HybridStatus` : Directory hybrid status (cloud-only vs synced with on-prem AD)

### TenantInformationDomain
Verified domains in the tenant, with each domain's type and whether it is the default.

- `Domain` : Domain name
- `Type` : Domain type (Managed, Federated, None)
- `IsDefault` : Whether this is the default domain

### UserInformation
Every user account: type (member/guest), status, department, office, licences and authentication methods.

- `Id` : Directory object ID
- `DisplayName` : Display name
- `UserPrincipalName` : User principal name (sign-in name)
- `UserType` : Account type (Member or Guest)
- `UserCategory` : Derived user category (e.g. internal / external)
- `Mail` : Primary email address
- `AccountEnabled` : Whether the account/app is enabled
- `AccountStatus` : Account status text (Enabled / Disabled)
- `AssignedLicenses` : Licenses assigned to the user
- `AssignedPlans` : Service plans assigned to the user
- `AssignedLicenseCount` : Number of licences assigned to the user
- `AssignedPlanCount` : Number of service plans assigned to the user
- `OnPremisesSyncEnabled` : Whether the object is synced from on-prem AD
- `MobilePhone` : Mobile phone number
- `JobTitle` : Job title
- `Department` : Department
- `CompanyName` : Company name
- `OfficeLocation` : Office location
- `Manager` : The user's manager
- `ManagerDisplayName` : Display name of the manager
- `ManagerId` : Object ID of the user's manager
- `AuthenticationMethods` : Registered authentication methods (MFA, ...)
- `AuthenticationMethodCount` : Number of registered authentication methods


