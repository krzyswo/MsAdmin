# MsAdmin documentation index

This repository contains a large collection of PowerShell utilities for Microsoft 365 administration. This root README centralizes discovery by linking to every script and reserving space for shared metadata (workload, authentication, and key parameters). Once scripts adopt a common metadata block, the table below can be regenerated automatically to include those fields.

## Directory naming improvements
- **Adopt lowercase kebab-case**: use lowercase letters, digits, and hyphens instead of spaces (for example, `meeting-room-creation-automation` instead of `Meeting Room Creation Automation`).
- **Keep names descriptive but concise**: start with the action (export, audit, manage), then the workload (exo/graph/spo/teams), followed by the target (users, groups, mailboxes).
- **Avoid trailing spaces and special characters**: remove trailing spaces and prefer hyphens over underscores or mixed separators.
- **One script per folder**: keep the folder name aligned with the primary script name so automation can locate the correct README/PS1 pair.

> Migration tip: rename folders with `Rename-Item` (PowerShell) or `mv` (bash) and update references in documentation/automation pipelines to the new paths. Consider staging renames in a branch to avoid breaking existing links.

## Script index
The columns below are placeholders until each script adopts a shared metadata block (recommended fields: Workload, Purpose, Required Modules, Authentication, Key Parameters).

| Script | Workload | Purpose | Authentication | Key parameters |
| --- | --- | --- | --- | --- |
| [archive-inactive-teams-general](archive-inactive-teams-general/archive-inactive-teams-general.ps1) | TBD | See script README | TBD | TBD |
| [assign-office365-user-manager](assign-office365-user-manager/assign-office365-user-manager.ps1) | TBD | See script README | TBD | TBD |
| [audit-entra-app-operations](audit-entra-app-operations/audit-entra-app-operations.ps1) | TBD | See script README | TBD | TBD |
| [audit-group-membership-changes](audit-group-membership-changes/audit-group-membership-changes.ps1) | TBD | See script README | TBD | TBD |
| [audit-m365-user-creations](audit-m365-user-creations/audit-m365-user-creations.ps1) | TBD | See script README | TBD | TBD |
| [audit-mailbox-mailbox-changes](audit-mailbox-mailbox-changes/audit-mailbox-mailbox-changes.ps1) | TBD | See script README | TBD | TBD |
| [audit-ms-teams-channel-creations](audit-ms-teams-channel-creations/audit-ms-teams-channel-creations.ps1) | TBD | See script README | TBD | TBD |
| [audit-send-as-emails](audit-send-as-emails/audit-send-as-emails.ps1) | TBD | See script README | TBD | TBD |
| [audit-shared-mailbox-activities](audit-shared-mailbox-activities/audit-shared-mailbox-activities.ps1) | TBD | See script README | TBD | TBD |
| [audit-shared-mailbox-email-deletions](audit-shared-mailbox-email-deletions/audit-shared-mailbox-email-deletions.ps1) | TBD | See script README | TBD | TBD |
| [audit-shared-mailbox-email-sender](audit-shared-mailbox-email-sender/audit-shared-mailbox-email-sender.ps1) | TBD | See script README | TBD | TBD |
| [audit-sharepoint-file-access](audit-sharepoint-file-access/audit-sharepoint-file-access.ps1) | TBD | See script README | TBD | TBD |
| [audit-sharepoint-file-downloads](audit-sharepoint-file-downloads/audit-sharepoint-file-downloads.ps1) | TBD | See script README | TBD | TBD |
| [audit-sharepoint-folder-activities](audit-sharepoint-folder-activities/audit-sharepoint-folder-activities.ps1) | TBD | See script README | TBD | TBD |
| [audit-spo-group-membership-changes](audit-spo-group-membership-changes/audit-spo-group-membership-changes.ps1) | TBD | See script README | TBD | TBD |
| [audit-teams-file-sharing-activities](audit-teams-file-sharing-activities/audit-teams-file-sharing-activities.ps1) | TBD | See script README | TBD | TBD |
| [audit-teams-membership-changes](audit-teams-membership-changes/audit-teams-membership-changes.ps1) | TBD | See script README | TBD | TBD |
| [Automaping_check](Automaping_check/Automaping_check.ps1) | TBD | See script README | TBD | TBD |
| [block-external-email-forwarding](block-external-email-forwarding/block-external-email-forwarding.ps1) | TBD | See script README | TBD | TBD |
| [block-signin-shared-resource-mailboxes](block-signin-shared-resource-mailboxes/block-signin-shared-resource-mailboxes.ps1) | TBD | See script README | TBD | TBD |
| [Check Automapping for Mailbox ](Check Automapping for Mailbox /Check Automapping for Mailbox .ps1) | TBD | See script README | TBD | TBD |
| [Check Mailbox inbox folder size](Check Mailbox inbox folder size/Check Mailbox inbox folder size.ps1) | TBD | See script README | TBD | TBD |
| [CheckIPv6StatusForAllDomains](CheckIPv6StatusForAllDomains/CheckIPv6StatusForAllDomains.ps1) | TBD | See script README | TBD | TBD |
| [configure-external-email-warning](configure-external-email-warning/configure-external-email-warning.ps1) | TBD | See script README | TBD | TBD |
| [configure-sent-mailbox-in-shared-mailbox](configure-sent-mailbox-in-shared-mailbox/configure-sent-mailbox-in-shared-mailbox.ps1) | TBD | See script README | TBD | TBD |
| [connect-exchange-online-general](connect-exchange-online-general/connect-exchange-online-general.ps1) | TBD | See script README | TBD | TBD |
| [connect-microsoft-graph-sdk](connect-microsoft-graph-sdk/connect-microsoft-graph-sdk.ps1) | TBD | See script README | TBD | TBD |
| [convert-dl-to-m365-group](convert-dl-to-m365-group/convert-dl-to-m365-group.ps1) | TBD | See script README | TBD | TBD |
| [convert-user-mailbox-to-shared](convert-user-mailbox-to-shared/convert-user-mailbox-to-shared.ps1) | TBD | See script README | TBD | TBD |
| [copy-distribution-list-members-and-owners](copy-distribution-list-members-and-owners/copy-distribution-list-members-and-owners.ps1) | TBD | See script README | TBD | TBD |
| [delete-older-emails-general](delete-older-emails-general/delete-older-emails-general.ps1) | TBD | See script README | TBD | TBD |
| [DirectLicenseAssigmentreport](DirectLicenseAssigmentreport/DirectLicenseAssigmentreport.ps1) | TBD | See script README | TBD | TBD |
| [disable-self-service-purchase](disable-self-service-purchase/disable-self-service-purchase.ps1) | TBD | See script README | TBD | TBD |
| [DisableIPv6ForAllDomains](DisableIPv6ForAllDomains/DisableIPv6ForAllDomains.ps1) | TBD | See script README | TBD | TBD |
| [email-report-report-export](email-report-report-export/email-report-report-export.ps1) | TBD | See script README | TBD | TBD |
| [Enable-ExchangeLitigationHold](Enable-ExchangeLitigationHold/Enable-ExchangeLitigationHold.ps1) | TBD | See script README | TBD | TBD |
| [enable-mailbox-mailbox-logging](enable-mailbox-mailbox-logging/enable-mailbox-mailbox-logging.ps1) | TBD | See script README | TBD | TBD |
| [enable-mfa-for-office365-admins](enable-mfa-for-office365-admins/enable-mfa-for-office365-admins.ps1) | TBD | See script README | TBD | TBD |
| [enabling autoarchiving and running Folder Assistant](enabling autoarchiving and running Folder Assistant/enabling autoarchiving and running Folder Assistant.ps1) | TBD | See script README | TBD | TBD |
| [Exchange Mailbox Capacity Management Tool](Exchange Mailbox Capacity Management Tool/Exchange Mailbox Capacity Management Tool.ps1) | TBD | See script README | TBD | TBD |
| [exchange-report-mailbox-statistics-export](exchange-report-mailbox-statistics-export/exchange-report-mailbox-statistics-export.ps1) | TBD | See script README | TBD | TBD |
| [ExchangeOnlineResource Mailbox Settings Report](ExchangeOnlineResource Mailbox Settings Report/ExchangeOnlineResource Mailbox Settings Report.ps1) | TBD | See script README | TBD | TBD |
| [EXO REPORTS](EXO REPORTS/EXO REPORTS.ps1) | TBD | See script README | TBD | TBD |
| [Export Exchange Online Distribution Group Members to CSV](Export Exchange Online Distribution Group Members to CSV/Export Exchange Online Distribution Group Members to CSV.ps1) | TBD | See script README | TBD | TBD |
| [export-report-365-external-sharing-report](export-report-365-external-sharing-report/export-report-365-external-sharing-report.ps1) | TBD | See script README | TBD | TBD |
| [export-report-365-guest-user-report](export-report-365-guest-user-report/export-report-365-guest-user-report.ps1) | TBD | See script README | TBD | TBD |
| [export-report-365-manager-report](export-report-365-manager-report/export-report-365-manager-report.ps1) | TBD | See script README | TBD | TBD |
| [export-report-activity-report](export-report-activity-report/export-report-activity-report.ps1) | TBD | See script README | TBD | TBD |
| [export-report-ad-devices-report](export-report-ad-devices-report/export-report-ad-devices-report.ps1) | TBD | See script README | TBD | TBD |
| [export-report-admin-report](export-report-admin-report/export-report-admin-report.ps1) | TBD | See script README | TBD | TBD |
| [export-report-anonymous-links-report](export-report-anonymous-links-report/export-report-anonymous-links-report.ps1) | TBD | See script README | TBD | TBD |
| [export-report-apps-report](export-report-apps-report/export-report-apps-report.ps1) | TBD | See script README | TBD | TBD |
| [export-report-based-license-report](export-report-based-license-report/export-report-based-license-report.ps1) | TBD | See script README | TBD | TBD |
| [export-report-channels-external-members](export-report-channels-external-members/export-report-channels-external-members.ps1) | TBD | See script README | TBD | TBD |
| [export-report-distribution-group-members](export-report-distribution-group-members/export-report-distribution-group-members.ps1) | TBD | See script README | TBD | TBD |
| [export-report-document-libraries-report](export-report-document-libraries-report/export-report-document-libraries-report.ps1) | TBD | See script README | TBD | TBD |
| [export-report-email-addresses](export-report-email-addresses/export-report-email-addresses.ps1) | TBD | See script README | TBD | TBD |
| [export-report-emails-audit-report](export-report-emails-audit-report/export-report-emails-audit-report.ps1) | TBD | See script README | TBD | TBD |
| [export-report-expiry-report](export-report-expiry-report/export-report-expiry-report.ps1) | TBD | See script README | TBD | TBD |
| [export-report-external-user-file-access-report](export-report-external-user-file-access-report/export-report-external-user-file-access-report.ps1) | TBD | See script README | TBD | TBD |
| [export-report-external-users-report](export-report-external-users-report/export-report-external-users-report.ps1) | TBD | See script README | TBD | TBD |
| [export-report-file-version-history-report](export-report-file-version-history-report/export-report-file-version-history-report.ps1) | TBD | See script README | TBD | TBD |
| [export-report-files-audit-report](export-report-files-audit-report/export-report-files-audit-report.ps1) | TBD | See script README | TBD | TBD |
| [export-report-forwarding-report](export-report-forwarding-report/export-report-forwarding-report.ps1) | TBD | See script README | TBD | TBD |
| [export-report-group-members-report](export-report-group-members-report/export-report-group-members-report.ps1) | TBD | See script README | TBD | TBD |
| [export-report-group-membership-report](export-report-group-membership-report/export-report-group-membership-report.ps1) | TBD | See script README | TBD | TBD |
| [export-report-group-membership](export-report-group-membership/export-report-group-membership.ps1) | TBD | See script README | TBD | TBD |
| [export-report-groups-report](export-report-groups-report/export-report-groups-report.ps1) | TBD | See script README | TBD | TBD |
| [export-report-groups-reports](export-report-groups-reports/export-report-groups-reports.ps1) | TBD | See script README | TBD | TBD |
| [export-report-guest-details](export-report-guest-details/export-report-guest-details.ps1) | TBD | See script README | TBD | TBD |
| [export-report-guest-users-last-login-report](export-report-guest-users-last-login-report/export-report-guest-users-last-login-report.ps1) | TBD | See script README | TBD | TBD |
| [export-report-inactive-users-report](export-report-inactive-users-report/export-report-inactive-users-report.ps1) | TBD | See script README | TBD | TBD |
| [export-report-items-by-external-users](export-report-items-by-external-users/export-report-items-by-external-users.ps1) | TBD | See script README | TBD | TBD |
| [export-report-license-assignment-report](export-report-license-assignment-report/export-report-license-assignment-report.ps1) | TBD | See script README | TBD | TBD |
| [export-report-list-item-counts](export-report-list-item-counts/export-report-list-item-counts.ps1) | TBD | See script README | TBD | TBD |
| [export-report-mailbox-mailbox-access-report](export-report-mailbox-mailbox-access-report/export-report-mailbox-mailbox-access-report.ps1) | TBD | See script README | TBD | TBD |
| [export-report-mailbox-mailbox-activities](export-report-mailbox-mailbox-activities/export-report-mailbox-mailbox-activities.ps1) | TBD | See script README | TBD | TBD |
| [export-report-mailbox-permissions-report](export-report-mailbox-permissions-report/export-report-mailbox-permissions-report.ps1) | TBD | See script README | TBD | TBD |
| [export-report-mailbox-report](export-report-mailbox-report/export-report-mailbox-report.ps1) | TBD | See script README | TBD | TBD |
| [export-report-mailbox-reports](export-report-mailbox-reports/export-report-mailbox-reports.ps1) | TBD | See script README | TBD | TBD |
| [export-report-mailbox-size-report](export-report-mailbox-size-report/export-report-mailbox-size-report.ps1) | TBD | See script README | TBD | TBD |
| [export-report-mailboxes-report](export-report-mailboxes-report/export-report-mailboxes-report.ps1) | TBD | See script README | TBD | TBD |
| [export-report-meeting-attendance-report](export-report-meeting-attendance-report/export-report-meeting-attendance-report.ps1) | TBD | See script README | TBD | TBD |
| [export-report-meeting-reports](export-report-meeting-reports/export-report-meeting-reports.ps1) | TBD | See script README | TBD | TBD |
| [export-report-permissions-report](export-report-permissions-report/export-report-permissions-report.ps1) | TBD | See script README | TBD | TBD |
| [export-report-registrations-certificates-secrets](export-report-registrations-certificates-secrets/export-report-registrations-certificates-secrets.ps1) | TBD | See script README | TBD | TBD |
| [export-report-registrations-expiry-report](export-report-registrations-expiry-report/export-report-registrations-expiry-report.ps1) | TBD | See script README | TBD | TBD |
| [export-report-sharepoint-links](export-report-sharepoint-links/export-report-sharepoint-links.ps1) | TBD | See script README | TBD | TBD |
| [export-report-signin-report](export-report-signin-report/export-report-signin-report.ps1) | TBD | See script README | TBD | TBD |
| [export-report-status-report](export-report-status-report/export-report-status-report.ps1) | TBD | See script README | TBD | TBD |
| [export-report-storage-consumption-report](export-report-storage-consumption-report/export-report-storage-consumption-report.ps1) | TBD | See script README | TBD | TBD |
| [export-report-subscription-expiry-report](export-report-subscription-expiry-report/export-report-subscription-expiry-report.ps1) | TBD | See script README | TBD | TBD |
| [export-report-subsites-report](export-report-subsites-report/export-report-subsites-report.ps1) | TBD | See script README | TBD | TBD |
| [export-report-teams-private-channel-reports](export-report-teams-private-channel-reports/export-report-teams-private-channel-reports.ps1) | TBD | See script README | TBD | TBD |
| [export-report-teams-report](export-report-teams-report/export-report-teams-report.ps1) | TBD | See script README | TBD | TBD |
| [export-report-traffic-statistics](export-report-traffic-statistics/export-report-traffic-statistics.ps1) | TBD | See script README | TBD | TBD |
| [export-report-urls-and-storage-report](export-report-urls-and-storage-report/export-report-urls-and-storage-report.ps1) | TBD | See script README | TBD | TBD |
| [export-report-user-activity-report](export-report-user-activity-report/export-report-user-activity-report.ps1) | TBD | See script README | TBD | TBD |
| [export-report-user-audit-log](export-report-user-audit-log/export-report-user-audit-log.ps1) | TBD | See script README | TBD | TBD |
| [export-report-user-last-logon-report](export-report-user-last-logon-report/export-report-user-last-logon-report.ps1) | TBD | See script README | TBD | TBD |
| [export-report-user-license-report](export-report-user-license-report/export-report-user-license-report.ps1) | TBD | See script README | TBD | TBD |
| [export-report-user-login-history](export-report-user-login-history/export-report-user-login-history.ps1) | TBD | See script README | TBD | TBD |
| [export-report-user-membership-report](export-report-user-membership-report/export-report-user-membership-report.ps1) | TBD | See script README | TBD | TBD |
| [export-report-user-signin-report](export-report-user-signin-report/export-report-user-signin-report.ps1) | TBD | See script README | TBD | TBD |
| [export-report-users-last-successful-signin-report](export-report-users-last-successful-signin-report/export-report-users-last-successful-signin-report.ps1) | TBD | See script README | TBD | TBD |
| [export-report-users-membership-reports](export-report-users-membership-reports/export-report-users-membership-reports.ps1) | TBD | See script README | TBD | TBD |
| [export-report-users-registered-auth-methods](export-report-users-registered-auth-methods/export-report-users-registered-auth-methods.ps1) | TBD | See script README | TBD | TBD |
| [export-report-users-report](export-report-users-report/export-report-users-report.ps1) | TBD | See script README | TBD | TBD |
| [export-report-with-disabled-users-report](export-report-with-disabled-users-report/export-report-with-disabled-users-report.ps1) | TBD | See script README | TBD | TBD |
| [find-inactive-distribution-lists](find-inactive-distribution-lists/find-inactive-distribution-lists.ps1) | TBD | See script README | TBD | TBD |
| [find-inactive-teams-general](find-inactive-teams-general/find-inactive-teams-general.ps1) | TBD | See script README | TBD | TBD |
| [find-inbox-rules-with-external-forwarding](find-inbox-rules-with-external-forwarding/find-inbox-rules-with-external-forwarding.ps1) | TBD | See script README | TBD | TBD |
| [generate-password-expiry-reports](generate-password-expiry-reports/generate-password-expiry-reports.ps1) | TBD | See script README | TBD | TBD |
| [generate-report-shared-channel-report](generate-report-shared-channel-report/generate-report-shared-channel-report.ps1) | TBD | See script README | TBD | TBD |
| [get-distribution-lists-external-users](get-distribution-lists-external-users/get-distribution-lists-external-users.ps1) | TBD | See script README | TBD | TBD |
| [get-report-365-groups-storage-report](get-report-365-groups-storage-report/get-report-365-groups-storage-report.ps1) | TBD | See script README | TBD | TBD |
| [get-report-mailbox-permission-report](get-report-mailbox-permission-report/get-report-mailbox-permission-report.ps1) | TBD | See script README | TBD | TBD |
| [identify-non-compliant-shared-mailboxes](identify-non-compliant-shared-mailboxes/identify-non-compliant-shared-mailboxes.ps1) | TBD | See script README | TBD | TBD |
| [install-and-connect-exo-module](install-and-connect-exo-module/install-and-connect-exo-module.ps1) | TBD | See script README | TBD | TBD |
| [Install-ReqModules_AdminCheck](Install-ReqModules_AdminCheck/Install-ReqModules_AdminCheck.ps1) | TBD | See script README | TBD | TBD |
| [Listing_mailbox_content](Listing_mailbox_content/Listing_mailbox_content.ps1) | TBD | See script README | TBD | TBD |
| [Mailbox Audt Logs](Mailbox Audt Logs/Mailbox Audt Logs.ps1) | TBD | See script README | TBD | TBD |
| [Manage Exchange Online Mobile Devices for a User](Manage Exchange Online Mobile Devices for a User/Manage Exchange Online Mobile Devices for a User.ps1) | TBD | See script README | TBD | TBD |
| [manage-inactive-m365-users](manage-inactive-m365-users/manage-inactive-m365-users.ps1) | TBD | See script README | TBD | TBD |
| [manage-plus-addressing-general](manage-plus-addressing-general/manage-plus-addressing-general.ps1) | TBD | See script README | TBD | TBD |
| [map-sharepoint-drive-general](map-sharepoint-drive-general/map-sharepoint-drive-general.ps1) | TBD | See script README | TBD | TBD |
| [Meeting Room Creation Automation](Meeting Room Creation Automation/Meeting Room Creation Automation.ps1) | TBD | See script README | TBD | TBD |
| [prevent-bing-install-general](prevent-bing-install-general/prevent-bing-install-general.ps1) | TBD | See script README | TBD | TBD |
| [private-channel-management-general](private-channel-management-general/private-channel-management-general.ps1) | TBD | See script README | TBD | TBD |
| [remove-duplicate-group-licenses](remove-duplicate-group-licenses/remove-duplicate-group-licenses.ps1) | TBD | See script README | TBD | TBD |
| [remove-email-forwarding-management](remove-email-forwarding-management/remove-email-forwarding-management.ps1) | TBD | See script README | TBD | TBD |
| [reset-mfa-methods-general](reset-mfa-methods-general/reset-mfa-methods-general.ps1) | TBD | See script README | TBD | TBD |
| [reset-phone-mfa-methods](reset-phone-mfa-methods/reset-phone-mfa-methods.ps1) | TBD | See script README | TBD | TBD |
| [Add one|multiple members to RF](Room Finders/Add one|multiple members to RF.ps1) | TBD | See script README | TBD | TBD |
| [Listing Room Finders v2](Room Finders/Listing Room Finders v2.ps1) | TBD | See script README | TBD | TBD |
| [Listing Room Finders](Room Finders/Listing Room Finders.ps1) | TBD | See script README | TBD | TBD |
| [room-report-mailbox-report](room-report-mailbox-report/room-report-mailbox-report.ps1) | TBD | See script README | TBD | TBD |
| [send-app-credential-expiry-notifications](send-app-credential-expiry-notifications/send-app-credential-expiry-notifications.ps1) | TBD | See script README | TBD | TBD |
| [send-password-expiry-notifications](send-password-expiry-notifications/send-password-expiry-notifications.ps1) | TBD | See script README | TBD | TBD |
| [set-recurring-oof-replies](set-recurring-oof-replies/set-recurring-oof-replies.ps1) | TBD | See script README | TBD | TBD |
| [setup-ms-online-module](setup-ms-online-module/setup-ms-online-module.ps1) | TBD | See script README | TBD | TBD |
| [sharepoint-file-activity-audit](sharepoint-file-activity-audit/sharepoint-file-activity-audit.ps1) | TBD | See script README | TBD | TBD |
| [sharepoint-report-link-activity-report](sharepoint-report-link-activity-report/sharepoint-report-link-activity-report.ps1) | TBD | See script README | TBD | TBD |
| [Transport Rules Checker](Transport Rules Checker/Transport Rules Checker.ps1) | TBD | See script README | TBD | TBD |
