# Microsoft 365 PowerShell Toolkit

Een verzameling PowerShell-scripts voor beheer, beveiliging en rapportage in Microsoft 365, Entra ID, Exchange Online en SharePoint Online. Ontwikkeld en ingezet in de praktijk bij Microsoft 365-omgevingen van uiteenlopende organisaties.

Gemaakt door **Maikel Moonen** – [Moonen IT Consulting](https://meek-mandazi-f8610f.netlify.app/)

## Uitgangspunten

- **Veilig standaardgedrag**: scripts die iets wijzigen draaien standaard als *dry-run* en voeren pas wijzigingen door met `-Execute`.
- **Geen vaste tenantgegevens**: tenant-ID's, app-registraties en accounts worden altijd als parameter meegegeven.
- **Rapportage voor beheer én management**: waar zinvol leveren scripts een opgemaakt Excel- of HTML-rapport op, naast een CSV-log.
- **Documentatie in het script**: elk script heeft een help-blok. Gebruik `Get-Help .\<script>.ps1 -Full`.

## Overzicht

### Conditional Access
| Script | Omschrijving |
|---|---|
| [Export-TenantPolicies.ps1](Conditional-Access/Export-TenantPolicies.ps1) | Exporteert alle Conditional Access-policies van een tenant naar één JSON-bestand, inclusief de namen van gerefereerde groepen, apps en locaties. |
| [Import-TenantPolicies.ps1](Conditional-Access/Import-TenantPolicies.ps1) | Importeert die export in een andere tenant: verwijzingen worden gemapt of opgeschoond, alle policies komen binnen als *Report-only*. |
| [Export-ConditionalAccessPolicies.ps1](Conditional-Access/Export-ConditionalAccessPolicies.ps1) | Leesbaar Excel-rapport van alle CA-policies (voorwaarden, uitsluitingen, grant controls). |
| [Check-MfaRegistrationPolicy.ps1](Conditional-Access/Check-MfaRegistrationPolicy.ps1) | Controleert of er een policy is die MFA-registratie afdwingt en signaleert policies in Report-only. |

### Entra ID
| Script | Omschrijving |
|---|---|
| [Get-EntraIdMfaSignInReport.ps1](Entra-ID/Get-EntraIdMfaSignInReport.ps1) | Analyseert sign-in logs op daadwerkelijk MFA-gebruik en levert een gap-analyse per gebruiker. |
| [Get-EntraGuestAccountReport.ps1](Entra-ID/Get-EntraGuestAccountReport.ps1) | Excel-dashboard van inactieve gastaccounts, als basis voor opschoning. |
| [Check-EntraIDClientSecrets.ps1](Entra-ID/Check-EntraIDClientSecrets.ps1) | Overzicht van verlopen en bijna verlopen client secrets van alle app-registraties. |
| [Copy-AdminAccount.ps1](Entra-ID/Copy-AdminAccount.ps1) | Kloont een (admin)account naar een nieuw cloud-only account, inclusief groepen en directory-rollen. |
| [Reset-EnterpriseAppSsoMode.ps1](Entra-ID/Reset-EnterpriseAppSsoMode.ps1) | Zet de SSO-modus van een Enterprise App terug en geeft een vastzittende SAML Entity ID vrij. |

### Exchange Online
| Script | Omschrijving |
|---|---|
| [Export-ExchangeAutoForwardingReport.ps1](Exchange-Online/Export-ExchangeAutoForwardingReport.ps1) | Rapport van alle actieve auto-forwarding (mailbox-forwarding én inbox rules), met nadruk op externe bestemmingen. |
| [Find-ExchangeOnlineObject.ps1](Exchange-Online/Find-ExchangeOnlineObject.ps1) | Zoekt een e-mailadres breed op: recipients, aliassen, Microsoft 365-groepen en soft-deleted mailboxen. |
| [Diagnose-0x8004010F.ps1](Exchange-Online/Diagnose-0x8004010F.ps1) | Diagnose van Outlook-fout 0x8004010F met een HTML-rapport en concrete adviezen. |
| [Delete-OldMailboxItems.ps1](Exchange-Online/Delete-OldMailboxItems.ps1) | Verwijdert met Compliance Search items ouder dan een bewaartermijn uit een mailbox. |

### SharePoint Online
| Script | Omschrijving |
|---|---|
| [Export-SPOSiteAttestatieRapport.ps1](SharePoint/Export-SPOSiteAttestatieRapport.ps1) | Attestatierapport (Excel) van alle sitecollecties met hun eigenaren, inactiviteit en risicoscore. Geschikt voor Azure Automation. |
| [Remove-SharePointOwnersTenantWide.ps1](SharePoint/Remove-SharePointOwnersTenantWide.ps1) | Verwijdert opgegeven accounts tenant-breed als eigenaar, zonder sites of groepen zonder eigenaar achter te laten. Parallel uitgevoerd. |
| [Repair-SharePointOwnerCleanupErrors.ps1](SharePoint/Repair-SharePointOwnerCleanupErrors.ps1) | Verwerkt de fouten uit een cleanup-run opnieuw, serieel en met retries. |

### Licenties
| Script | Omschrijving |
|---|---|
| [Check-MDILicenses.ps1](Licensing/Check-MDILicenses.ps1) | Inventariseert welke gebruikers een licentie hebben die Microsoft Defender for Identity omvat. |
| [Invoke-VisioPlan2UsageReport.ps1](Licensing/Invoke-VisioPlan2UsageReport.ps1) | Excel-dashboard van Visio Plan 2-gebruik per gebruiker, om ongebruikte licenties te vinden. |

### Partner / GDAP
| Script | Omschrijving |
|---|---|
| [Remove-GdapOrphanedAssignments.ps1](Partner-GDAP/Remove-GdapOrphanedAssignments.ps1) | Vindt en verwijdert access assignments in GDAP-relaties die naar niet-bestaande groepen verwijzen. |
| [Set-GdapContributorGroup.ps1](Partner-GDAP/Set-GdapContributorGroup.ps1) | Vervangt de gekoppelde securitygroep in alle GDAP-relaties in één keer. |

### Entra Connect
| Script | Omschrijving |
|---|---|
| [EntraConnectUpgrade-GUI.ps1](Entra-Connect/EntraConnectUpgrade-GUI.ps1) | Grafische (WinForms) begeleiding van een volledige Entra Connect-upgrade: ruim 30 stappen, met back-ups, controles en een dry-run modus. |

## Vereisten

- PowerShell 7.2 of hoger aanbevolen. Sommige scripts werken ook in Windows PowerShell 5.1 (zie `#Requires` bovenin het script).
- Afhankelijk van het script:
  ```powershell
  Install-Module Microsoft.Graph -Scope CurrentUser
  Install-Module ExchangeOnlineManagement -Scope CurrentUser
  Install-Module PnP.PowerShell -Scope CurrentUser
  Install-Module ImportExcel -Scope CurrentUser
  ```
- De benodigde Graph-scopes of app-permissies staan in het help-blok van elk script.

## Gebruik

```powershell
# Help en voorbeelden bekijken
Get-Help .\Conditional-Access\Import-TenantPolicies.ps1 -Full

# Eerst altijd een dry-run
.\SharePoint\Remove-SharePointOwnersTenantWide.ps1 @connectie -TargetAccount 'oud.account@contoso.com' -BackupOwnerAccount 'svc-backup@contoso.com'

# Daarna pas uitvoeren
.\SharePoint\Remove-SharePointOwnersTenantWide.ps1 @connectie -TargetAccount 'oud.account@contoso.com' -BackupOwnerAccount 'svc-backup@contoso.com' -Execute
```

> Test scripts altijd eerst in een test-tenant of met een dry-run. Gebruik is op eigen risico.

## Contact

Hulp nodig bij Microsoft 365, Entra ID of automatisering met PowerShell? Neem contact op via [Moonen IT Consulting](https://meek-mandazi-f8610f.netlify.app/).

## Licentie

[MIT](LICENSE)
