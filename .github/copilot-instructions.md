# PowerShell Script Guidelines

## Code Style

### Parameter Declarations
- Use `[Parameter()]` attribute with descriptive `HelpMessage`
- Use validation attributes: `[ValidateNotNullOrEmpty()]`, `[ValidateRange()]`
- Declare parameter types explicitly: `[string]`, `[int]`, `[switch]`, etc.
- Set `Mandatory=$True` or `Mandatory=$False` explicitly

## Architecture

### Template System
Files under the `Graph` folder use a template-based code generation system for shared functionality:

**Shared Code Modules** (in `Graph/Build/`):
- `Auth.ps1` - Authentication (app-only and delegated flows)
- `Logging.ps1` - Logging functionality
- `Graph.ps1` - Microsoft Graph API helpers
- `Mailbox.ps1` - Mailbox-specific operations

**Template Markers**:
Scripts use special comment markers to inject shared code:
- `<# ModuleName.ps1 %PARAMS_START% #>` / `<# ModuleName.ps1 %PARAMS_END% #>`
- `<# ModuleName.ps1 %FUNCTIONS_START% #>` / `<# ModuleName.ps1 %FUNCTIONS_END% #>`

**DO NOT** manually edit code between these markers - use `Apply-TemplateCode.ps1` instead.

### Microsoft Graph Scripts
Graph API scripts follow these patterns:
- Use `AppId`, `AppSecretKey`, `TenantId` parameters for authentication
- Support both application (client credential) and delegated permissions
- Default `$TenantId` to "common" for multi-tenant scenarios
- Default `$User` to "me" for delegated permissions
- Include `ConnectionTimeout` parameter (default 30 seconds)

## Build and Test

### Applying Templates to New Scripts
```powershell
$ScriptFolder = "C:\Github Repos\David-Barrett-MS\PowerShell\Graph"
$SharedFolder = "$ScriptFolder\Build"
$NewScript = "$ScriptFolder\New-Script.ps1"

# For basic Graph API script
.\Apply-TemplateCode.ps1 -TargetScriptOrFolder $NewScript -SharedCodeFolder $SharedFolder -SharedCode @("Auth.ps1", "Logging.ps1", "Graph.ps1") -Template "Template-Default.ps1"

# For mailbox-related script
.\Apply-TemplateCode.ps1 -TargetScriptOrFolder $NewScript -SharedCodeFolder $SharedFolder -SharedCode @("Auth.ps1", "Logging.ps1", "Graph.ps1", "Mailbox.ps1") -Template "Template-Mailbox.ps1"
```

### Updating Existing Scripts
```powershell
# Update single script
.\Apply-TemplateCode.ps1 -TargetScriptOrFolder "$ScriptFolder\Test-GraphSearch.ps1" -SharedCodeFolder $SharedFolder -SharedCode @("Auth.ps1", "Logging.ps1", "Graph.ps1") -BackupFolder $BackupFolder

# Update all scripts in folder
.\Apply-TemplateCode.ps1 -TargetScriptOrFolder $ScriptFolder -SharedCodeFolder $SharedFolder -SharedCode @("Auth.ps1", "Logging.ps1", "Graph.ps1") -BackupFolder $BackupFolder
```

## Conventions

### Script Organization
- **Graph/** - Microsoft Graph API scripts and shared build modules
- **Miscellaneous/** - General utility scripts
- **Office 365 Administration/** - O365 admin operations
- **Outlook/** - Outlook-specific operations
- **SMTP/** - Email sending with SMTP OAuth

### Naming
- Use Verb-Noun naming convention for script files
- Test scripts prefix with `Test-`
- Use descriptive names that indicate purpose

### Error Handling
Follow the patterns established in the shared code modules for consistent error handling across scripts.

### Authentication
- Application permissions (app-only): Require `$AppSecretKey` parameter
- Delegated permissions: OAuth flow with `$RedirectUrl`
- Document required permissions in `.PREREQUISITES` section
