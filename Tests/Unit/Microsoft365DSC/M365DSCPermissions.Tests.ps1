BeforeAll {
    Import-Module "$PSScriptRoot/../../../Modules/Microsoft365DSC/Modules/M365DSCDllLoader.psm1" -Force -Global
    Initialize-M365DSCDllLoader
    Import-Module "$PSScriptRoot/../../../Modules/Microsoft365DSC/Microsoft365DSC.psd1" -Global -WarningAction SilentlyContinue
}

Describe 'Update-M365DSCAzureAdApplication' {
    BeforeAll {
        $Script:GraphAppId = '00000003-0000-0000-c000-000000000000'
        $Script:PowerAppsAppId = '475226c6-020e-4fb2-8a90-7a972cbfc1d4'
    }

    BeforeEach {
        Mock -ModuleName M365DSCPermissions -CommandName Confirm-M365DSCDependencies
        Mock -ModuleName M365DSCPermissions -CommandName New-M365DSCConnection
        Mock -ModuleName M365DSCPermissions -CommandName Write-Host
        Mock -ModuleName M365DSCPermissions -CommandName Update-MgApplication
        Mock -ModuleName M365DSCPermissions -CommandName Get-MgApplication -MockWith {
            return @{
                Id                  = '11111111-1111-1111-1111-111111111111'
                AppId               = '22222222-2222-2222-2222-222222222222'
                AppRoles            = @()
                PasswordCredentials = @(@{ DisplayName = 'Existing secret' })
            }
        }
        Mock -ModuleName M365DSCPermissions -CommandName Get-MgServicePrincipal -MockWith {
            return @{
                AppId    = $Script:GraphAppId
                AppRoles = @(@{ Id = '5b567255-7703-4780-807c-7be8301ae99b'; Value = 'Group.Read.All' })
            }
        }
        Mock -ModuleName M365DSCPermissions -CommandName Get-MgServicePrincipal -ParameterFilter { $Filter -eq "appId eq '$Script:PowerAppsAppId'" } -MockWith {
            return @{
                AppId    = $Script:PowerAppsAppId
                AppRoles = @()
            }
        }
    }

    Context 'When an API offers none of the requested permissions as application permission' {
        BeforeAll {
            $Script:Permissions = @(
                @{ Api = 'Graph'; PermissionName = 'Group.Read.All' }
                @{ Api = 'powerAppsService'; PermissionName = 'User' }
            )
        }

        It 'Requests access to the other APIs only' {
            Update-M365DSCAzureAdApplication -ApplicationName 'Microsoft365DSC' -Permissions $Script:Permissions -Type Secret

            Should -Invoke -ModuleName M365DSCPermissions -CommandName Update-MgApplication -Exactly -Times 1 -ParameterFilter {
                $null -ne $BodyParameter.requiredResourceAccess -and
                @($BodyParameter.requiredResourceAccess).Count -eq 1 -and
                $BodyParameter.requiredResourceAccess[0].resourceAppId -eq $Script:GraphAppId -and
                @($BodyParameter.requiredResourceAccess[0].resourceAccess).Count -eq 1
            }
        }

        It 'Warns about the permission it skips' {
            Update-M365DSCAzureAdApplication -ApplicationName 'Microsoft365DSC' -Permissions $Script:Permissions -Type Secret

            Should -Invoke -ModuleName M365DSCPermissions -CommandName Write-Host -Exactly -Times 1 -ParameterFilter {
                $Object -like "*API 'powerAppsService' has no application permission 'User'*"
            }
        }
    }
}
