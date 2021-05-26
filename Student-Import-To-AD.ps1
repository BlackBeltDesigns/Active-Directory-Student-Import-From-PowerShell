<# 
    Set our Preference Variables and
    create the transcript log
#>
$ErrorActionPreference="SilentlyContinue"
Stop-Transcript | Out-Null
$ErrorActionPreference = "Continue"
Start-Transcript -path "C:\SftpRoot\loginid\logs\ou.log"

<#
    Adding old fallbacks just in case
#>
Import-Module ActiveDirectory -ErrorAction SilentlyContinue

<#
    VARIABLES

    Make all options easy to edit and maintain
#>
$CSVFilePath = "C:\SftpRoot\loginid\student-ad.csv"

$Schools = @(
    [pscustomobject]@{Name='ES';Id='105';Grades=@('1','2','3','4')}
    [pscustomobject]@{Name='NW';Id='120';Grades=@('1','2','3','4')}
    [pscustomobject]@{Name='SR';Id='115';Grades=@('1','2','3','4')}
    [pscustomobject]@{Name='EC';Id='130';Grades=@('PK','K')}
    [pscustomobject]@{Name='PI';Id='125';Grades=@('5','6')}
    [pscustomobject]@{Name='MS';Id='510';Grades=@('7','8')}
    [pscustomobject]@{Name='HS';Id='705';Grades=@('9','10','11','12')}
)

$Groups = '802.1x Staff','802.1x Student','Domain Admins','Domain Users' # Add Groups in that the students should be added to.
$SubDomain = 'student' # SubDomain is used for email only.

<#
    Name: SchoolName
    Params: $schoolID
    
    Instructions: Call function and pass a school ID 
    in order to get the readable name.
#>
Function SchoolName {
    [CmdletBinding()]
    param(
        [int] $passedID
    )
    switch($passedID) {
        130{return "EC"}
        105{return "ES"}
        115{return "SR"}
        120{return "NW"}
        125{return "PI"}
        510{return "MS"}
        705{return "HS"}
    }
} # END SchoolName

<#
    Name: GradeLevel
    Params: $passedID
    
    Instructions: Call function and pass a grade level 
    in order to get the readable name.
#>
Function GradeLevel {
    [CmdletBinding()]
    param(
        [string] $passedID
    )
    switch($passedID) {
        "0"{return "K"}
        "-1"{return "PK"}
        default{return "$passedID"}
    }
} # END GradeLevel

<#
    Name: CreateOUs
    Params: NA
    
    Instructions: Call function to verify and create 
    Organizational Units in AD.
#>
Function CreateOUs() {
    [CmdletBinding()]
    # Create Inner Variables
    $DomainPath = 'DC='+$env:USERDNSDOMAIN.Substring(0,$env:USERDNSDOMAIN.Length-4).ToLower()+',DC='+$env:USERDNSDOMAIN.Substring($env:USERDNSDOMAIN.Length-3).ToLower()

    if(![adsi]::Exists("LDAP://OU=Students,$DomainPath")){New-ADOrganizationalUnit Students -ProtectedFromAccidentalDeletion $True; Write-Host "Students OU being created." -foregroundcolor Green}else{Write-Host "Students OU already exists. Skipping..." -foregroundcolor Yellow}

    # Check for the schools OU paths
    $Schools | ForEach-Object {
        $SchoolPath = "OU=Students,$DomainPath"
        $SchoolName = $_.'Name'

        if(![adsi]::Exists("LDAP://OU=$SchoolName,$SchoolPath")){New-ADOrganizationalUnit $SchoolName -Path $SchoolPath -ProtectedFromAccidentalDeletion $False; Write-Host "$SchoolName OU being created." -foregroundcolor Green}else{Write-Host "$SchoolName OU already exists. Skipping..." -foregroundcolor Yellow}

        $GradePath = "OU=$SchoolName,$SchoolPath"
        
        ForEach($Grade in $_.'Grades'){
            if(![adsi]::Exists("LDAP://OU=$Grade,$GradePath")){New-ADOrganizationalUnit $Grade -Path $GradePath -ProtectedFromAccidentalDeletion $False; Write-Host "$Grade OU being created." -foregroundcolor Green}else{Write-Host "$Grade OU already exists. Skipping..." -foregroundcolor Yellow}
        }
    }
    
} CreateOUs # END CreateOUs

<#
    Name: CreateGroups
    Params: NA
    
    Instructions: Call function to verify and create 
    Organizational Groups in AD.
#>
Function CreateGroups {
    [CmdletBinding()]
    # Check for and create the Groups in the provided array above
    $Groups | ForEach-Object {
        # Create Path variable
        $gPath = 'CN=Users,DC='+$env:USERDNSDOMAIN.Substring(0,$env:USERDNSDOMAIN.Length-4).ToLower()+',DC='+$env:USERDNSDOMAIN.Substring($env:USERDNSDOMAIN.Length-3).ToLower()
        $gName = $_

        if(![adsi]::Exists("LDAP://CN=$gName,$gPath")){New-ADGroup -GroupCategory "Security" -GroupScope "Global" -Name "$gName" -Path "$gPath" -SamAccountName "$gName"; Write-Host "$gName Group is being created." -foregroundcolor Green}else{Write-Host "$gName Group already exists. Skipping..." -foregroundcolor Yellow}
    }
} CreateGroups # END CreateGroups

<#
    Name: CreateUser
    Params: [object]User

    Instructions: Call function and pass in a User object 
    which has all values needed to create New-ADUser.
#>
Function CreateUser {
    [CmdletBinding()]
    param(
        [PSObject[]] $PassedUser
    )
    try {
        New-ADUser @PassedUser
    } catch {
        Write-Error "Error creating user: $($PassedUser.lname), $($PassedUser.fname)"
    }
} # END CreateUser

<#
    Name: UpdateUser
    Params: [object]User

    Instructions: Call function and pass in a User object 
    which has all values needed to update Set-ADUser.
#>
Function UpdateUser {
    [CmdletBinding()]
    param(
        [array]$PassedUser
    )
    try {
        Get-ADUser -Identity "$($PassedUser.studentID)" | Set-ADUser $PassedUser
    } catch {
        Write-Error "Error updating user: $($PassedUser.lname), $($PassedUser.fname)"
    }
} # END UpdateUser

<#
    Name: ShowProgress
    Params: PSObject

    Instructions: Call function and pass object
    to display a progress indicator for the process.
#>
Function ShowProgress {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true, Position=0, ValueFromPipeline=$true)]
        [PSObject[]]$InputObject,
        [string]$Activity = "Processing items"
    )

        [int]$TotItems = $Input.Count
        [int]$Count = 0

        $Input|foreach {
            $_
            $Count++
            [int]$percentComplete = ($Count/$TotItems* 100)
            Write-Progress -Activity $Activity -PercentComplete $percentComplete -Status ("Working - " + $percentComplete + "%")
        }
} # END ShowProgress

<#
    Begin the real work here. 
    Above is variable declarations and 
    building operations for the AD-Groups
    and the AD-OUs.
#>
#Store the data from Users.csv in the $Users variable
Write-Host 'Importing the data. Can be long, please wait.'
If ([System.IO.File]::Exists($CSVFilePath)) {
    
    $Users = Import-csv $CSVFilePath `
        -Header lname, fname, mname, studentID, gradeLevel, schoolID

        $index = 0
        $total = @($Users).Count
        $starttime = $lasttime = Get-Date

        #$Users |
        #Select *, @{Name='schoolName';Expression={SchoolName $_.schoolID}} | `
        #Select *, @{Name='actualGrade';Expression={GradeLevel $_.gradeLevel}} | `
        #Sort-Object -Property "schoolID", "gradeLevel", "lname", "fname", "mname"
                
        foreach($user in $Users) {

            $index++
            $currtime = (Get-Date) - $starttime
            $avg = $currtime.TotalSeconds / $index
            $last = ((Get-Date) - $lasttime).TotalSeconds
            $left = $total - $index
            $WrPrgParam = @{
                Activity = (
                    "Processing User $($user.lname), $($user.fname)",
                    "Total: $($currtime -replace '\..*')",
                    "Avg: $('{0:N2}' -f $avg)",
                    "Last: $('{0:N2}' -f $last)",
                    "ETA: $('{0:N2}' -f ($avg * $left / 60))",
                    "min ($([string](Get-Date).AddSeconds($avg*$left) -replace '^.* '))"
                ) -join ' '
                Status = "$index of $total ($left left) [$('{0:N2}' -f ($index / $total * 100))%]"
                CurrentOperation = "$user.lname, $user.fname being processed..."
                PercentComplete = $index / $total * 100
            }
            Write-Progress @WrPrgParam
            $lasttime = Get-Date

            <# 
                Check for inclusion of a Middle Name.
                Build appropriate variables needed.
            #>
            $HasMiddleName = ![string]::IsNullOrEmpty($user.mname)

            #ShowProgress -Activity "Middle Name check and assignment:" `

            If ($HasMiddleName) {
                $global:Password = ($user.fname.substring(0,1) + $user.mname.substring(0,1) + $user.lname.substring(0,1) + $user.studentID)
                $global:MiddleInitial = $user.mname.substring(0,1)
            }else{
                $global:Password = ($user.fname.substring(0,1) + $user.lname.substring(0,1) + $user.studentID)
                $global:MiddleInitial = $Null
            }

            <#
                Check for inclusion of a SubDomain.
                Build out appropriate variables needed.
            #>
            
            #ShowProgress -Activity "Sub-Domain check progress:" `
            
            If (![string]::IsNullOrEmpty($SubDomain)){
                $global:StudentEmail = ($user.studentID+"@"+$SubDomain+"."+$env:USERDNSDOMAIN.Substring(0,$env:USERDNSDOMAIN.Length-4).ToLower()+"."+$env:USERDNSDOMAIN.Substring($env:USERDNSDOMAIN.Length-3).ToLower())
            } else {
                $global:StudentEmail = ($user.studentID+"@"+$env:USERDNSDOMAIN.Substring(0,$env:USERDNSDOMAIN.Length-4).ToLower()+"."+$env:USERDNSDOMAIN.Substring($env:USERDNSDOMAIN.Length-3).ToLower())
            }

            <#
                Check for existing user.
                Either update an existing or create a new
            #>
            try {
                Get-ADUser -Identity $user.studentID
                # Account exists. Update name fields.
                $StudentParameters = @{
                    ChangePasswordAtLogon = $False
                    Description = "$($user.studentID)"
                    DisplayName = "$($user.lname+", "+$user.fname)"
                    Mail = "$($StudentEmail)"
                    Enabled = $True
                    GivenName = "$($user.fname)"
                    PasswordNeverExpires = $True
                    SamAccountName = "$($user.studentID)"
                    Surname = "$($user.lname)"
                    UserPrincipalName = "$($user.studentID+"@"+$env:USERDNSDOMAIN.ToLower())"
                }

                # Set trouble properties this way instead.
                if( ($Null -ne $MiddleInitial) -and ("" -ne $MiddleInitial) ) {
                    $StudentParameters.Add("Initials", "$MiddleInitial")
                }
                
                $StudentParameterString = $StudentParameters | Out-String
                Write-Host $StudentParameterString
                UpdateUser $StudentParameters
            } catch [Microsoft.ActiveDirectory.Management.ADIdentityResolutionException] {
                # Account does NOT exist. Create new user.
                $StudentParameters = @{
                    ChangePasswordAtLogon = $False
                    Description = $($user.studentID)
                    DisplayName = ($user.lname+", "+$user.fname)
                    EmailAddress = $($StudentEmail)
                    Enabled = $True
                    GivenName = $($user.fname)
                    #Initials = $($MiddleInitial)
                    Name = $($user.studentID)
                    PasswordNeverExpires = $True
                    Path = "OU=$user.actualGrade,OU=$user.schoolName,OU=Students,DC="+$env:USERDNSDOMAIN.Substring(0,$env:USERDNSDOMAIN.Length-4).ToLower()+",DC="+$env:USERDNSDOMAIN.Substring($env:USERDNSDOMAIN.Length-3).ToLower()
                    SamAccountName = S($user.studentID)
                    Surname = $($user.lname)
                    UserPrincipalName = ($user.studentID+"@"+$env:USERDNSDOMAIN.ToLower())
                }
                CreateUser $StudentParameters
            }
    }
}

<#
    Let's wrap it up now!!!

    Give the date (should be the same)
    and stop the transcript log. 
#>
# Log the date
Get-Date

# Stop the transcript log
Stop-Transcript
