<# MIT License

Copyright (c) Andy Blackman.

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE
#>

<#
.SYNOPSIS
Saves emails from a users account to the filesystem

.DESCRIPTION
Saves all emails from a named folder and its subfolders in a specified users account to the filesystem in the same folder structure

.PARAMETER userPrincipalName
The userPrincipalName of the mailbox to extract from


.PARAMETER emailFolderName
The name of the folder in the specified mailbox to extract from

.PARAMETER fileFolderRoot
The folder to save the emails under

.PARAMETER replaceInvalidCharacterWith
Replace any characters in the subject name with the character specified.  Defaults to "-"

.PARAMETER duplicateFileSeparator
Uses this character when multiple emails have the same datetime and subject name.  This will be used with an incremental number as a suffix to the filename

.EXAMPLE
.\Export-emailsToFiles.ps1 -userPrincipalName joe.bloggs@example.domain.com -emailFolderName "Inbox" -fileFolderRoot C:\EmailExtract
Exports all emails in the Inbox folder and any subfolders to the folder "C:\EmailExtract"

.NOTES
You must have Mail.Read.Shared permissions on MS Graph and delegated access on the mailbox you are going to extract from to successfully run this script
#>

[CmdletBinding()]
param (
    [Parameter(Position=0, Mandatory=$true, HelpMessage="userPrincipalName of the mailbox")]
    [string]
    $userPrincipalName,
    [Parameter(Position=1, Mandatory=$true, HelpMessage="Name of the email folder to extract from")]
    [string]
    $emailfolderName,
    [Parameter(Position=2, Mandatory=$true, HelpMessage="Filesystem folder to extract the emails to")]
    [string]
    $fileFolderRoot,
    [Parameter(Position=3, Mandatory=$false, HelpMessage="Character to replace invalid file system characters with when saving emails")]
    [string]
    $replaceInvalidCharacterWith="-",
    [Parameter(Position=4, Mandatory=$false, HelpMessage="Character to use as a prefix to a count when saving emails with duplicate names to the same folder")]
    [string]
    $duplicateFileSeparator="_"

)

[string] $illegalFileCharacters='[^\x20-\x7F]|~|"|#|%|\&|\*|:|<|>|\?|\/|\\|{|\||}'
function find-mailFolder {
    [CmdletBinding()]
    param (
        [string] $folderId,
        [string] $userId,
        [string] $folderName
    )
        
        if ($folderId) {
            $result=Get-MgUserMailFolderChildFolder -MailFolderId $folderid -UserId $userId -All:$true
        } else {
            $result=Get-MgUserMailFolder -UserId $userId -All:$true
        }
        $found=$result|where {$_.DisplayName -eq $folderName}
        if (-not $found) {
            foreach($folder in $result) {
                Write-Verbose "Searching Folder: $($folder.DisplayName)"
                $found=find-mailFolder -folderId $folder.id -UserId $userId -folderName $folderName
                if ($found) {
                    return $found
                }
            }  
        } else {
            return $found
        }
    
}

function invoke-saveEmails {
    param (
        [string] $folderId,
        [string] $userId,
        [string] $fileFolder
    ) 

    $emails=Get-MgUserMailFolderMessage -MailFolderId $folderid -UserId $userId -all:$true
    $folders=Get-MgUserMailFolderChildFolder -MailFolderId $folderid -UserId $userId -All:$true

    foreach ($email in $emails) {
        $filename=(($email.SentDateTime.Tostring("yyyyMMdd_HHmmss") + " - " + $email.subject) -replace $illegalFileCharacters, $replaceInvalidCharacterWith).trim()
        [int]$count=0
        Write-Verbose "$($fileFolder)\$($filename)"
        while (Test-Path -literalpath "$($fileFolder)\$($filename).eml") {
            $filename=($filename -replace "$($duplicateFileSeparator)$($count)$", "").trim()
            $count+=1
            $filename+="$($duplicateFileSeparator)$($count)"
        }
        if ($count -gt 0) {
            Write-verbose "Altered Filename to $($filename)"
        }
        $filename+=".eml"
        Get-MgUserMessageContent -UserId $userId -MessageId $email.Id -OutFile "$($fileFolder)\$($filename)" 
    }
    foreach ($folder in $folders) {
        Write-Verbose "Searching Folder: $($result.DisplayName)"
        $newFolderName=$filefolder+"\"+ ($folder.DisplayName -replace $illegalFileCharacters, $replaceInvalidCharacterWith).trim()
        invoke-saveEmails -folderId $folder.id -userId $userId -fileFolder $newFolderName
    }

}

Connect-MgGraph -Scopes "Mail.Read.Shared"
$startingFolder=find-mailFolder -userId $userPrincipalName -folderName $emailfolderName
invoke-saveEmails -folderId $startingFolder.id -userId $userPrincipalName -fileFolder $fileFolderRoot
