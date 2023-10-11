# extract-emailsToFiles
DESCRIPTION
    Saves all emails from a named folder and its subfolders in a specified users account to the filesystem in the same folder structure

SYNTAX
    .\export-emailsToFiles.ps1 [-userPrincipalName] <String> [-emailfolderName] <String> [-fileFolderRoot] <String>
    [[-replaceInvalidCharacterWith] <String>] [[-duplicateFileSeparator] <String>] [<CommonParameters>]

PARAMETERS
    -userPrincipalName <String>
        The userPrincipalName of the mailbox to extract from

    -emailfolderName <String>
        The name of the folder in the specified mailbox to extract from

    -fileFolderRoot <String>
        The folder to save the emails under

    -replaceInvalidCharacterWith <String>
        Replace any characters in the subject name with the character specified.  Defaults to "-"

    -duplicateFileSeparator <String>
        Uses this character when multiple emails have the same datetime and subject name.  This will be used with an incremental number as a suffix to the
        filename

    -------------------------- EXAMPLE 1 --------------------------

    PS C:\>.\Export-emailsToFiles.ps1 -userPrincipalName joe.bloggs@example.domain.com -emailFolderName "Inbox" -fileFolderRoot C:\EmailExtract

    Exports all emails in the Inbox folder and any subfolders to the folder "C:\EmailExtract"
