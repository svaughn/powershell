$study_abroad_tests_folder_name = "Study Abroad Tests"
$study_abroad_tests_rule_name = $study_abroad_tests_folder_name
$study_abroad_tests_token = '[sa-test]'

# record what directory we are in
$original_pwd = (Get-Location).Path

# go to the folder where the Interop dll is going to be under
cd  C:\Windows\assembly\

# find the dll
$interop_assemply_location = (Get-ChildItem -Recurse  Microsoft.Office.Interop.Outlook.dll).Directory

# switch to that directory so that any version of powershell will find it
cd $interop_assemply_location 

# load the assmbly for outlook
Add-Type -AssemblyName "Microsoft.Office.Interop.Outlook" 

# go back to the folder we were in at the start
cd "$original_pwd"

# connect to outlook
$Outlook = New-Object -comobject Outlook.Application

# find the users mailbox and then look for the study abroad folder
$namespace = $Outlook.GetNameSpace("MAPI")
$user_folder = $namespace.Folders | Where-Object  name -like "$($env:USERNAME)*" 
$special_folder = $user_folder.Folders | Where-Object  name -like "Special"
$study_abroad_tests_folder_name = "Study Abroad Tests"
$study_abroad_tests_folder = $special_folder.Folders | Where-Object  name -like $study_abroad_tests_folder_name

# the folder is not there, create it
if (!$study_abroad_tests_folder) {
    "creating $study_abroad_tests_folder_name email folder"
    $study_abroad_tests_folder = $special_folder.Folders.Add($study_abroad_tests_folder_name)
}

# look for the study abroad rule
$rules = $outlook.session.DefaultStore.GetRules()
$study_abroad_tests_rule = $rules | where-object name -like "$study_abroad_tests_rule_name"

# if the rule is not there, create it
if (!$study_abroad_tests_rule) {
    "creating $study_abroad_tests_folder_name email rule"
    $study_abroad_tests_rule = $rules.Create($study_abroad_tests_rule_name, `
                                  [Microsoft.Office.Interop.Outlook.OlRuleType]::olRuleReceive)

    $SubjectCondition = $study_abroad_tests_rule.Conditions.Subject
    $SubjectCondition.Enabled = $true
    $SubjectCondition.Text = @("$study_abroad_tests_token")
    $MoveTarget = $namespace.getFolderFromID($study_abroad_tests_folder.EntryID)
    $MoveRuleAction = $study_abroad_tests_rule.Actions.MoveToFolder
    $MoveRuleAction.Enabled = $true
    [Microsoft.Office.Interop.Outlook._MoveOrCopyRuleAction].InvokeMember(
        "Folder",
        [System.Reflection.BindingFlags]::SetProperty,
        $null,
        $MoveRuleAction,
        $MoveTarget)

    $rules.Save()    
}
