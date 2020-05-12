# ************************
# * Name - User information
# * Created - 12/05
# * Author - Adam
# * Version - 0.1
# ************************


#import csv and user input for result

Try{
$Path = Import-Csv C:\Temp\users.csv | Sort MailboxSizeGB -Descending
}
catch{
    clear
    "Please store User file into C:\temp\  and then run code again. File has not been stored correctly."
    pause
    break
}
clear
"Welcome to the Danger Zone.`nPlease select from the following options: `n1)Check user count`n2)Check Mailbox size`n3)Check for non identical email address vs UPN
4)Check NYC Mailbox size`n5)Check for Mailboxes larger than 10gb in NYC`n6)Check for top ten users`n7)Summary of sites`n8)Exit"
$input = Read-Host -Prompt 'Input'
#loop to pull it back to main menu until exit is selected
While ($input -ne 8){
#switch for menu selection
Switch ($input)
{
 #runs total user count
 1 {
 "Total user count is " + $path.count
 pause
 }
 
 #total mailboxsize
 2{
 $MailboxSize = $Path.mailboxsizegb | measure-object -sum | select sum
 "Total size " + $MailboxSize.Sum
 pause
 }
 #Non identical
 3 {
        $result = ""
        "The following addresses have a mismatch:"
         ForEach ($user in $path) {

            $email = $user.EmailAddress
            $UPN = $user.UserPrincipalName

            if ($email -ne $UPN){

            $result = $email + " " + $UPN
            $result
            }
           
  }
  pause
  }
  #NYC Mailbox size
  4{
        $SizeResult = 0
         ForEach ($user in $path) {

          $site = $user.Site
           

            if ($site -eq "NYC"){

            $NYCSize = $user.MailboxSizeGB | measure-object -sum | select sum
            $SizeResult += $NYCSize.sum

            }
            
  }
  $SizeResult
  pause
  }
  #larger than 10GB
  5 {
          $count = 0
          ForEach ($user in $path) {

          $MBSize = $user.MailboxSizeGB
          $account = $user.AccountType
          
            if ($MBSize -gt "10" -and $account -eq "Employee"){

            $count ++

            }

  }
  $count
  pause
  }
  #top ten users
  6{

           $know = ""
           ForEach ($user in $path) {

           $site = $user.Site
           

            if ($site -eq "NYC"){
                
             $EMadd = $user.EmailAddress
             $split = "@"
             $option = [System.StringSplitOptions]::RemoveEmptyEntries
             $a = $EMadd.split($split, $option)
             $know += $a[0],""
             
            }
}
    $know
    pause
}

#Run site check and export to .csv

7{
$result = $path| Group-Object -Property Site|
Select-Object @{Name='Site'   ;Expression={$_.Values[0]}},
              @{Name='TotalUserCount' ;Expression={($_.Group|Measure-Object UserPrincipalName -Sum ).Count}},
              @{Name='EmployeeCount'   ;Expression={($_.Group -match "Employee"|Measure-Object AccountType -Sum).Count}},
              @{Name='ContractorCount'   ;Expression={($_.Group -match "Contractor"|Measure-Object MailboxsizeGB -Sum).Count}},
              @{Name='TotalMailboxSizeGB'   ;Expression={($_.Group|Measure-Object MailboxsizeGB -Sum).Sum}},
              @{Name='AverageMailboxSizeGB'   ;Expression={[math]::Round(($_.Group|Measure-Object MailboxsizeGB -Average).average,1)}}
                        

$result | Export-Csv C:\temp\Results.csv -nti
"Results have been exported successfully. Set to C:\temp\Results"
Pause
  }
8 {
    Break
    } 
    default
    {
    "incorrect selection"
    pause
    }                     
 }
 clear
"Welcome to the Danger Zone`nPlease select from the following options: `n1)Check user count.`n2)Check Mailbox size`n3)Check for non identical email address vs UPN
4)Check NYC Mailbox size`n5)Check for Mailboxes larger than 10gb in NYC`n6)Check for top ten users`n7)Summary of sites`n8)Exit"
 $input = Read-Host -Prompt 'Input'
 }