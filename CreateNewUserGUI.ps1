 #-------------------------------------------------------------#
#         CODE STARTED FOR CREATE NEW USER GUI BELOW          #
# THIS IS IN ALPHA AND MAY NOT FUNCTION AS CURRENTLY INTENDED #
#-------------------------------------------------------------#

　
function CallNewUserGUI {
$MyDocumentsBaloo = [environment]::getfolderpath("mydocuments") + "\WindowsPowerShell\BalooTrooper.png"

#XML code for GUI objects
$inputXML = @"
<Window x:Class="Bear.MainWindow"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:Bear"
        mc:Ignorable="d"
        Title="Bear Necessities | v2.3 | Create New Users" Height="510" Width="750" BorderBrush="#FF211414" Background="#FF6C6B6B" ResizeMode="CanMinimize" WindowStartupLocation="CenterScreen">

    <Grid>

        <Image Height="256" HorizontalAlignment="Left" Name="image1" Stretch="Fill" VerticalAlignment="Top" Width="192" SnapsToDevicePixels="False" Source="$MyDocumentsBaloo" />

        <TextBox Name="FirstName" Text="" Background="Black" CharacterCasing="Upper" Cursor="IBeam" Foreground="White" Height="27" HorizontalAlignment="Left" Margin="12,284,0,0" VerticalAlignment="Top" Width="211" />
        <Label Content="(First Name)" Height="28" HorizontalAlignment="Left" Margin="12,262,0,0" VerticalAlignment="Top" FontWeight="Bold" Width="106" />

        <TextBox Name="MiddleIn" Text="" Background="Black" CharacterCasing="Upper" Cursor="IBeam" Foreground="White" Height="27" HorizontalAlignment="Left" Margin="258,284,0,0" VerticalAlignment="Top" Width="211" />
        <Label Content="(Middle Initial)" FontWeight="Bold" Height="28" HorizontalAlignment="Left" Margin="258,262,0,0" VerticalAlignment="Top" Width="97" />

        <TextBox Name="LastName" Text="" Background="Black" CharacterCasing="Upper" Cursor="IBeam" Foreground="White" Height="27" HorizontalAlignment="Left" Margin="505,284,0,0" VerticalAlignment="Top" Width="211" />
        <Label Content="(Last Name)" FontWeight="Bold" Height="28" HorizontalAlignment="Left" Margin="505,262,0,0" VerticalAlignment="Top" Width="97" />

        <TextBox Name="Organization" Text="" Background="Black" CharacterCasing="Upper" Cursor="IBeam" Foreground="White" Height="27" HorizontalAlignment="Left" Margin="12,332,0,0" VerticalAlignment="Top" Width="211" />
        <Label Content="(Organization)" FontWeight="Bold" Height="28" HorizontalAlignment="Left" Margin="12,310,0,0" VerticalAlignment="Top" Width="97" />

        <TextBox Name="Department" Text="" Background="Black" CharacterCasing="Upper" Cursor="IBeam" Foreground="White" Height="27" HorizontalAlignment="Left" Margin="258,332,0,0" VerticalAlignment="Top" Width="211" />
        <Label Content="(Department)" FontWeight="Bold" Height="28" HorizontalAlignment="Left" Margin="258,310,0,0" VerticalAlignment="Top" Width="97" />

        <TextBox Name="Phone" Text="" Background="Black" CharacterCasing="Upper" Cursor="IBeam" Foreground="White" Height="27" HorizontalAlignment="Left" Margin="505,332,0,0" VerticalAlignment="Top" Width="211" />
        <Label Content="(Phone)" FontWeight="Bold" Height="28" HorizontalAlignment="Left" Margin="505,310,0,0" VerticalAlignment="Top" Width="97" />

        <TextBox Name="Email" Text="" Background="Black" CharacterCasing="Upper" Cursor="IBeam" Foreground="White" Height="27" HorizontalAlignment="Left" Margin="12,380,0,0" VerticalAlignment="Top" Width="211" />
        <Label Content="(Official User Email)" FontWeight="Bold" Height="28" HorizontalAlignment="Left" Margin="12,358,0,0" VerticalAlignment="Top" Width="127" />

        <TextBox Name="JobTitle" Text="" Background="Black" CharacterCasing="Upper" Cursor="IBeam" Foreground="White" Height="27" HorizontalAlignment="Left" Margin="258,380,0,0" VerticalAlignment="Top" Width="211" />
        <Label Content="(Job Title)" FontWeight="Bold" Height="28" HorizontalAlignment="Left" Margin="258,358,0,0" VerticalAlignment="Top" Width="97" />

        <TextBox Name="Description" Text="Domain User" Background="Black" CharacterCasing="Upper" Cursor="IBeam" Foreground="White" Height="27" HorizontalAlignment="Left" Margin="505,380,0,0" VerticalAlignment="Top" Width="211" />
        <Label Content="(Description)" FontWeight="Bold" Height="28" HorizontalAlignment="Left" Margin="505,358,0,0" VerticalAlignment="Top" Width="97" />

        <TextBox Name="Designation" Text="" Background="Black" CharacterCasing="Upper" Cursor="IBeam" Foreground="White" Height="27" HorizontalAlignment="Left" Margin="12,428,0,0" VerticalAlignment="Top" Width="211" />
        <Label Content="(Designation)" FontWeight="Bold" Height="28" HorizontalAlignment="Left" Margin="12,406,0,0" VerticalAlignment="Top" Width="97" />

        <TextBox Name="SupervisorEmail" Text="" Background="Black" CharacterCasing="Upper" Cursor="IBeam" Foreground="White" Height="27" HorizontalAlignment="Left" Margin="258,428,0,0" VerticalAlignment="Top" Width="211" />
        <Label Content="(Supervisor's Email)" FontWeight="Bold" Height="28" HorizontalAlignment="Left" Margin="258,406,0,0" VerticalAlignment="Top" Width="128" />

	<Button Name="NewUser" Background="Black" BorderBrush="Black" BorderThickness="2" Content="Create New User" Foreground="White" Height="30" HorizontalAlignment="Left" Margin="559,14,0,0" VerticalAlignment="Top" Width="144" FontSize="13" FontWeight="Bold" FontFamily="Arial" />

	<Label Content="(User Template)" FontWeight="Bold" Height="28" HorizontalAlignment="Left" Margin="198,65,0,0" VerticalAlignment="Top" Width="106" />
 	<ListBox Name="Templates" AllowDrop="True" Background="Black" BorderBrush="Black" BorderThickness="2" Foreground="White" Height="167" HorizontalAlignment="Left" ItemsSource="{Binding}" Margin="198,0,0,215" VerticalAlignment="Bottom" Width="211">

            <ListBoxItem Name="Template1" Content="Template1" />
            <ListBoxItem Name="Template2" Content="Template2" />
            <ListBoxItem Name="Template3" Content="Template3" />
            <ListBoxItem Name="Template4" Content="Template4" />
            <ListBoxItem Name="Template5" Content="Template5" />
            <ListBoxItem Name="Template6" Content="Template6" />
            <ListBoxItem Name="Template7" Content="Template7" />
            <ListBoxItem Name="Template8" Content="Template8" />
            </ListBox>

    </Grid>
</Window>               
 
"@       
 
$inputXML = $inputXML -replace 'mc:Ignorable="d"','' -replace "x:N",'N'  -replace '^<Win.*', '<Window'
 
[void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework')
[xml]$XAML = $inputXML
#Read XAML
 
    $reader=(New-Object System.Xml.XmlNodeReader $xaml)
    try{
    $Form=[Windows.Markup.XamlReader]::Load( $reader )}

catch{
Write-Host "Unable to load Windows.Markup.XamlReader. Double-check syntax and ensure .net is installed."}
 
#===========================================================================
# Store Form Objects In PowerShell
#===========================================================================
 
$xaml.SelectNodes("//*[@Name]") | %{Set-Variable -Name "WPF$($_.Name)" -Value $Form.FindName($_.Name)}

#Connect to Controls
$firstnameTextBox = $Form.FindName('Firstname')
$middleinTextBox = $Form.FindName('MiddleIn')
$lastnameTextBox = $Form.FindName('LastName')
$organizationTextBox = $Form.FindName('Organization')
$emailTextBox = $Form.FindName('Email')
$designationTextBox = $Form.FindName('Designation')
$departmentTextBox = $Form.FindName('Department')
$jobtitleTextBox = $Form.FindName('JobTitle')
$phoneTextBox = $Form.FindName('Phone')
$descriptionTextBox = $Form.FindName('Description')
$supervisoremailTextBox = $Form.FindName('SupervisorEmail')
$templatesListBox = $Form.FindName('Templates')
$newuserButton = $Form.FindName('NewUser')

#===========================================================================
# Actually make the objects work
#===========================================================================

#Create New User Button 
$newuserButton.Add_Click({

function CreateNewUser { 

<#
.SYNOPSIS 
Written by:
Greg & JBear 11/2/2016

Last Edited: 
JBear 12/24/2016 - Edited to interact with GUI.

Requires: ActiveDirectory Module
            & PowerShell Version 3 or higher

Creates a new active directory user from a template.

Purpose of script to assist Help Desk with the creation of End-User accounts in Active Directory.
#>

　
#Script requires ActiveDirectory Module to be loaded
Import-Module ActiveDirectory
    

    #User account information variables
    $Designation = $(

        If($designationTextBox.Text -EQ "Canada") {
            "Civilian (CND)"
        }
                       
        ElseIf($designationTextBox.Text -EQ "USA") {
            "Civilian (USA)"
        }

        ElseIf($designationTextBox.Text-EQ "Military") {
            "Foreign Military"
        }

        ElseIf($designationTextBox.Text -EQ "Contractor") {
            "Contractor"
        }
        Else {
            "Other"
        })
                                                                     
    $Displayname = $(
     
        If([string]::IsNullOrWhiteSpace($middleinTextBox.Text)) {
            $lastnameTextBox.Text + ", " + $firstnameTextBox.Text + " $Designation"
        }
        
        Else {
            $lastnameTextBox.Text + ", " + $firstnameTextBox.Text + " " + $middleinTextBox.Text + " $Designation"
        })
 
    $UserFirstname = $firstnameTextBox.Text
    $UserInitial = $middleinTextBox.Text
    $UserLastname = $lastnameTextBox.Text 
    $SupervisorEmail = $supervisoremailTextBox.Text
    $UserCompany = $organizationTextBox.Text
    $UserDepartment =  $departmentTextBox.Text
    $UserJobTitle = $jobtitleTextBox.Text
    $OfficePhone = $phoneTextBox.Text
    $Description = $descriptionTextBox.Text

    $Email = $emailTextBox.Text
    $Info = $(

	    $Date = Get-Date
	    "Account Created: " + $Date.ToShortDateString() + " " + $Date.ToShortTimeString() + " - " +  [Environment]::UserName
    )

    $FindSuperV = Get-ADUser -Filter {(mail -like $User.SupervisorEmail)}
    $FindSuperV = $FindSuperV | select -First "1" -ExpandProperty SamAccountName

    $Password = '7890&*()uiopUIOP'
    $Template = $templatesListBox.Name

　
　
#Load Visual Basic .NET Framework
[Void][System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')

　
    #Do{ process } Until( )
    Do{ 

        #Continue if $True
        While($True) {
            $SAM = [Microsoft.VisualBasic.Interaction]::InputBox("Enter desired Username for $Displayname :", "Create Username", "") 
            
            #Will loop if no value is supplied for $SAM
            If($SAM -ne "$Null"){

                #If AD user exists, throw error warning; loop back to $SAM input
                Try {

                    #On error, jump to Catch { }
                    $FindSAM = Get-ADUser $SAM -ErrorAction Stop
                    $SAMError = [Microsoft.VisualBasic.Interaction]::MsgBox("Username [$SAM] already in use by: " + $FindSAM.Name + "`nPlease try again...", "OKOnly,SystemModal", "Error")
                }#Try

                #On -EA Stop, specified account doesn't exist; continue with creation
                Catch {
                    $SAMFound = $False 
                    Break;   
                }#Catch
            }#If
        }#While
    }#Do

#Break from Do { } when $SAMFound is $False
Until($SAMFound -eq $False)
            

    #Parameters from Template User Object
    $AddressPropertyNames = @("StreetAddress","State","PostalCode","POBox","Office","Country","City")

    $SchemaNamingContext = (Get-ADRootDSE).schemaNamingContext

    $PropertiesToCopy = Get-ADObject -Filter "objectCategory -eq 'CN=Attribute-Schema,$SchemaNamingContext' -and searchflags -eq '16'" -SearchBase $SchemaNamingContext -Properties * |  
        Select -ExpandProperty lDAPDisplayname

    $PropertiesToCopy += $AddressPropertyNames

    $Password_SS = ConvertTo-SecureString -String $Password -AsPlainText -Force
    $Template_Obj = Get-ADUser -Identity $Template -Properties $PropertiesToCopy

    $OU = $Template_Obj.DistinguishedName -replace '^cn=.+?(?<!\\),'

    #Replace SAMAccountName of Template User with new account for properties like the HomeDrive that need to be dynamic
    $Template_Obj.PSObject.Properties | where {
        $_.Value -match ".*$($Template_Obj.SAMAccountName).*" -and
        $_.Name -ne "SAMAccountName" -and
        $_.IsSettable -eq $True
        } | ForEach {

            Try{
                $_.Value = $_.Value -replace "$($Template_Obj.SamAccountName)","$SAM"
            }#Try

            Catch {

                #DoNothing
            }#Catch
        }#ForEach

    #ADUser parameters
    $params = @{
            "Instance"=$Template_Obj
            "Name"=$DisplayName
            "DisplayName"=$DisplayName
            "GivenName"=$UserFirstname
            "SurName"=$UserLastname
            "Initials"=$UserInitial
            "AccountPassword"=$Password_SS
            "Enabled"=$false
            "ChangePasswordAtLogon"=$false
            "UserPrincipalName"=$UserPrincipalName
            "SAMAccountName"=$SAM
            "Path"=$OU
            "OfficePhone"=$OfficePhone
            "EmailAddress"=$Email
            "Company"=$UserCompany
            "Department"=$UserDepartment
            "Description"=$Description   
            "Title"=$UserJobTitle 
            "SmartCardLogonRequired"=$True
        }#params

    $AddressPropertyNames | foreach {$params.Add("$_","$($Template_obj."$_")")}

    New-ADUser @params
    
    Set-AdUser "$SAM" -Manager $FindSuperV -Replace @{Info="$Info"}

    $TempMembership = Get-ADUser -Identity $Template -Properties MemberOf
    $TempMembership = $TempMembership | Select -ExpandProperty MemberOf
     
    $TempMembership | Add-ADGroupMember -Members $SAM

        If($FindSuperV -EQ $Null){
        
            $NoEmail = [Microsoft.VisualBasic.Interaction]::MsgBox("Please add Manager's Email Address to their User Account!`n" + $User.SupervisorEmail, "OKOnly,SystemModal", "Error")
        }

    
}#End CreateNewUser

CreateNewUser

})

#Show Form
$Form.ShowDialog() | out-null
}
 
