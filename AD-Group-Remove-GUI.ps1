Enter file contents #####################
# GUI Code          #
#####################

$InputXML = @"
<Window x:Class="WpfApplication1.MainWindow"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:WpfApplication1"
        mc:Ignorable="d"
        Title="AD User Tool" Width="800" Height="600">
      <Grid>
        <Image x:Name="Logo" HorizontalAlignment="Left" Height="100" VerticalAlignment="Top" Width="100" Source="C:\Support\logo.png" Margin="10,0,0,0"/>
        <TextBlock x:Name="HeaderTextBlock" HorizontalAlignment="Left" TextWrapping="Wrap" VerticalAlignment="Top" Margin="171,38,0,0" Height="40" Width="360"><Run Foreground="#FF1A3B81" Text="This tool is used to modify Active Directory users from a foreign domain inside the local domain."/></TextBlock>
        <Label x:Name="UsernameLabel" Content="AD Username:" HorizontalAlignment="Left" Margin="15,120,0,0" VerticalAlignment="Top"/>
        <TextBox x:Name="UsernameInputBox" HorizontalAlignment="Left" Height="23" TextWrapping="Wrap" VerticalAlignment="Top" Width="115" Margin="118,123,0,0" ToolTip="Enter the account name (ex: OlanderM)" />   
        <Label x:Name="GroupMembershipsLabel" Content="Group Memberships" HorizontalAlignment="Left" Margin="265,120,0,0" VerticalAlignment="Top" Width="126" FontWeight="Bold" />
        <TextBlock x:Name="AccountStatusLabel" HorizontalAlignment="Left" Margin="502,125,0,0" TextWrapping="Wrap" Text="Account is:" VerticalAlignment="Top" />
        <TextBlock x:Name="AccountStatus" HorizontalAlignment="Left" Height="23" Margin="567,125,0,0" TextWrapping="Wrap" Text="" VerticalAlignment="Top" Width="58" ToolTip="Current account status" FontWeight="Bold" />
        <ListBox x:Name="GroupsBox" HorizontalAlignment="Left" Height="290" Margin="270,153,0,0" VerticalAlignment="Top" Width="446" SelectionMode="Extended" ToolTip="List of groups that the user is a member of." />
        <Label x:Name="DescriptionLabel" Content="Description" HorizontalAlignment="Left" Margin="265,450,0,0" VerticalAlignment="Top" Width="126" FontWeight="Bold" />
        <TextBox x:Name="DescriptionBox" HorizontalAlignment="Left" Height="70" Margin="270,475,0,0" TextWrapping="Wrap"  VerticalAlignment="Top" Width="446"  VerticalScrollBarVisibility="Auto" ToolTip="Group description" /> 
        <Button x:Name="ScanButton" Content="Scan" HorizontalAlignment="Left" VerticalAlignment="Top" Width="115" Margin="118,175,0,0" Height="47" FontSize="22" ToolTip="Searches Active Directory for the user" />
        <Button x:Name="RemoveButton" Content="Remove" HorizontalAlignment="Left" VerticalAlignment="Top" Width="115" Margin="118,250,0,0" Height="47" FontSize="22" ToolTip="Removes the selected group memberships" />
        <Button x:Name="CloseButton" Content="Close" HorizontalAlignment="Left" VerticalAlignment="Top" Width="115" Margin="118,450,0,0" Height="48" FontSize="22" ToolTip="Close" />
        <TextBlock x:Name="CreatedByText" HorizontalAlignment="Left" Height="17" Margin="20,534,0,0" TextWrapping="Wrap" Text="Created by Marcus Olander in 2016" VerticalAlignment="Top" Width="163" FontSize="10" />
        <Popup x:Name="InfoPopup" HorizontalAlignment="Left" Margin="10,10,0,13" VerticalAlignment="Top" IsOpen="False" Placement="MousePoint" AllowsTransparency="True">   
        <Border Margin="0,0,8,8" Background="White" BorderThickness="1">
          <Border.Effect>
            <DropShadowEffect BlurRadius="25" Opacity="0.4"/>
          </Border.Effect>
          <TextBlock x:Name="PopupText" Background="White" FontFamily="Lucida Console" />   
        </Border>
      </Popup>
    </Grid>
</Window>
"@

# Cleans and processess the XML code so that it can be converted into a
# PowerShell GUI. The error message gets thrown if there is a problem, usually
# because .NET is not installed. The script should run on Windows Server 2012 R2.
$InputXML = $InputXML -Replace 'mc:Ignorable="d"','' -Replace "x:N",'N'  -Replace '^<Win.*', '<Window'
[void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework')
[xml]$XAML = $InputXML
$Reader=(New-Object System.Xml.XmlNodeReader $XAML)
Try { $Form=[Windows.Markup.XamlReader]::Load($Reader) }
Catch { Write-Host "Unable to load Windows.Markup.XamlReader. Double-check syntax and ensure .net is installed." }
$XAML.SelectNodes("//*[@Name]") | % { Set-Variable -Name "WPF$($_.Name)" -Value $Form.FindName($_.Name) }

#####################
# Script Start      #
#####################

# Import Active Directory module as well as supress any error messages.
# We also add some code needed to hide the PowerShell window before
# launching the GUI.
Import-Module ActiveDirectory
$ErrorActionPreference = "SilentlyContinue"
Add-Type -Name Window -Namespace Console -MemberDefinition '
 [DllImport("Kernel32.dll")]
 public static extern IntPtr GetConsoleWindow();

 [DllImport("user32.dll")]
 public static extern bool ShowWindow(IntPtr hWnd, Int32 nCmdShow);
 '

# Grabbing the default domain controllers for each domain.
# We also set the target OU. To speed up the program this path should be
# as narrow as possible. This is something that could use *CHANGE*.
# We currently don't need the local DC variable, so it is commented out.
#$LocalDC = (Get-ADDomainController).HostName
$ForeignDomainName = "contoso.com"
$ForeignDC = (Get-ADDomainController -Domain "$ForeignDomainName" -Discover).HostName
$TargetOU = "OU=Groups,DC=contoso,DC=com"

#####################
# Functions         #
#####################

# Declating the function which simply clears
# out the content of the group membership box.
Function Clear-Groups { $WPFGroupsBox.Items.Clear() }

# This function is called before the GUI is shown to hide the
# PowerShell Window.
Function Hide-Console {
  $ConsolePtr = [Console.Window]::GetConsoleWindow()
  [Console.Window]::ShowWindow($ConsolePtr, 0)
}

# This function will generate a popup window when a group item
# is double-clicked, and will display some basic information.
# Can use *CHANGE* if more/less info is desired.
Function Info-Popup {
  $GroupInfo = (Get-ADGroup $WPFGroupsBox.SelectedItem -Properties CanonicalName,Created |
           Format-List Name,CanonicalName,Created,DistinguishedName,GroupCategory,GroupScope,SID  |
           Out-String).Trim()
           $WPFPopupText.Text = $GroupInfo
           $WPFInfoPopup.IsOpen = $True
}

# This function gets called when the "Remove" button is clicked.
# It will remove the user from the selected group(s).
Function Remove-ADUserFromGroup {
  If (($WPFAccountStatus.Text -ne "Disabled") -and ($WPFAccountStatus.Text -ne "Enabled")) {
    Clear-Groups
    $WPFGroupsBox.AddText("Please enter a name and click Scan.")
    $WPFAccountStatus.Text = ""
  } Else {
    ForEach ($SelectedGroup in $WPFGroupsBox.SelectedItems) {
      $GroupDN = (Get-ADGroup -Identity $SelectedGroup).DistinguishedName
      $Remove = [ADSI]("LDAP://" + $GroupDN)
      $Remove.Remove($User.ADsPath)
    }
    Clear-Groups ; Scan-ADUser
  }
}

# Declaring the "Scan-ADUser" function, which gets called when
# the user press the "Scan" button (or press enter).
Function Scan-ADUser {
If (!$WPFUsernameInputBox.Text) {
  Clear-Groups
  $WPFGroupsBox.AddText("Please enter a name.")
  $WPFAccountStatus.Text = ""
} Else {
    If ($Password -eq $Null) {
      $Global:Password = Get-Credential -Credential "$ForeignDomainName\$Env:Username"
    }
    If ((Get-ADUser -Identity $WPFUsernameInputBox.Text -Server "$ForeignDC" -Credential $Password).Enabled) {
      $WPFAccountStatus.Text = "Enabled"
    } Else {
      $WPFAccountStatus.Text = "Disabled"
    }
    $UserSID = (Get-ADuser -Server "$ForeignDC" -Credential $Password -Identity $WPFUsernameInputBox.Text).SID
    If ($UserSID -ne $Null) {
    # This will need *CHANGE* to work in other domains.
      $Global:UserCN = "CN=$UserSID,CN=ForeignSecurityPrincipals,DC=contoso,DC=com"
      $Global:User = [ADSI]"LDAP://$UserCN"
      $Groups = $User.MemberOf
      ForEach ($Group in $Groups) {
        $WPFGroupsBox.AddText((Get-ADGroup -Identity $Group).Name)
      }
    } Else {
      Clear-Groups
      $WPFGroupsBox.AddText("User not found or incorrect Password.")
      Clear-Variable "Password" -Scope Global -Force
      $WPFAccountStatus.Text = ""
    }
  }
}

# This function will update the description box with
# the group description of the currently selected group.
Function Update-DescriptionBox {
  $WPFDescriptionBox.Text = ""
  $WPFDescriptionBox.AddText(((Get-ADGroup $WPFGroupsBox.SelectedItem -Properties Description).Description | Out-String))
}

#############################
# Mouse/Keyboard triggers   #
#############################

## Username input box when pressing Enter or Escape
$WPFUsernameInputBox.Add_KeyDown({
  If ($Args[1].Key -eq 'Return') {
    Clear-Groups ; Scan-ADUser
  }
  ElseIf ($Args[1].Key -eq 'Escape') {
    $form.Close()
  }
})

## Box listing groups, various actions depending on mouse/keyboard action
$WPFGroupsBox.Add_MouseUp({ Update-DescriptionBox })

$WPFGroupsBox.Add_KeyUp({
  if (($Args[1].Key -eq 'Down') -or ($Args[1].Key -eq 'Up')) {
    Update-DescriptionBox
  }
})

$WPFGroupsBox.Add_MouseDoubleClick({ Info-Popup })

$WPFGroupsBox.Add_KeyDown({
 if ($Args[1].Key -eq 'Return') {
      Remove-ADUserFromGroup
    }
})

# Close the info popup window when clicking on it.
$WPFInfoPopup.Add_MouseUp({ $WPFInfoPopup.IsOpen = $False })

#####################
# GUI Buttons       #
#####################

## Scan Button
$WPFScanButton.Add_Click({Clear-Groups ; Scan-ADUser})

# Remove Button
$WPFRemoveButton.Add_Click({ Remove-ADUserFromGroup })

# Close Button
$WPFCloseButton.Add_Click({ $Form.Close() })

#####################
# GUI Launch        #
#####################

# Hide the PowerShell window on launch
Hide-Console

# Displays the GUI
$Form.ShowDialog() | Out-Nullhere
