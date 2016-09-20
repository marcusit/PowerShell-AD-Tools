$LocalDomain = ((Get-ADDomain).Name).ToUpper()

################################################################################################################################
################################################################################################################################
# These three variables still need to be made changeable through the GUI.
################################################################################################################################
################################################################################################################################
$ForeignDomain = "CONTOSO"
$TargetOU = "OU=Contoso Groups,DC=test,DC=lab"
$Logo = "C:\support\logo.png"
################################################################################################################################
################################################################################################################################

#####################
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
        Title="Active Directory Tool" Width="800" Height="650">
  <TabControl x:Name="tabControl" HorizontalAlignment="Left" Height="630" VerticalAlignment="Top" Width="800" Margin="-3,0,0,0">
    <TabItem Header="Add User" Height="28" VerticalAlignment="Top" FontSize="12" Margin="0,0,-2,0">
      <Grid Width="800" Margin="0,0,-8,-7" Height="580" VerticalAlignment="Top">
        <Image x:Name="ADDLogo" HorizontalAlignment="Left" Height="100" VerticalAlignment="Top" Width="100" Source="$Logo" Margin="10,0,0,0" />
        <TextBlock x:Name="ADDHeaderTextBlock" HorizontalAlignment="Left" TextWrapping="Wrap" VerticalAlignment="Top" Margin="171,38,0,0" Height="40" Width="360">
          <Run Text="This AD tool is a work in progress. It will be used to compliment dsa.msc when working with users in another domain." Foreground="#FF1A3B81" />
        </TextBlock>
        <Label x:Name="ADDUsernameLabelDomain" Content="$ForeignDomain" HorizontalAlignment="Left" Margin="15,105,0,0" VerticalAlignment="Top"/>
        <Label x:Name="ADDUsernameLabelUsername" Content="Username:" HorizontalAlignment="Left" Margin="15,120,0,0" VerticalAlignment="Top"/>
        <TextBlock x:Name="ADDErrors" Text="" HorizontalAlignment="Left" TextWrapping="Wrap"  Margin="119,305,0,0" VerticalAlignment="Top" FontWeight="Bold" Width="120" />
        <TextBox x:Name="ADDUsernameInputBox" HorizontalAlignment="Left" Height="20" TextWrapping="Wrap" VerticalAlignment="Top" Width="115" Margin="118,124,0,0" ToolTip="Enter the account name (ex: OlanderM)" />   
        <Label x:Name="ADDFilterLabel" Content="Filter" HorizontalAlignment="Left" Margin="265,120,0,0" VerticalAlignment="Top" Width="126" FontWeight="Bold" />
        <TextBox x:Name="ADDFilterInputBox" HorizontalAlignment="Left" Height="20" TextWrapping="Wrap" VerticalAlignment="Top" Width="405" Margin="310,124,0,0" ToolTip="Search for groups" />   
        <ListBox x:Name="ADDGroupsBox" HorizontalAlignment="Left" Height="290" Margin="270,153,0,0" VerticalAlignment="Top" Width="446" SelectionMode="Extended" ToolTip="List of groups that the user is a member of" />
        <Label x:Name="ADDDescriptionLabel" Content="Description" HorizontalAlignment="Left" Margin="265,450,0,0" VerticalAlignment="Top" Width="126" FontWeight="Bold" />
        <TextBox x:Name="ADDDescriptionBox" HorizontalAlignment="Left" Height="70" Margin="270,475,0,0" TextWrapping="Wrap"  VerticalAlignment="Top" Width="446"  VerticalScrollBarVisibility="Auto" ToolTip="Group description" /> 
        <Button x:Name="ADDScanButton" Content="Scan" HorizontalAlignment="Left" VerticalAlignment="Top" Width="115" Margin="118,175,0,0" Height="47" FontSize="22" ToolTip="Generates a list of all groups in AD "/>
        <Button x:Name="ADDAddButton" Content="Add" HorizontalAlignment="Left" VerticalAlignment="Top" Width="115" Margin="118,250,0,0" Height="47" FontSize="22" ToolTip="Adds the user to the selected group(s)" />
        <Button x:Name="ADDCloseButton" Content="Close" HorizontalAlignment="Left" VerticalAlignment="Top" Width="115" Margin="118,450,0,0" Height="48" FontSize="22" ToolTip="Close" />
        <TextBlock x:Name="ADDCreatedByText" HorizontalAlignment="Left" Height="17" Margin="20,534,0,0" TextWrapping="Wrap" Text="Created by Marcus Olander in 2016" VerticalAlignment="Top" Width="163" FontSize="10" />
        <Popup x:Name="ADDInfoPopup" HorizontalAlignment="Left" Margin="10,10,0,13" VerticalAlignment="Top" IsOpen="False" Placement="MousePoint" AllowsTransparency="True">   
          <Border Margin="0,0,8,8" Background="White" BorderThickness="1">
            <Border.Effect>
              <DropShadowEffect BlurRadius="25" Opacity="0.4" />
            </Border.Effect>
            <TextBlock x:Name="ADDPopupText" Background="White" FontFamily="Lucida Console" />
          </Border>
        </Popup>
      </Grid>
    </TabItem>
    <TabItem Header="Remove User" Height="28" VerticalAlignment="Top" Margin="0,0,-2,0">
      <Grid>
        <Image x:Name="REMLogo" HorizontalAlignment="Left" Height="100" VerticalAlignment="Top" Width="100" Source="C:\Support\PCI.png" Margin="10,0,0,0" />
        <TextBlock x:Name="REMHeaderTextBlock" HorizontalAlignment="Left" TextWrapping="Wrap" VerticalAlignment="Top" Margin="171,38,0,0" Height="40" Width="360">
          <Run Foreground="#FF1A3B81" Text="This AD tool is a work in progress. It will be used to compliment dsa.msc when working with users in another domain." />
        </TextBlock>
        <Label x:Name="REMUsernameLabelDomain" Content="$ForeignDomain" HorizontalAlignment="Left" Margin="15,105,0,0" VerticalAlignment="Top" />
        <Label x:Name="REMUsernameLabelUsername" Content="Username:" HorizontalAlignment="Left" Margin="15,120,0,0" VerticalAlignment="Top" />
        <TextBox x:Name="REMUsernameInputBox" HorizontalAlignment="Left" Height="20" TextWrapping="Wrap" VerticalAlignment="Top" Width="115" Margin="118,124,0,0" ToolTip="Enter the account name (ex: OlanderM)" />   
        <Label x:Name="REMGroupMembershipsLabel" Content="Group Memberships" HorizontalAlignment="Left" Margin="265,120,0,0" VerticalAlignment="Top" Width="126" FontWeight="Bold" />
        <TextBlock x:Name="REMREMAccountStatusLabel" HorizontalAlignment="Left" Margin="502,125,0,0" TextWrapping="Wrap" Text="Account is:" VerticalAlignment="Top" />
        <TextBlock x:Name="REMAccountStatus" HorizontalAlignment="Left" Height="23" Margin="567,125,0,0" TextWrapping="Wrap" Text="" VerticalAlignment="Top" Width="58" ToolTip="Current account status" FontWeight="Bold" />
        <ListBox x:Name="REMGroupsBox" HorizontalAlignment="Left" Height="290" Margin="270,153,0,0" VerticalAlignment="Top" Width="446" SelectionMode="Extended" ToolTip="List of groups that the user is a member of" />
        <Label x:Name="REMDescriptionLabel" Content="Description" HorizontalAlignment="Left" Margin="265,450,0,0" VerticalAlignment="Top" Width="126" FontWeight="Bold" />
        <TextBox x:Name="REMDescriptionBox" HorizontalAlignment="Left" Height="70" Margin="270,475,0,0" TextWrapping="Wrap"  VerticalAlignment="Top" Width="446"  VerticalScrollBarVisibility="Auto" ToolTip="Group description" /> 
        <Button x:Name="REMScanButton" Content="Scan" HorizontalAlignment="Left" VerticalAlignment="Top" Width="115" Margin="118,175,0,0" Height="47" FontSize="22" ToolTip="Searches Active Directory for the user" />
        <Button x:Name="REMRemoveButton" Content="Remove" HorizontalAlignment="Left" VerticalAlignment="Top" Width="115" Margin="118,250,0,0" Height="47" FontSize="22" ToolTip="Removes the selected group memberships" />
        <Button x:Name="REMCloseButton" Content="Close" HorizontalAlignment="Left" VerticalAlignment="Top" Width="115" Margin="118,450,0,0" Height="48" FontSize="22" ToolTip="Close" />
        <TextBlock x:Name="REMCreatedByText" HorizontalAlignment="Left" Height="17" Margin="20,534,0,0" TextWrapping="Wrap" Text="Created by Marcus Olander in 2016" VerticalAlignment="Top" Width="163" FontSize="10" />
        <Popup x:Name="REMInfoPopup" HorizontalAlignment="Left" Margin="10,10,0,13" VerticalAlignment="Top" IsOpen="False" Placement="MousePoint" AllowsTransparency="True">   
          <Border Margin="0,0,8,8" Background="White" BorderThickness="1">
            <Border.Effect>
              <DropShadowEffect BlurRadius="25" Opacity="0.4" />
            </Border.Effect>
            <TextBlock x:Name="REMPopupText" Background="White" FontFamily="Lucida Console" />   
          </Border>
        </Popup>
      </Grid>
    </TabItem>
    <TabItem Header="Account Info" Height="28" VerticalAlignment="Top" Margin="0,0,0,0">
      <Grid>
        <Image x:Name="INFLogo" HorizontalAlignment="Left" Height="100" VerticalAlignment="Top" Width="100" Source="C:\Support\PCI.png" Margin="10,0,0,0" />
        <TextBlock x:Name="INFHeaderTextBlock" HorizontalAlignment="Left" TextWrapping="Wrap" VerticalAlignment="Top" Margin="171,38,0,0" Height="40" Width="360">
          <Run Foreground="#FF1A3B81" Text="This AD tool is a work in progress. It will be used to compliment dsa.msc when working with users in another domain." />
        </TextBlock>
        <Label x:Name="INFUsernameLabelDomain" Content="$ForeignDomain" HorizontalAlignment="Left" Margin="15,105,0,0" VerticalAlignment="Top" />
        <Label x:Name="INFUsernameLabelUsername" Content="Username:" HorizontalAlignment="Left" Margin="15,120,0,0" VerticalAlignment="Top" />
        <TextBox x:Name="INFUsernameInputBox" HorizontalAlignment="Left" Height="20" TextWrapping="Wrap" VerticalAlignment="Top" Width="115" Margin="118,124,0,0" ToolTip="Enter the account name (ex: OlanderM)" />   
        <Button x:Name="INFScanButton" Content="Scan" HorizontalAlignment="Left" VerticalAlignment="Top" Width="115" Margin="118,175,0,0" Height="47" FontSize="22" ToolTip="Grabs the account status" />
        <Button x:Name="INFUnlockButton" Content="Unlock" HorizontalAlignment="Left" VerticalAlignment="Top" Width="115" Margin="118,250,0,0" Height="47" FontSize="22" ToolTip="Unlock the account" />
        <Button x:Name="INFResetkButton" Content="Reset" HorizontalAlignment="Left" VerticalAlignment="Top" Width="115" Margin="118,325,0,0" Height="47" FontSize="22" ToolTip="Reset the password to 'Password1'" />
        <Button x:Name="INFCloseButton" Content="Close" HorizontalAlignment="Left" VerticalAlignment="Top" Width="115" Margin="118,450,0,0" Height="48" FontSize="22" ToolTip="Close" />
      </Grid>
    </TabItem>
    <TabItem Header="Mass Operations" Height="28" VerticalAlignment="Top" Margin="0,0,0,0">
      <Grid>
        <Image x:Name="MASLogo" HorizontalAlignment="Left" Height="100" VerticalAlignment="Top" Width="100" Source="C:\Support\PCI.png" Margin="10,0,0,0" />
        <TextBlock x:Name="MASHeaderTextBlock" HorizontalAlignment="Left" TextWrapping="Wrap" VerticalAlignment="Top" Margin="171,38,0,0" Height="40" Width="360">
          <Run Foreground="#FF1A3B81" Text="This AD tool is a work in progress. It will be used to compliment dsa.msc when working with users in another domain." />
        </TextBlock>
        <Label x:Name="MASGroupnameLabelDomain" Content="$LocalDomain" HorizontalAlignment="Left" Margin="15,105,0,0" VerticalAlignment="Top" />
        <Label x:Name="MASGroupnameLabelGroupname" Content="Groupname:" HorizontalAlignment="Left" Margin="15,120,0,0" VerticalAlignment="Top" />
        <TextBox x:Name="MASGroupnameInputBox" HorizontalAlignment="Left" Height="20" TextWrapping="Wrap" VerticalAlignment="Top" Width="115" Margin="118,124,0,0" ToolTip="Enter the account name (ex: OlanderM)" /> 
        <RichTextBox x:Name="MASUsernameInputBox" Height="421" Width="446" Margin="193,77,0,0" Block.LineHeight="1">
          <FlowDocument>
            <Paragraph>
              <Run Text=""/>
            </Paragraph>
          </FlowDocument>
        </RichTextBox>
        <Button x:Name="MASScanButton" Content="Scan" HorizontalAlignment="Left" VerticalAlignment="Top" Width="115" Margin="118,175,0,0" Height="47" FontSize="22" ToolTip="Grabs the account status" />
        <Button x:Name="MASUnlockButton" Content="Unlock" HorizontalAlignment="Left" VerticalAlignment="Top" Width="115" Margin="118,250,0,0" Height="47" FontSize="22" ToolTip="Unlock the account" />
        <Button x:Name="MASResetkButton" Content="Reset" HorizontalAlignment="Left" VerticalAlignment="Top" Width="115" Margin="118,325,0,0" Height="47" FontSize="22" ToolTip="Reset the password to 'Password1'" />
        <Button x:Name="MASCloseButton" Content="Close" HorizontalAlignment="Left" VerticalAlignment="Top" Width="115" Margin="118,450,0,0" Height="48" FontSize="22" ToolTip="Close" />
      </Grid>
    </TabItem>
  </TabControl>
</Window>
"@

# Cleans and processess the XML code so that it can be converted into a
# PowerShell GUI. The error message gets thrown if there is a problem, usually
# because .NET is not installed. The script should run on Windows Server 2012 R2.
$InputXML = $InputXML -Replace 'mc:Ignorable="d"','' -Replace "x:N",'N'  -Replace '^<Win.*', '<Window'
[void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework')
[xml]$XAML = $InputXML
$Reader=(New-Object System.Xml.XmlNodeReader $XAML)
Try { $Form=[Windows.Markup.XamlReader]::Load( $Reader ) }
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

$ForeignDC = (Get-ADDomainController -Domain $ForeignDomain -Discover).HostName
$LocalDomainDN = (Get-ADDomain).DistinguishedName


#####################
# Functions         #
#####################

# This function will add the user to the
# selected group(s).
Function ADDAdd-UserToGroups {
If (!$WPFADDUsernameInputBox.Text) {
  $WPFADDErrors.Text = "Please enter a name."
  $WPFREMAccountStatus.Text = ""
} Else {
  $WPFADDErrors.Text = ""
  If ($Password -eq $Null) {
   $Global:Password = Get-Credential -Credential "$ForeignDomain\$Env:Username"
  }
  $UserID = (Get-ADuser -Server "$ForeignDC" -Credential $Password -Identity $WPFADDUsernameInputBox.Text)
    If ($Userid -ne $Null) {
      if (!$WPFADDGroupsBox.SelectedItems) {
        $WPFADDErrors.Text = "Please select at least one group. Click on Scan to list available groups."
      }
      ForEach ($ADDSelectedGroup in $WPFADDGroupsBox.SelectedItems) {
        Add-ADGroupMember -Identity $ADDSelectedGroup -Members $UserID
        If ($?) {
          $WPFADDErrors.Text = $WPFADDUsernameInputBox.Text + " added to selected group(s)."
        } Else {
          $WPFADDErrors.Text = "User not added to group. Incorrect username or insufficient permissions."
        }
      }
    } Else {
      $WPFADDErrors.Text = "User not found or incorrect password."
      Clear-Variable "Password" -Scope Global -Force
    }
  }
}

# Declating the function which simply clears
# out the content of the group membership box.
Function ADDClear-Groups { $WPFADDGroupsBox.Items.Clear() }
Function REMClear-Groups { $WPFREMGroupsBox.Items.Clear() }

# This function will enumerate the groups box
# with a list of all the groups in a target OU.
Function ADDEnumerate-Groups {
  $Global:ADDAllGroups = (Get-AdGroup -Filter * -SearchBase $TargetOU -Properties Description | Sort-Object)
  ForEach ($ADDGroup in $ADDAllGroups) { $WPFADDGroupsBox.AddText($ADDGroup.name) }
  $WPFADDScanButton.IsEnabled = $False
}
# This function is called before the GUI is shown to hide the
# PowerShell Window.
Function Hide-Console {
  $ConsolePtr = [Console.Window]::GetConsoleWindow()
  [Console.Window]::ShowWindow($ConsolePtr, 0)
}

# This function will generate a popup window when a group item
# is double-clicked, and will display some basic information.
# Can use *CHANGE* if more/less info is desired.
Function ADDInfo-Popup {
  $ADDGroupInfo = (Get-ADGroup $WPFADDGroupsBox.SelectedItem -Properties CanonicalName,Created |
           Format-List Name,CanonicalName,Created,DistinguishedName,GroupCategory,GroupScope,SID  |
           Out-String).Trim()
           $WPFADDPopupText.Text = $ADDGroupInfo
           $WPFADDInfoPopup.IsOpen = $True
}

Function REMInfo-Popup {
  $REMGroupInfo = (Get-ADGroup $WPFREMGroupsBox.SelectedItem -Properties CanonicalName,Created |
           Format-List Name,CanonicalName,Created,DistinguishedName,GroupCategory,GroupScope,SID  |
           Out-String).Trim()
           $WPFREMPopupText.Text = $REMGroupInfo
           $WPFREMInfoPopup.IsOpen = $True
}

# This function will search through the list of all
# groups and updates the groups box with the result.
Function ADDSearch-Box {
  $WPFADDErrors.Text = ""
  ADDClear-Groups
  If ($ADDAllGroups.Name -Match $ADDSearchString) {
    $Global:ADDSearchResult = $ADDAllGroups | Select-String -InputObject {$_.Name} -Pattern $ADDSearchString
    ForEach ($ADDResult in $ADDSearchResult) {
      $WPFADDGroupsBox.AddText($ADDResult)
    }
  }
  Elseif ($ADDSearchString -Like "*OemMinus*") {
    $ADDSearchString = $ADDSearchString -Replace "OemMinus",'-'
    $Global:ADDSearchResult = $ADDAllGroups | Select-String -InputObject {$_.Name} -Pattern $ADDSearchString
    ForEach ($ADDResult in $ADDSearchResult) {
      $WPFADDGroupsBox.AddText($ADDResult)
    }
  }
  If (!$ADDSearchString) {
    ADDClear-Groups
    ForEach ($ADDGroup in $ADDAllGroups) {
      $WPFADDGroupsBox.AddText($ADDGroup.Name)
    }
  }
}

# This function gets called when the "Remove" button is clicked.
# It will remove the user from the selected groups.
Function REMRemove-ADUserFromGroup {
  If (($WPFREMAccountStatus.Text -ne "Disabled") -and ($WPFREMAccountStatus.Text -ne "Enabled")) {
    REMClear-Groups
    $WPFREMGroupsBox.AddText("Please enter a name and click Scan.")
    $WPFREMAccountStatus.Text = ""
  } Else {
    ForEach ($REMSelectedGroup in $WPFREMGroupsBox.SelectedItems) {
      $REMGroupDN = (Get-ADGroup -Identity $REMSelectedGroup).DistinguishedName
      $REMRemove = [ADSI]("LDAP://" + $REMGroupDN)
      $REMRemove.Remove($REMUser.ADsPath)
    }
    REMClear-Groups ; REMScan-ADUser
  }
}

# Declaring the "Scan-ADUser" function, which gets called when
# the user press the "Scan" button (or press enter).
Function REMScan-ADUser {
If (!$WPFREMUsernameInputBox.Text) {
  REMClear-Groups
  $WPFREMGroupsBox.AddText("Please enter a username.")
  $WPFREMAccountStatus.Text = ""
} Else {
    If ($Password -eq $Null) {
      $Global:Password = Get-Credential -Credential "$ForeignDomain\$Env:Username"
    }
    If ((Get-ADUser -Identity $WPFREMUsernameInputBox.Text -Server "$ForeignDC" -Credential $Password).Enabled) {
      $WPFREMAccountStatus.Text = "Enabled"
    } Else {
      $WPFREMAccountStatus.Text = "Disabled"
    }
    $REMUserSID = (Get-ADuser -Server "$ForeignDC" -Credential $Password -Identity $WPFREMUsernameInputBox.Text).SID
    If ($REMUserSID -ne $Null) {
      $Global:REMUserCN = "CN=" + $REMUserSID + "," + (Get-ADDomain).ForeignSecurityPrincipalsContainer
      $Global:REMUser = [ADSI]"LDAP://$REMUserCN"
      $REMGroups = $REMUser.MemberOf
      ForEach ($REMGroup in $REMGroups) {
        $WPFREMGroupsBox.AddText((Get-ADGroup -Identity $REMGroup).Name)
      }
    } Else {
      REMClear-Groups
      $WPFREMGroupsBox.AddText("User not found or incorrect Password.")
      Clear-Variable "Password" -Scope Global -Force
      $WPFREMAccountStatus.Text = ""
    }
  }
}

# This function will update the description box with
# the group description of the currently selected group.
Function ADDUpdate-DescriptionBox {
  $WPFADDErrors.Text = ""
  $WPFADDDescriptionBox.Text = ""
  $WPFADDDescriptionBox.AddText(($ADDAllGroups | ? {$_.Name -eq $WPFADDGroupsBox.SelectedItem} | % {$_.Description}))
}

Function REMUpdate-DescriptionBox {
  $WPFREMDescriptionBox.Text = ""
  $WPFREMDescriptionBox.AddText(((Get-ADGroup $WPFREMGroupsBox.SelectedItem -Properties Description).Description | Out-String))
}

#############################
# Mouse/Keyboard triggers   #
#############################

# Detects keystrokes in the filter search box.
$WPFADDFilterInputBox.Add_KeyUp({
  If (($Args[1].Key -match '[-a-z]') -or ($Args[1].Key -eq 'Back')) {
    $ADDSearchString = $WPFADDFilterInputBox.Text
    ADDSearch-Box
  }
})

$WPFADDFilterInputBox.Add_KeyDown({
  If ($Args[1].Key -eq 'Escape') {
    $Form.Close()
  }
  Elseif ($Args[1].Key -eq 'Return') {
    ADDClear-Groups ; ADDEnumerate-Groups
  }
})

# Detects Enter/Escape in the username input field.
$WPFADDUsernameInputBox.Add_KeyDown({
  If ($Args[1].Key -eq 'Escape') {
    $Form.Close()
  }
  Elseif ($Args[1].Key -eq 'Enter') {
     ADDAdd-UserToGroups
   }
})

# Various actions for the groups box.
$WPFADDGroupsBox.Add_KeyUp({
  If (($Args[1].Key -eq 'Down') -or ($Args[1].Key -eq 'Up')) {
    ADDUpdate-DescriptionBox
  }
})

$WPFADDGroupsBox.Add_KeyDown({
  If ($Args[1].Key -eq 'Return') {
    ADDAdd-UserToGroups
  }
})

$WPFADDGroupsBox.Add_MouseUp({ ADDUpdate-DescriptionBox })
$WPFREMGroupsBox.Add_MouseUp({ REMUpdate-DescriptionBox })

$WPFADDGroupsBox.Add_MouseDoubleClick({ ADDInfo-Popup })
$WPFREMGroupsBox.Add_MouseDoubleClick({ REMInfo-Popup })

# Close the info popup window when clicking on it.
$WPFADDInfoPopup.Add_MouseUp({ $WPFADDInfoPopup.IsOpen = $False })
$WPFREMInfoPopup.Add_MouseUp({ $WPFREMInfoPopup.IsOpen = $False })

## Username input box when pressing Enter or Escape
$WPFREMUsernameInputBox.Add_KeyDown({
  If ($Args[1].Key -eq 'Return') {
    REMClear-Groups ; REMScan-ADUser
  }
  Elseif ($Args[1].Key -eq 'Escape') {
    $form.Close()
  }
})

$WPFREMGroupsBox.Add_KeyUp({
  if (($Args[1].Key -eq 'Down') -or ($Args[1].Key -eq 'Up')) {
    REMUpdate-DescriptionBox
  }
})

$WPFREMGroupsBox.Add_KeyDown({
 if ($Args[1].Key -eq 'Return') {
      REMRemove-ADUserFromGroup
    }
})

#####################
# GUI Buttons       #
#####################

## Scan Button
$WPFADDScanButton.Add_Click({ ADDClear-Groups ; ADDEnumerate-Groups })
$WPFREMScanButton.Add_Click({ REMClear-Groups ; REMScan-ADUser})

# Add Button - Add tab
$WPFADDAddButton.Add_Click({ ADDAdd-UserToGroups })

# Remove Button - Remove tab
$WPFREMRemoveButton.Add_Click({ REMRemove-ADUserFromGroup })

# Close Button
$WPFREMCloseButton.Add_Click({ $Form.Close() })
$WPFADDCloseButton.Add_Click({ $Form.Close() })

#[System.Windows.Forms.RichTextBoxStreamType]::PlainText

$MASListOfUsers = New-Object System.Windows.Documents.TextRange($WPFMASUsernameInputBox.Document.ContentStart,$WPFMASUsernameInputBox.Document.ContentEnd)
$WPFMASCloseButton.Add_Click({ Foreach ($User in ($MASListOfUsers.Text -Split '\s{2,}')) { Write-Host $User } })

#####################
# GUI Launch        #
#####################

# Hide the PowerShell window on launch
Hide-Console

# Displays the GUI
$Form.ShowDialog() | Out-Null
