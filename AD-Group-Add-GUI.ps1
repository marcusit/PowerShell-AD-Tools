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
        Title="AD User Tool" Width="800" Height="600">
      <Grid>
        <Image x:Name="Logo" HorizontalAlignment="Left" Height="100" VerticalAlignment="Top" Width="100" Source="C:\support\logo.png" Margin="10,0,0,0"/>
        <TextBlock x:Name="HeaderTextBlock" HorizontalAlignment="Left" TextWrapping="Wrap" VerticalAlignment="Top" Margin="171,38,0,0" Height="40" Width="360"><Run Foreground="#FF1A3B81" Text="This tool is used to modify Active Directory users from a foreign domain inside the local domain."/></TextBlock>
        <Label x:Name="UsernameLabel" Content="AD Username:" HorizontalAlignment="Left" Margin="15,120,0,0" VerticalAlignment="Top"/>
        <TextBlock x:Name="Errors" Text="" HorizontalAlignment="Left" TextWrapping="Wrap"  Margin="119,305,0,0" VerticalAlignment="Top" FontWeight="Bold" Width="120" />
        <TextBox x:Name="UsernameInputBox" HorizontalAlignment="Left" Height="20" TextWrapping="Wrap" VerticalAlignment="Top" Width="115" Margin="118,124,0,0" ToolTip="Enter the account name (ex: OlanderM)" />   
        <Label x:Name="FilterLabel" Content="Filter" HorizontalAlignment="Left" Margin="265,120,0,0" VerticalAlignment="Top" Width="126" FontWeight="Bold" />
        <TextBox x:Name="FilterInputBox" HorizontalAlignment="Left" Height="20" TextWrapping="Wrap" VerticalAlignment="Top" Width="405" Margin="310,124,0,0" ToolTip="Search for groups" />   
        <ListBox x:Name="GroupsBox" HorizontalAlignment="Left" Height="290" Margin="270,153,0,0" VerticalAlignment="Top" Width="446" SelectionMode="Extended" ToolTip="List of groups that the user is a member of." />
        <Label x:Name="DescriptionLabel" Content="Description" HorizontalAlignment="Left" Margin="265,450,0,0" VerticalAlignment="Top" Width="126" FontWeight="Bold" />
        <TextBox x:Name="DescriptionBox" HorizontalAlignment="Left" Height="70" Margin="270,475,0,0" TextWrapping="Wrap"  VerticalAlignment="Top" Width="446"  VerticalScrollBarVisibility="Auto" ToolTip="Group description" /> 
        <Button x:Name="ScanButton" Content="Scan" HorizontalAlignment="Left" VerticalAlignment="Top" Width="115" Margin="118,175,0,0" Height="47" FontSize="22" ToolTip="Generates a list of all groups in AD "/>
        <Button x:Name="RemoveButton" Content="Add" HorizontalAlignment="Left" VerticalAlignment="Top" Width="115" Margin="118,250,0,0" Height="47" FontSize="22" ToolTip="Adds the user to the selected group(s) "/>
        <Button x:Name="CloseButton" Content="Close" HorizontalAlignment="Left" VerticalAlignment="Top" Width="115" Margin="118,450,0,0" Height="48" FontSize="22" ToolTip="Close" />
        <TextBlock x:Name="CreatedByText" HorizontalAlignment="Left" Height="17" Margin="20,534,0,0" TextWrapping="Wrap" Text="Created by Marcus Olander in 2016" VerticalAlignment="Top" Width="163" FontSize="10" />
        <Popup x:Name="InfoPopup" HorizontalAlignment="Left" Margin="10,10,0,13" VerticalAlignment="Top" IsOpen="False" Placement="MousePoint" AllowsTransparency="True">   
        <Border Margin="0,0,8,8" Background="White" BorderThickness="1">
        <Border.Effect>
            <DropShadowEffect BlurRadius="25" Opacity="0.4"/>
         </Border.Effect>
          <TextBlock x:Name="PopupText" Background="White" FontFamily="Lucida Console">   
          </TextBlock>
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
Try { $Form=[Windows.Markup.XamlReader]::Load( $Reader ) }
Catch { Write-Host "Unable to load Windows.Markup.XamlReader. Double-check syntax and ensure .net is installed." }
$XAML.SelectNodes("//*[@Name]") | %{Set-Variable -Name "WPF$($_.Name)" -Value $Form.FindName($_.Name)}

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

# This function will add the user to the
# selected group(s).
Function Add-UserToGroups {
If (!$WPFUsernameInputBox.Text) {
  $WPFErrors.Text = "Please enter a name."
  $WPFAccountStatus.Text = ""
} Else {
  $WPFErrors.Text = ""
  If ($Password -eq $Null) {
   $Global:Password = Get-Credential -Credential "$ForeignDomainName\$Env:Username"
  }
  $UserID = (Get-ADuser -Server "$ForeignDC" -Credential $Password -Identity $WPFUsernameInputBox.Text)
    If ($Userid -ne $Null) {
      if (!$WPFGroupsBox.SelectedItems) {
        $WPFErrors.Text = "Please select at least one group. Click on Scan to list available groups."
      }
      ForEach ($SelectedGroup in $WPFGroupsBox.SelectedItems) {
        Add-ADGroupMember -Identity $SelectedGroup -Members $UserID
        If ($?) {
          $WPFErrors.Text = $WPFUsernameInputBox.Text + " added to selected group(s)."
        } Else {
          $WPFErrors.Text = "User not added to group. Incorrect username or insufficient permissions."
        }
      }
    } Else {
      $WPFErrors.Text = "User not found or incorrect password."
      Clear-Variable "Password" -Scope Global -Force
    }
  }
}

# Declating the function which simply clears
# out the content of the group membership box.
Function Clear-Groups { $WPFGroupsBox.Items.Clear() }

# This function will enumerate the groups box
# with a list of all the groups in a target OU.
Function Enumerate-Groups {
  $Global:AllGroups = (Get-AdGroup -Filter * -SearchBase $TargetOU -Properties Description | Sort-Object)
  ForEach ($Group in $AllGroups) {$WPFGroupsBox.AddText($Group.Name)}
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
Function Info-Popup {
$GroupInfo = (Get-ADGroup $WPFGroupsBox.SelectedItem -Properties CanonicalName,Created |
         Format-List Name,CanonicalName,Created,DistinguishedName,GroupCategory,GroupScope,SID  |
         Out-String).Trim()
         $WPFPopupText.Text = $GroupInfo
         $WPFInfoPopup.IsOpen = $True
}

# This function will search through the list of all
# groups and updates the groups box with the result.
Function Search-Box {
  $WPFErrors.Text = ""
  Clear-Groups
  If ($AllGroups.Name -Match $SearchString) {
    $Global:SearchResult = $AllGroups | Select-String -InputObject {$_.Name} -Pattern $SearchString
    ForEach ($Result in $SearchResult) {
      $WPFGroupsBox.AddText($Result)
    }
  }
  ElseIf ($SearchString -Like "*OemMinus*") {
    $SearchString = $SearchString -Replace "OemMinus",'-'
    $Global:SearchResult = $AllGroups | Select-String -InputObject {$_.Name} -Pattern $SearchString
    ForEach ($Result in $SearchResult) {
      $WPFGroupsBox.AddText($Result)
    }
  }
  If (!$SearchString) {
    Clear-Groups
    ForEach ($Group in $AllGroups) {
      $WPFGroupsBox.AddText($Group.Name)
    }
  }
}

# This function will update the description box with
# the group description of the currently selected group.
Function Update-DescriptionBox {
  $WPFErrors.Text = ""
  $WPFDescriptionBox.Text = ""
  $WPFDescriptionBox.AddText(($AllGroups | ? {$_.Name -eq $WPFGroupsBox.SelectedItem} | % {$_.Description}))
}

#############################
# Mouse/Keyboard triggers   #
#############################

# Detects keystrokes in the filter search box.
$WPFFilterInputBox.Add_KeyUp({
  If (($Args[1].Key -Match '[-a-z]') -or ($Args[1].Key -eq 'Back')) {
    $SearchString = $WPFFilterInputBox.Text
    Search-Box
  }
})

$WPFFilterInputBox.Add_KeyDown({
  If ($Args[1].Key -eq 'Escape') {
    $Form.Close()
  }
  ElseIf ($Args[1].Key -eq 'Return') {
    Clear-Groups ; Enumerate-Groups
  }
})

# Detects Enter/Escape in the username input field.
$WPFUsernameInputBox.Add_KeyDown({
  If ($Args[1].Key -eq 'Escape') {
    $Form.Close()
  }
  ElseIf ($Args[1].Key -eq 'Enter') {
     Add-UserToGroups
   }
})

# Various actions for the groups box.
$WPFGroupsBox.Add_KeyUp({
  If (($Args[1].Key -eq 'Down') -or ($Args[1].Key -eq 'Up')) {
    Update-DescriptionBox
  }
})

$WPFGroupsBox.Add_KeyDown({
  If ($Args[1].Key -eq 'Return') {
    Add-UserToGroups
  }
})

$WPFGroupsBox.Add_MouseUp({ Update-DescriptionBox })

$WPFGroupsBox.Add_MouseDoubleClick({ Info-Popup })

# Close the info popup window when clicking on it.
$WPFInfoPopup.Add_MouseUp({ $WPFInfoPopup.IsOpen = $False })

#####################
# GUI Buttons       #
#####################

## Scan Button
$WPFScanButton.Add_Click({ Clear-Groups ; Enumerate-Groups })

# Remove Button
$WPFRemoveButton.Add_Click({ Add-UserToGroups })

# Close Button
$WPFCloseButton.Add_Click({ $Form.Close() })

#####################
# GUI Launch        #
#####################

# Hide the PowerShell window on launch
Hide-Console

# Displays the GUI
$Form.ShowDialog() | Out-Null
