powershell.exe -windowstyle hidden -command {
#Adding PresentationFramework for Run XAML Variable.
Add-Type -AssemblyName PresentationFramework
#Read Notepad file for get Default Value in this application
[ARRAY]$notepad = Get-Content .\SDMail.txt
$PCName2 = $notepad[0]
$username2 = $notepad[1]
$userid2 = $notepad[2]
$subject2 = $notepad[3]
$to2 =$notepad[4]
$cc2 =$notepad[5]
#Write XAML code in variable and Declear type as a xml.
[xml]$xaml= @"
<Window
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        x:Name="Window"
        Title="MainWindow" Height="350" Width="630">
	<Grid Margin="0,0,0,14">
		<Grid.ColumnDefinitions>
			<ColumnDefinition Width="99*"/>
			<ColumnDefinition Width="540*"/>
			<ColumnDefinition/>
		</Grid.ColumnDefinitions>
		<Label Content="PC Name :" HorizontalAlignment="Left" VerticalAlignment="Top" FontFamily="Calibri" FontSize="24" FontWeight="Bold" Margin="10,10,0,0" Grid.ColumnSpan="2" Height="39" Width="112"/>
		<Label Content="User ID :" HorizontalAlignment="Left" VerticalAlignment="Top" FontFamily="Calibri" FontSize="24" FontWeight="Bold" Margin="10,88,0,0" Height="39" Width="89"/>
		<Label Content="User Name :" HorizontalAlignment="Left" VerticalAlignment="Top" FontFamily="Calibri" FontSize="24" FontWeight="Bold" Margin="10,49,0,0" Grid.ColumnSpan="2" Height="39" Width="133"/>
		<Label Content="Subject :" HorizontalAlignment="Left" VerticalAlignment="Top" FontFamily="Calibri" FontSize="24" FontWeight="Bold" Margin="10,127,0,0" Height="39" Width="89"/>
		<Label Content="To :" HorizontalAlignment="Left" VerticalAlignment="Top" FontFamily="Calibri" FontSize="24" FontWeight="Bold" Margin="10,166,0,0" Height="39" Width="45"/>
		<Label Content="CC :" HorizontalAlignment="Left" VerticalAlignment="Top" FontFamily="Calibri" FontSize="24" FontWeight="Bold" Margin="10,204,0,0" Height="39" Width="47"/>
		<Label Content="Created By Raghaji Raul" HorizontalAlignment="Left" Height="100" Margin="0,283,0,-92" VerticalAlignment="Top" Width="143" Grid.ColumnSpan="2"/>

        <TextBox x:Name ="PCName" HorizontalAlignment="Left" TextWrapping="Wrap" Text="$PCName2" VerticalAlignment="Top" Width="308" Margin="82,10,0,0" FontSize="24" SpellCheck.IsEnabled="True" Height="39" Grid.Column="1"/>
		<TextBox x:Name ="UserName" HorizontalAlignment="Left" TextWrapping="Wrap" Text="$Username2" VerticalAlignment="Top" Width="308" Margin="82,49,0,0" FontSize="24" SpellCheck.IsEnabled="True" Height="39" Grid.Column="1"/>
		<TextBox x:Name ="UserID" HorizontalAlignment="Left" TextWrapping="Wrap" Text="$Userid2" VerticalAlignment="Top" Width="308" Margin="82,89,0,0" FontSize="24" SpellCheck.IsEnabled="True" Height="39" Grid.Column="1"/>
		<TextBox x:Name ="Subject" HorizontalAlignment="Left" TextWrapping="Wrap" Text="$subject2" VerticalAlignment="Top" Width="308" Margin="82,128,0,0" FontSize="24" SpellCheck.IsEnabled="True" Height="39" Grid.Column="1"/>
		<TextBox x:Name ="To" HorizontalAlignment="Left" TextWrapping="Wrap" Text="$to2" VerticalAlignment="Top" Width="308" Margin="82,167,0,0" FontSize="24" SpellCheck.IsEnabled="True" Height="39" Grid.Column="1"/>
		<TextBox x:Name ="Cc" HorizontalAlignment="Left" TextWrapping="Wrap" Text="$cc2" VerticalAlignment="Top" Width="308" Margin="82,205,0,0" FontSize="24" SpellCheck.IsEnabled="True" Height="39" Grid.Column="1"/>

		<Button x:Name ="OK" Content="OK" HorizontalAlignment="Left" Margin="0,249,0,0" VerticalAlignment="Top" Width="150" RenderTransformOrigin="0.5,0.5" Height="39" FontSize="24" FontFamily="Calibri" FontWeight="Bold" Grid.Column="1"/>
		<Button x:Name ="Cancel" Content="Cancel" Grid.Column="1" HorizontalAlignment="Left" Margin="218,249,0,0" VerticalAlignment="Top" Width="150" RenderTransformOrigin="0.5,0.5" Height="39" FontFamily="Calibri" FontSize="24" FontWeight="Bold"/>
		
	</Grid>
</Window>

"@

#Create Object for read xml file.
$reader = (New-Object System.Xml.XmlNodeReader $xaml)
#Create Object for GUI object from xaml.
$window = [windows.markup.XamlReader]::Load($reader)

#Find Text box from created XAML Window and store it in variable.
$PCName = $window.FindName("PCName")
$UserName = $window.FindName("UserName")
$UserID = $window.FindName("UserID")
$Subject = $window.FindName("Subject")
$To = $window.FindName("To")
$Cc = $window.FindName("Cc")

#Find Button from created XAML Window and store it in variable.
$OK = $window.FindName("OK")
$Cancel = $window.FindName("Cancel")

#Declear acton for botton.
$OK.Add_Click({
#Get value from text box and store it into variable.
$global:PCName1 = $PCName.Text;
$Global:UserID1 = $UserID.Text;
$Global:UserName1 = $UserName.Text;
$Global:subject1 = $subject.Text;
$Global:to1 = $To.Text;
$Global:cc1 = $Cc.Text;
#Close window
$window.Close();



# Outlook Connection
$Outlook = New-Object -ComObject Outlook.Application

# Send an Email from Outlook
$mail =$Outlook.CreateItem(0)
$mail.To = "$To1"
$mail.CC = "$Cc1"
$mail.Subject = "$Subject1"

[string]$html = Get-Content .\MailHTML.htm
$html = $html -replace "PCName:x", $PCName1
$html = $html -replace "UserName:x", $Username1
$html = $html -replace "UserID:x" , $UserID1

$mail.HTMLBody = "$html"

$mail.Send()
Start-Sleep -Seconds 10
$Outlook.close


})
$Cancel.Add_Click({$window.Close()})
#Pop window
$window.showDialog() |Out-Null

}