#-------------------------------------------------------------#
#----Initial Declarations-------------------------------------#
#-------------------------------------------------------------#
Set-ExecutionPolicy -ExecutionPolicy bypass process
Add-Type -AssemblyName PresentationCore, PresentationFramework, System.Windows.Forms

$Xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" Width="1530" Height="778" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="0,0,0,0">
  <Grid Margin="0,-1,0,1">
    <Border BorderBrush="Black" BorderThickness="1" Grid.Row="1" Grid.Column="0">
      <TabControl SelectedIndex="0">
        <TabItem Header="Employee List">
          <Grid Background="#FFE5E5E5">
            <Grid.RowDefinitions>
              <RowDefinition Height="55*"/>
              <RowDefinition Height="634*"/>
            </Grid.RowDefinitions>
            <Border BorderBrush="Black" BorderThickness="1" Grid.Row="1" Grid.Column="0">
              <DataGrid ItemsSource="{Binding EmployeeDataGrid}" Name="EmployeeDataGrid"/>
            </Border>
            <Grid Grid.Row="0" Grid.Column="0" Name="Tab1TopGrid">
              <Grid.ColumnDefinitions>
                <ColumnDefinition Width="135*"/>
                <ColumnDefinition Width="1369*"/>
              </Grid.ColumnDefinitions>
              <Border BorderBrush="Black" BorderThickness="1">
                <StackPanel>
                  <Button Content="Import" HorizontalAlignment="Left" VerticalAlignment="Top" Width="75" Margin="5,5,0,0" Name="ButtonEmployeeImport"/>
                  <Button Content="Save" HorizontalAlignment="Left" VerticalAlignment="Top" Width="75" Margin="5,5,0,0" Name="ButtonEmployeeSave"/>
                </StackPanel>
              </Border>
            </Grid>
          </Grid>
        </TabItem>
        <TabItem Header="SCCM Import">
          <Grid Background="#FFE5E5E5">
            <Grid>
              <Grid.RowDefinitions>
                <RowDefinition Height="55*"/>
                <RowDefinition Height="634*"/>
              </Grid.RowDefinitions>
              <Border BorderBrush="Black" BorderThickness="1" Grid.Row="0" Grid.Column="0">
                <StackPanel>
                  <Button Content="Import" HorizontalAlignment="Left" VerticalAlignment="Top" Width="75" Margin="5,5,0,0" Name="ButtonSCCMImport"/>
                  <Button Content="Save" HorizontalAlignment="Left" VerticalAlignment="Top" Width="75" Margin="5,5,0,0" Name="ButtomSCCMExport"/>
                </StackPanel>
              </Border>
              <Border BorderBrush="Black" BorderThickness="1" Grid.Row="1" Grid.Column="0">
                <DataGrid ItemsSource="{Binding SCCMDataGrid}" Name="SCCMDataGrid"/>
              </Border>
            </Grid>
          </Grid>
        </TabItem>
      </TabControl>
    </Border>
  </Grid>
</Window>

"@

#-------------------------------------------------------------#
#----Control Event Handlers-----------------------------------#
#-------------------------------------------------------------#


#region Logic
function Convert-EmployeeList {
    param(
        [Parameter(Mandatory=$true, ValueFromPipeline=$true)][array]$InputObject
    )
    process {
        $Knownas = $null
        if ($InputObject.'Known as' -ne ""){
        if ($InputObject.'Known as'.split(" ").count -gt 1){
            $KnownAs = $InputObject.'Known as'.split(" ")[0]
        }
        elseif ($InputObject.'Known as'.split(" ").count -eq 1){
            $KnownAs = $InputObject.'Known as'
        }
    }
    elseif ($InputObject.'Known as' -eq ""){
        $KnownAs = $InputObject.'First Name'}
        [PSCustomObject]@{
            DisplayName = if ($InputObject.'first name' -ne $KnownAs ) {
                "$KnownAs $($InputObject.'last name')"
            } else {
                "$($InputObject.'first name') $($InputObject.'last name')"
            }
            EmployeeName = $InputObject."Employee Name"
            Lastname = $InputObject.'Last Name'
            FirstName = $InputObject.'First Name'
            Status = $InputObject.Status
            PersonallArea = $InputObject.'PERSONNEL AREA'
            SubArea = $InputObject.'Personnel Sub Area Description'
            Knownas = $InputObject.'Known as'
            UserID = $InputObject.'User ID'
        }
    }
}

function Import-EmployeeList {
    $FileBrowser = [System.Windows.Forms.OpenFileDialog]::new()
    $FileBrowser.InitialDirectory = $env:HOMEPATH + "\Downloads"
    $FileBrowser.Filter = "XLSX Files (*.xlsx)|*.xlsx|All Files (*.*)|*.*"
    $FileBrowser.ShowDialog()
    $State.EmployeeDataGrid = Import-excel $FileBrowser.Filename | Convert-EmployeeList
}

function Export-EmployeeList {
    $FileBrowser = [System.Windows.Forms.SaveFileDialog]::new()
    $FileBrowser.InitialDirectory = $env:HOMEPATH + "\Downloads"
    $FileBrowser.Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"
    $FileBrowser.FileName = "EmployeeListExport_CarolaIsAwesome.csv"
    $FileBrowser.ShowDialog()
    $State.EmployeeDataGrid | Export-csv -notypeinformation -Path $FileBrowser.FileName
}

class SCCMReport {
    [string]$Serial_Number
    [string]$Name
    [string]$Full_User_Name             
    [string]$Primary_User
    [string]$Model
    [string]$Last_Seen_Online

    SCCMReport () {}

    SCCMReport ($Name, $Primary_User, $Full_User_Name, $Model, $Serial_Number, $Last_Seen_Online) {
        $this.Name = $Name
        $this.Primary_User = $Primary_User
        $this.Full_User_Name = $Full_User_Name
        $this.Model = $Model
        $this.Serial_Number = $Serial_Number
        $this.Last_Seen_Online = $Last_Seen_Online
    }

    static  [SCCMReport]SCCMOutput ($InputObject){
        return [SCCMReport]::new($InputObject.Name, $InputObject.Primary_User, $InputObject.Full_User_Name, $InputObject.Model, $InputObject.Serial_Number, $InputObject.Last_Seen_Online)
    }
}

function Out-SCCM {
    [CmdletBinding()]
    param(
        [parameter(ValueFromPipeline)][array]$InputObject,
        [string]$DownloadPath
    )
    begin {
        $Output = @()
    }
    process {
        $Output += [SCCMReport]::SCCMOutput($InputObject)
    }
    end {
        $Output | export-excel -Path $DownloadPath -WorksheetName "SCCMReport" -AutoFilter:$False  -FreezeTopRow -Calculate:$False -NoNumberConversion "Serial_Number"
    }
}

function Import-SCCM {
    [CmdletBinding()]
    param(
        [parameter(ValueFromPipeline)][array]$InputObject
    )
    process {
        [SCCMReport]::SCCMOutput($InputObject)
    }
}
function Import-SCCMList {
    $FileBrowser = [System.Windows.Forms.OpenFileDialog]::new()
    $FileBrowser.InitialDirectory = $env:HOMEPATH + "\Downloads"
    $FileBrowser.Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"
    $FileBrowser.ShowDialog()
    $State.SCCMDataGrid = Import-csv $FileBrowser.Filename | Import-SCCM
}
function Export-SCCM {
    $FileBrowser = [System.Windows.Forms.SaveFileDialog]::new()
    $FileBrowser.InitialDirectory = $env:HOMEPATH + "\Downloads"
    $FileBrowser.Filter = "XLSX Files (*.xlsx)|*.xlsx|All Files (*.*)|*.*"
    $FileBrowser.FileName = "SCCMExport_CarolaIsAwesome.xlsx"
    $FileBrowser.ShowDialog()
    $State.SCCMDataGrid | Out-SCCM -DownloadPath $FileBrowser.FileName
}
#endregion 


#-------------------------------------------------------------#
#----Script Execution-----------------------------------------#
#-------------------------------------------------------------#

$Window = [Windows.Markup.XamlReader]::Parse($Xaml)

[xml]$xml = $Xaml

$xml.SelectNodes("//*[@Name]") | ForEach-Object { Set-Variable -Name $_.Name -Value $Window.FindName($_.Name) }


$ButtonEmployeeImport.Add_Click({Import-EmployeeList $this $_})
$ButtonEmployeeSave.Add_Click({Export-EmployeeList $this $_})
$ButtonSCCMImport.Add_Click({Import-SCCMList $this $_})
$ButtomSCCMExport.Add_Click({Export-SCCM $this $_})

$State = [PSCustomObject]@{}


Function Set-Binding {
    Param($Target,$Property,$Index,$Name,$UpdateSourceTrigger)
 
    $Binding = New-Object System.Windows.Data.Binding
    $Binding.Path = "["+$Index+"]"
    $Binding.Mode = [System.Windows.Data.BindingMode]::TwoWay
    if($UpdateSourceTrigger -ne $null){$Binding.UpdateSourceTrigger = $UpdateSourceTrigger}


    [void]$Target.SetBinding($Property,$Binding)
}

function FillDataContext($props){

    For ($i=0; $i -lt $props.Length; $i++) {
   
   $prop = $props[$i]
   $DataContext.Add($DataObject."$prop")
   
    $getter = [scriptblock]::Create("Write-Output `$DataContext['$i'] -noenumerate")
    $setter = [scriptblock]::Create("param(`$val) return `$DataContext['$i']=`$val")
    $State | Add-Member -Name $prop -MemberType ScriptProperty -Value  $getter -SecondValue $setter
               
       }
   }



$DataObject =  ConvertFrom-Json @"

{
    "EmployeeDataGrid" : "",
    "SCCMDataGrid" : ""
}

"@

$DataContext = New-Object System.Collections.ObjectModel.ObservableCollection[Object]
FillDataContext @("EmployeeDataGrid","SCCMDataGrid") 

$Window.DataContext = $DataContext
Set-Binding -Target $EmployeeDataGrid -Property $([System.Windows.Controls.DataGrid]::ItemsSourceProperty) -Index 0 -Name "EmployeeDataGrid"  
Set-Binding -Target $SCCMDataGrid -Property $([System.Windows.Controls.DataGrid]::ItemsSourceProperty) -Index 1 -Name "SCCMDataGrid"  
$Window.ShowDialog()