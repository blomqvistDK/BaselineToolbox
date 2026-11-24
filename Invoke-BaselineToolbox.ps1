#Requires -Version 5.1
<#PSScriptInfo

.VERSION 1.0
.AUTHOR Mikael Blomqvist
.RELEASENOTES
v1.0 - Initial release: combined CIS Baseline Toolbox with current tab set
       Tab 1: PDF to Markdown Converter
       Tab 2: Export Change History
       Tab 3: Markdown to HTML
       Tab 4: Baseline Merger
       Tab 5: Intune Policy Creator
       Tab 6: Intune Baseline Profile Report
       Tab 7: Baseline Assessment

.PRIVATEDATA
#>

<#
.SYNOPSIS
    CIS Baseline Toolbox — combined WPF GUI for CIS benchmark workflows.

.DESCRIPTION
    A WPF-based GUI combining the main CIS benchmark workflow tools in one tabbed window:
      Tab 1 — PDF to Markdown: Convert a CIS benchmark PDF into formatted Markdown output
      Tab 2 — Export Change History: Compare benchmark PDFs and export change history HTML
      Tab 3 — Markdown to HTML: Convert Markdown benchmark content into an interactive HTML report
      Tab 4 — Baseline Merger: Merge new/old baselines and generate HTML comparison reports
      Tab 5 — Intune Policy Creator: Create and export Settings Catalog policies from Markdown
    Tab 6 — Intune Policy Report: Inventory policies or compare settings with policy backups
      Tab 7 — Baseline Assessment: Assess imported baseline backups against tenant policy data and generate HTML findings

.EXAMPLE
    .\Invoke-BaselineToolbox.ps1

.NOTES
    Version:        1.0
    Author:         Mikael Blomqvist
    Creation Date:  2026-04-30

    Required Microsoft Graph Scopes:
      - DeviceManagementConfiguration.Read.All    (Baseline Report, Baseline Assessment)
      - DeviceManagementConfiguration.ReadWrite.All (Intune Policy Creator)
      - Organization.Read.All                      (Tenant name in HTML reports)

        Feature processing is implemented in sibling parameter-driven scripts.
        See README.md for standalone examples and dependencies.
#>

[CmdletBinding()]
param(
    [string]$ResolverTenantId = $env:BASELINE_TOOLBOX_RESOLVER_TENANTID
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Import-Module (Join-Path $PSScriptRoot 'Shared\BaselineToolkit.Common.psm1') -DisableNameChecking
. (Join-Path $PSScriptRoot 'Features\Export-BaselineAssessmentWorkbook.ps1')
. (Join-Path $PSScriptRoot 'Features\Invoke-BaselineMerger.ps1')
. (Join-Path $PSScriptRoot 'Features\Invoke-BaselineAssessment.ps1')
. (Join-Path $PSScriptRoot 'Features\Invoke-IntuneBaselineReport.ps1')

Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase
Add-Type -AssemblyName System.Windows.Forms

$script:ToolVersion = '1.0'
$script:thisScriptPath = $PSCommandPath
$script:ResolverTenantId = if ($ResolverTenantId -and $ResolverTenantId.Trim()) { $ResolverTenantId.Trim() } else { $null }
Set-BaselineResolverTenant -TenantId $script:ResolverTenantId

# ══════════════════════════════════════════════════════════════════════════════
#  SHARED FUNCTIONS
# ══════════════════════════════════════════════════════════════════════════════

function Test-Prerequisites {
    [CmdletBinding()]
    param()
    $issues = @()
    $pythonExe = $null
    $pythonKind = $null
    $pythonCmd = Get-Command python -ErrorAction SilentlyContinue
    if ($pythonCmd) { $pythonExe = $pythonCmd.Source; $pythonKind = 'python' }
    else {
        $pyCmd = Get-Command py -ErrorAction SilentlyContinue
        if ($pyCmd) { $pythonExe = $pyCmd.Source; $pythonKind = 'py' }
    }
    if (-not $pythonExe) {
        $issues += 'Python 3 was not found in PATH.'
        return [pscustomobject]@{ IsOk = $false; Issues = $issues }
    }
    try {
        $versionOutput = if ($pythonKind -eq 'python') { & $pythonExe --version 2>&1 } else { & $pythonExe -3 --version 2>&1 }
        if ($versionOutput -notmatch 'Python\s+3(\.|$)') { $issues += "Python 3 is required, but detected: $versionOutput" }
    } catch { $issues += "Failed to validate Python version: $($_.Exception.Message)" }
    try {
        if ($pythonKind -eq 'python') { & $pythonExe -c "import PyPDF2" *> $null } else { & $pythonExe -3 -c "import PyPDF2" *> $null }
        if ($LASTEXITCODE -ne 0) { $issues += 'PyPDF2 is not importable for Python 3.' }
    } catch { $issues += "Failed to validate PyPDF2 import: $($_.Exception.Message)" }
    return [pscustomobject]@{ IsOk = ($issues.Count -eq 0); Issues = $issues }
}

#region Markdown Parsing Helpers (for Merger)

#endregion

#region Merge Logic

#endregion

#region HTML Report Generation (for Merger)

#endregion

#region RecItem class for DataGrid checkbox binding (Policy Creator Selection tab)
$_recItemType = 'RecItem' -as [type]
if (-not $_recItemType -or -not $_recItemType.GetProperty('IsManual') -or -not $_recItemType.GetProperty('IsAutoCreationUnsupported') -or -not $_recItemType.GetProperty('IsCustomerInfoRequired')) {
    if ($_recItemType) {
        [System.Windows.MessageBox]::Show(
            "A previous version of RecItem is loaded in this session.`nPlease close and re-open the script in a new PowerShell window.",
            'Restart Required', 'OK', 'Warning')
        return
    }
    Add-Type -Language CSharp @'
using System.ComponentModel;
public class RecItem : INotifyPropertyChanged
{
    private bool _selected;
    public bool Selected
    {
        get { return _selected; }
        set { _selected = value; OnPropertyChanged("Selected"); }
    }
    public string CISRef        { get; set; }
    public string SettingName   { get; set; }
    public string SettingValue  { get; set; }
    public string PolicyName    { get; set; }
    public string Profile       { get; set; }
    public string Recommendation { get; set; }
    public string Notes         { get; set; }
    public bool   IsManual      { get; set; }
    public bool   IsAutoCreationUnsupported { get; set; }
    public bool   IsCustomerInfoRequired { get; set; }

    public event PropertyChangedEventHandler PropertyChanged;
    protected void OnPropertyChanged(string name)
    {
        if (PropertyChanged != null)
            PropertyChanged(this, new PropertyChangedEventArgs(name));
    }
}
'@
}
#endregion

#region Parse-MdFile function (for Policy Creator Selection tab)

#endregion

# ══════════════════════════════════════════════════════════════════════════════
#  WPF XAML DEFINITION
# ══════════════════════════════════════════════════════════════════════════════

[xml]$xaml = @"
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="CIS Baseline Toolbox" Height="880" Width="1100"
    MinHeight="700" MinWidth="900"
    ResizeMode="CanResize" WindowStartupLocation="CenterScreen"
    Background="{DynamicResource BgBrush}">

    <Window.Resources>
        <SolidColorBrush x:Key="BgBrush"           Color="#0D151D"/>
        <SolidColorBrush x:Key="AccentBrush"       Color="#36c18f"/>
        <SolidColorBrush x:Key="AccentHover"       Color="#2ea87c"/>
        <SolidColorBrush x:Key="SecondaryBtnBrush" Color="#111A24"/>
        <SolidColorBrush x:Key="SecondaryBtnHover" Color="#1C2A38"/>
        <SolidColorBrush x:Key="DangerBrush"       Color="#dc2626"/>
        <SolidColorBrush x:Key="DangerHoverBrush"  Color="#b91c1c"/>
        <SolidColorBrush x:Key="SuccessBrush"      Color="#16a34a"/>
        <SolidColorBrush x:Key="AttentionBrush"     Color="#d9a441"/>
        <SolidColorBrush x:Key="TextBrush"         Color="#E7EDF3"/>
        <SolidColorBrush x:Key="ProfileRowTextBrush" Color="#D8E2EA"/>
        <SolidColorBrush x:Key="SubTextBrush"      Color="#93A4B5"/>
        <SolidColorBrush x:Key="BorderBrush2"      Color="#263849"/>
        <SolidColorBrush x:Key="InputBgBrush"      Color="#111922"/>
        <SolidColorBrush x:Key="SurfaceBrush"      Color="#16212D"/>

        <Style TargetType="TextBlock">
            <Setter Property="Foreground" Value="{StaticResource TextBrush}"/>
        </Style>

        <Style x:Key="SectionTitle" TargetType="TextBlock">
            <Setter Property="Foreground" Value="{StaticResource TextBrush}"/>
            <Setter Property="FontSize" Value="14"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Margin" Value="0,14,0,7"/>
        </Style>

        <Style x:Key="NavCategory" TargetType="TextBlock">
            <Setter Property="Foreground" Value="{StaticResource SubTextBrush}"/>
            <Setter Property="FontSize" Value="10"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Margin" Value="8,16,0,8"/>
        </Style>

        <Style x:Key="ActivityLog" TargetType="TextBox">
            <Setter Property="Background" Value="{StaticResource InputBgBrush}"/>
            <Setter Property="Foreground" Value="{StaticResource TextBrush}"/>
            <Setter Property="FontFamily" Value="Cascadia Mono, Consolas"/>
            <Setter Property="FontSize" Value="11"/>
            <Setter Property="Padding" Value="8"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="TextWrapping" Value="Wrap"/>
            <Setter Property="VerticalScrollBarVisibility" Value="Auto"/>
        </Style>

        <Style x:Key="ActivityHeading" TargetType="TextBlock">
            <Setter Property="Foreground" Value="{StaticResource TextBrush}"/>
            <Setter Property="FontSize" Value="13"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
        </Style>

        <Style x:Key="ActivitySummary" TargetType="TextBlock">
            <Setter Property="Foreground" Value="{StaticResource SubTextBrush}"/>
            <Setter Property="FontSize" Value="11"/>
            <Setter Property="TextTrimming" Value="CharacterEllipsis"/>
            <Setter Property="VerticalAlignment" Value="Center"/>
            <Setter Property="Margin" Value="12,0,0,0"/>
        </Style>

        <Style TargetType="CheckBox">
            <Setter Property="Foreground" Value="{StaticResource TextBrush}"/>
        </Style>

        <Style TargetType="RadioButton">
            <Setter Property="Foreground" Value="{StaticResource TextBrush}"/>
        </Style>

        <Style TargetType="TextBox">
            <Setter Property="Foreground" Value="{StaticResource TextBrush}"/>
            <Setter Property="Background" Value="{StaticResource InputBgBrush}"/>
            <Setter Property="CaretBrush" Value="{StaticResource TextBrush}"/>
            <Setter Property="BorderBrush" Value="{StaticResource BorderBrush2}"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="MinHeight" Value="34"/>
            <Setter Property="FontSize" Value="12"/>
            <Style.Triggers>
                <Trigger Property="IsKeyboardFocused" Value="True">
                    <Setter Property="BorderBrush" Value="{StaticResource AccentBrush}"/>
                </Trigger>
                <Trigger Property="IsEnabled" Value="False">
                    <Setter Property="Opacity" Value="0.65"/>
                </Trigger>
            </Style.Triggers>
        </Style>

        <Style TargetType="ComboBox">
            <Setter Property="Foreground" Value="{StaticResource TextBrush}"/>
            <Setter Property="Background" Value="{StaticResource InputBgBrush}"/>
            <Setter Property="BorderBrush" Value="{StaticResource BorderBrush2}"/>
            <Setter Property="MinHeight" Value="34"/>
            <Setter Property="FontSize" Value="12"/>
        </Style>

        <Style x:Key="PolicyCreatorHighContrastComboBox" TargetType="ComboBox">
            <Setter Property="Foreground" Value="#E7EDF3"/>
            <Setter Property="Background" Value="#111922"/>
            <Setter Property="BorderBrush" Value="#263849"/>
            <Setter Property="MinHeight" Value="34"/>
            <Setter Property="FontSize" Value="12"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="ComboBox">
                        <Grid>
                            <Border x:Name="SelectionBorder" Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="1" CornerRadius="2"/>
                            <Grid>
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="*"/>
                                    <ColumnDefinition Width="28"/>
                                </Grid.ColumnDefinitions>
                                <ContentPresenter Grid.Column="0" IsHitTestVisible="False" Content="{TemplateBinding SelectionBoxItem}" ContentTemplate="{TemplateBinding SelectionBoxItemTemplate}" ContentTemplateSelector="{TemplateBinding ItemTemplateSelector}" TextElement.Foreground="{TemplateBinding Foreground}" Margin="8,0,4,0" VerticalAlignment="Center" HorizontalAlignment="Left"/>
                                <ToggleButton Grid.Column="1" Focusable="False" Background="Transparent" BorderThickness="0" Foreground="{TemplateBinding Foreground}" IsChecked="{Binding IsDropDownOpen, Mode=TwoWay, RelativeSource={RelativeSource TemplatedParent}}" ClickMode="Press">
                                    <ToggleButton.Template>
                                        <ControlTemplate TargetType="ToggleButton">
                                            <Border Background="{TemplateBinding Background}">
                                                <Path Data="M 0 0 L 4 4 L 8 0 Z" Fill="{TemplateBinding Foreground}" HorizontalAlignment="Center" VerticalAlignment="Center"/>
                                            </Border>
                                        </ControlTemplate>
                                    </ToggleButton.Template>
                                </ToggleButton>
                            </Grid>
                            <Popup x:Name="Popup" Placement="Bottom" PlacementTarget="{Binding RelativeSource={RelativeSource TemplatedParent}}" IsOpen="{Binding IsDropDownOpen, Mode=TwoWay, RelativeSource={RelativeSource TemplatedParent}}" AllowsTransparency="True" Focusable="False" StaysOpen="False" PopupAnimation="Fade">
                                <Border Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="1" MinWidth="{TemplateBinding ActualWidth}">
                                    <ScrollViewer MaxHeight="{TemplateBinding MaxDropDownHeight}" CanContentScroll="True">
                                        <ItemsPresenter KeyboardNavigation.DirectionalNavigation="Contained"/>
                                    </ScrollViewer>
                                </Border>
                            </Popup>
                        </Grid>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsKeyboardFocusWithin" Value="True">
                                <Setter TargetName="SelectionBorder" Property="BorderBrush" Value="#36C18F"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
            <Setter Property="ItemContainerStyle">
                <Setter.Value>
                    <Style TargetType="ComboBoxItem">
                        <Setter Property="Foreground" Value="#E7EDF3"/>
                        <Setter Property="Background" Value="#111922"/>
                        <Setter Property="Padding" Value="8,5"/>
                        <Setter Property="Template">
                            <Setter.Value>
                                <ControlTemplate TargetType="ComboBoxItem">
                                    <Border x:Name="ItemBorder" Background="{TemplateBinding Background}" Padding="{TemplateBinding Padding}">
                                        <ContentPresenter VerticalAlignment="Center" RecognizesAccessKey="True" TextElement.Foreground="{TemplateBinding Foreground}"/>
                                    </Border>
                                    <ControlTemplate.Triggers>
                                        <Trigger Property="IsHighlighted" Value="True">
                                            <Setter TargetName="ItemBorder" Property="Background" Value="#1C2A38"/>
                                            <Setter Property="Foreground" Value="#E7EDF3"/>
                                        </Trigger>
                                        <Trigger Property="IsSelected" Value="True">
                                            <Setter TargetName="ItemBorder" Property="Background" Value="#36C18F"/>
                                            <Setter Property="Foreground" Value="#0D151D"/>
                                        </Trigger>
                                    </ControlTemplate.Triggers>
                                </ControlTemplate>
                            </Setter.Value>
                        </Setter>
                    </Style>
                </Setter.Value>
            </Setter>
        </Style>

        <Style TargetType="ComboBoxItem">
            <Setter Property="Background" Value="{StaticResource InputBgBrush}"/>
            <Setter Property="Foreground" Value="{StaticResource TextBrush}"/>
            <Setter Property="Padding" Value="8,5"/>
            <Style.Triggers>
                <Trigger Property="IsHighlighted" Value="True"><Setter Property="Background" Value="#1C2A38"/></Trigger>
                <Trigger Property="IsSelected" Value="True"><Setter Property="Background" Value="#1D352F"/></Trigger>
            </Style.Triggers>
        </Style>

        <Style TargetType="DataGrid">
            <Setter Property="Background" Value="#111922"/>
            <Setter Property="Foreground" Value="{StaticResource TextBrush}"/>
            <Setter Property="RowBackground" Value="#111922"/>
            <Setter Property="AlternatingRowBackground" Value="#151F2A"/>
            <Setter Property="AlternationCount" Value="2"/>
            <Setter Property="BorderBrush" Value="{StaticResource BorderBrush2}"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="HorizontalGridLinesBrush" Value="{StaticResource BorderBrush2}"/>
            <Setter Property="VerticalGridLinesBrush" Value="{StaticResource BorderBrush2}"/>
        </Style>

        <Style TargetType="DataGridColumnHeader">
            <Setter Property="Background" Value="#1B2734"/>
            <Setter Property="Foreground" Value="{StaticResource TextBrush}"/>
            <Setter Property="BorderBrush" Value="{StaticResource BorderBrush2}"/>
            <Setter Property="BorderThickness" Value="0,0,0,1"/>
            <Setter Property="Padding" Value="8,7"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
        </Style>

        <Style TargetType="DataGridRow">
            <Setter Property="Foreground" Value="{StaticResource TextBrush}"/>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True"><Setter Property="Background" Value="#1C2A38"/></Trigger>
                <Trigger Property="IsSelected" Value="True"><Setter Property="Background" Value="#1D352F"/></Trigger>
            </Style.Triggers>
        </Style>

        <Style TargetType="DataGridCell">
            <Setter Property="Foreground" Value="{StaticResource TextBrush}"/>
            <Setter Property="Background" Value="Transparent"/>
            <Setter Property="BorderBrush" Value="Transparent"/>
            <Setter Property="Padding" Value="6,4"/>
            <Style.Triggers>
                <Trigger Property="IsSelected" Value="True">
                    <Setter Property="Background" Value="#1D352F"/>
                    <Setter Property="Foreground" Value="{StaticResource TextBrush}"/>
                </Trigger>
                <Trigger Property="IsKeyboardFocusWithin" Value="True"><Setter Property="BorderBrush" Value="{StaticResource AccentBrush}"/></Trigger>
            </Style.Triggers>
        </Style>

        <Style TargetType="ProgressBar">
            <Setter Property="Foreground" Value="{StaticResource AccentBrush}"/>
            <Setter Property="Background" Value="{StaticResource InputBgBrush}"/>
            <Setter Property="BorderBrush" Value="{StaticResource BorderBrush2}"/>
            <Setter Property="Height" Value="8"/>
        </Style>

        <Style TargetType="TabControl">
            <Setter Property="Background" Value="{StaticResource BgBrush}"/>
            <Setter Property="BorderBrush" Value="{StaticResource BorderBrush2}"/>
        </Style>

        <Style x:Key="PageHostTabControl" TargetType="TabControl">
            <Setter Property="Background" Value="Transparent"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="TabControl">
                        <Grid ClipToBounds="True">
                            <ContentPresenter ContentSource="SelectedContent" Margin="{TemplateBinding Padding}"/>
                        </Grid>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <Style x:Key="NavItem" TargetType="RadioButton">
            <Setter Property="Foreground" Value="{StaticResource SubTextBrush}"/>
            <Setter Property="FontSize" Value="12"/>
            <Setter Property="Padding" Value="10,8"/>
            <Setter Property="Margin" Value="0,1"/>
            <Setter Property="BorderThickness" Value="3,0,0,0"/>
            <Setter Property="BorderBrush" Value="Transparent"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="RadioButton">
                        <Border x:Name="NavBorder" Background="Transparent" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}" CornerRadius="3" Padding="{TemplateBinding Padding}">
                            <ContentPresenter VerticalAlignment="Center" RecognizesAccessKey="True"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="NavBorder" Property="Background" Value="#1C2A38"/>
                                <Setter Property="Foreground" Value="{StaticResource TextBrush}"/>
                            </Trigger>
                            <Trigger Property="IsChecked" Value="True">
                                <Setter TargetName="NavBorder" Property="Background" Value="#1D352F"/>
                                <Setter Property="BorderBrush" Value="{StaticResource AccentBrush}"/>
                                <Setter Property="Foreground" Value="{StaticResource AccentBrush}"/>
                                <Setter Property="FontWeight" Value="SemiBold"/>
                            </Trigger>
                            <Trigger Property="IsKeyboardFocused" Value="True">
                                <Setter TargetName="NavBorder" Property="BorderBrush" Value="{StaticResource AccentBrush}"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter Property="Opacity" Value="0.58"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <Style TargetType="TabItem">
            <Setter Property="Foreground" Value="{StaticResource TextBrush}"/>
            <Setter Property="Background" Value="{StaticResource SurfaceBrush}"/>
            <Setter Property="BorderBrush" Value="{StaticResource BorderBrush2}"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="TabItem">
                        <Border x:Name="TabBorder"
                                Background="{TemplateBinding Background}"
                                BorderBrush="{TemplateBinding BorderBrush}"
                                BorderThickness="1"
                                CornerRadius="6,6,0,0"
                                Margin="0,0,2,0"
                                Padding="10,5">
                            <ContentPresenter x:Name="ContentSite"
                                              ContentSource="Header"
                                              HorizontalAlignment="Center"
                                              VerticalAlignment="Center"
                                              RecognizesAccessKey="True"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsSelected" Value="True">
                                <Setter TargetName="TabBorder" Property="Background" Value="{StaticResource AccentBrush}"/>
                                <Setter TargetName="TabBorder" Property="BorderBrush" Value="{StaticResource AccentBrush}"/>
                                <Setter Property="Foreground" Value="#0e151d"/>
                            </Trigger>
                            <MultiTrigger>
                                <MultiTrigger.Conditions>
                                    <Condition Property="IsMouseOver" Value="True"/>
                                    <Condition Property="IsSelected" Value="False"/>
                                </MultiTrigger.Conditions>
                                <Setter TargetName="TabBorder" Property="Background" Value="#243445"/>
                                <Setter TargetName="TabBorder" Property="BorderBrush" Value="#32475F"/>
                            </MultiTrigger>
                            <MultiTrigger>
                                <MultiTrigger.Conditions>
                                    <Condition Property="IsMouseOver" Value="True"/>
                                    <Condition Property="IsSelected" Value="True"/>
                                </MultiTrigger.Conditions>
                                <Setter TargetName="TabBorder" Property="Background" Value="{StaticResource AccentHover}"/>
                                <Setter TargetName="TabBorder" Property="BorderBrush" Value="{StaticResource AccentHover}"/>
                                <Setter Property="Foreground" Value="#0e151d"/>
                            </MultiTrigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter Property="Opacity" Value="0.55"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
            <Style.Triggers>
                <Trigger Property="IsSelected" Value="True">
                    <Setter Property="Background" Value="{StaticResource AccentBrush}"/>
                    <Setter Property="Foreground" Value="#0e151d"/>
                    <Setter Property="BorderBrush" Value="{StaticResource AccentBrush}"/>
                </Trigger>
            </Style.Triggers>
        </Style>

        <!-- Primary Button -->
        <Style x:Key="PrimaryBtn" TargetType="Button">
            <Setter Property="Background"      Value="{StaticResource AccentBrush}"/>
            <Setter Property="Foreground"      Value="#0D151D"/>
            <Setter Property="FontSize"        Value="12"/>
            <Setter Property="FontWeight"      Value="SemiBold"/>
            <Setter Property="Padding"         Value="12,5"/>
            <Setter Property="MinHeight"       Value="34"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Cursor"          Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border Background="{TemplateBinding Background}" CornerRadius="4" Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True"><Setter Property="Background" Value="{StaticResource AccentHover}"/></Trigger>
                            <Trigger Property="IsPressed" Value="True"><Setter Property="Background" Value="#238A68"/></Trigger>
                            <Trigger Property="IsEnabled" Value="False"><Setter Property="Opacity" Value="0.58"/></Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- Secondary Button -->
        <Style x:Key="SecondaryBtn" TargetType="Button">
            <Setter Property="Background" Value="{StaticResource SecondaryBtnBrush}"/>
            <Setter Property="Foreground" Value="{StaticResource TextBrush}"/>
            <Setter Property="FontSize" Value="12"/>
            <Setter Property="Padding" Value="12,5"/>
            <Setter Property="MinHeight" Value="34"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="BorderBrush" Value="{StaticResource BorderBrush2}"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}" CornerRadius="4" Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True"><Setter Property="Background" Value="{StaticResource SecondaryBtnHover}"/></Trigger>
                            <Trigger Property="IsPressed" Value="True"><Setter Property="Background" Value="#223343"/></Trigger>
                            <Trigger Property="IsEnabled" Value="False"><Setter Property="Opacity" Value="0.58"/></Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <Style x:Key="TertiaryBtn" TargetType="Button" BasedOn="{StaticResource SecondaryBtn}">
            <Setter Property="Background" Value="Transparent"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Padding" Value="8,4"/>
            <Setter Property="MinHeight" Value="28"/>
        </Style>

        <!-- Danger Button -->
        <Style x:Key="DangerBtn" TargetType="Button">
            <Setter Property="Background" Value="{StaticResource DangerBrush}"/>
            <Setter Property="Foreground" Value="#E7EDF3"/>
            <Setter Property="FontSize" Value="12"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Padding" Value="12,5"/>
            <Setter Property="MinHeight" Value="34"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border Background="{TemplateBinding Background}" CornerRadius="4" Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True"><Setter Property="Background" Value="{StaticResource DangerHoverBrush}"/></Trigger>
                            <Trigger Property="IsEnabled" Value="False"><Setter Property="Opacity" Value="0.58"/></Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
    </Window.Resources>

    <Grid Margin="0">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
        </Grid.RowDefinitions>
        <Border Grid.Row="0" Background="{DynamicResource SecondaryBtnBrush}" BorderBrush="{DynamicResource BorderBrush2}" BorderThickness="0,0,0,1" Padding="16,7">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <StackPanel Grid.Column="0" VerticalAlignment="Center">
                    <TextBlock Text="CIS Baseline Toolbox" FontSize="20" FontWeight="SemiBold" Foreground="{DynamicResource TextBrush}"/>
                    <TextBlock Text="Convert PDFs, merge baselines, and create Intune policies." FontSize="12" Foreground="{DynamicResource SubTextBrush}" Margin="0,2,0,0"/>
                </StackPanel>
                <TextBlock x:Name="TxtVersion" Grid.Column="1" Text="v1.0" VerticalAlignment="Center" FontSize="11" Foreground="{DynamicResource SubTextBrush}" Margin="16,0,4,0"/>
            </Grid>
        </Border>
        <Grid Grid.Row="1" Margin="8,0,8,8">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="184"/>
                <ColumnDefinition Width="8"/>
                <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>
            <Border Grid.Column="0" Background="{DynamicResource SecondaryBtnBrush}" Padding="0">
                <Grid>
                    <Grid.RowDefinitions>
                        <RowDefinition Height="*"/>
                        <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>
                    <ScrollViewer Grid.Row="0" VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Disabled">
                        <StackPanel Margin="8,8">
                            <TextBlock Text="CONVERSION" Style="{StaticResource NavCategory}" Margin="8,4,0,8"/>
                            <RadioButton x:Name="navPdfToMarkdown" GroupName="ToolNavigation" Tag="0" Content="PDF to Markdown" Style="{StaticResource NavItem}" IsChecked="True"/>
                            <RadioButton x:Name="navMarkdownToHtml" GroupName="ToolNavigation" Tag="2" Content="Markdown to HTML" Style="{StaticResource NavItem}"/>
                            <TextBlock Text="BASELINES" Style="{StaticResource NavCategory}"/>
                            <RadioButton x:Name="navBaselineMerger" GroupName="ToolNavigation" Tag="3" Content="Baseline Merger" Style="{StaticResource NavItem}"/>
                            <RadioButton x:Name="navBaselineAssessment" GroupName="ToolNavigation" Tag="6" Content="Baseline Assessment" Style="{StaticResource NavItem}"/>
                            <RadioButton x:Name="navAssessmentWorkbook" GroupName="ToolNavigation" Tag="7" Content="Baseline Workbook" Style="{StaticResource NavItem}"/>
                            <TextBlock Text="INTUNE" Style="{StaticResource NavCategory}"/>
                            <RadioButton x:Name="navIntunePolicyCreator" GroupName="ToolNavigation" Tag="4" Content="Intune Policy Creator" Style="{StaticResource NavItem}"/>
                            <RadioButton x:Name="navIntuneBaselineReport" GroupName="ToolNavigation" Tag="5" Content="Intune Policy Report" Style="{StaticResource NavItem}"/>
                            <TextBlock Text="TOOLS" Style="{StaticResource NavCategory}"/>
                            <RadioButton x:Name="navExportChangeHistory" GroupName="ToolNavigation" Tag="1" Content="Export Change History" Style="{StaticResource NavItem}"/>
                        </StackPanel>
                    </ScrollViewer>
                    <Border Grid.Row="1" BorderBrush="{DynamicResource BorderBrush2}" BorderThickness="0,1,0,0" Padding="8,8">
                        <StackPanel>
                            <Border x:Name="policyGraphPanel" Background="{DynamicResource SurfaceBrush}" BorderBrush="{DynamicResource BorderBrush2}" BorderThickness="1" CornerRadius="4" Padding="8,7" Margin="0,0,0,6" Visibility="Collapsed">
                                <StackPanel>
                                    <TextBlock Text="Microsoft Graph" FontSize="10" FontWeight="SemiBold" Foreground="{DynamicResource SubTextBrush}" Margin="0,0,0,5"/>
                                    <StackPanel Orientation="Horizontal" Margin="0,0,0,7">
                                        <Ellipse x:Name="policyGraphStatus" Width="8" Height="8" Fill="{DynamicResource DangerBrush}" VerticalAlignment="Center" Margin="0,0,6,0"/>
                                        <TextBlock x:Name="txtTenantInfo" Text="Not connected" FontSize="11" Foreground="{DynamicResource TextBrush}" VerticalAlignment="Center"/>
                                    </StackPanel>
                                    <TextBlock x:Name="txtPolicyTenantName" Text="" FontSize="11" FontWeight="SemiBold" Margin="0,0,0,5" Foreground="{DynamicResource TextBrush}" TextTrimming="CharacterEllipsis" ToolTip=""/>
                                    <Button x:Name="btnGraphConnect" Content="Connect" MinWidth="76" Height="28" HorizontalAlignment="Left" Style="{StaticResource SecondaryBtn}"/>
                                </StackPanel>
                            </Border>
                            <Border x:Name="profileReportGraphPanel" Background="{DynamicResource SurfaceBrush}" BorderBrush="{DynamicResource BorderBrush2}" BorderThickness="1" CornerRadius="4" Padding="8,7" Margin="0,0,0,6" Visibility="Collapsed">
                                <StackPanel>
                                    <TextBlock Text="Microsoft Graph" FontSize="10" FontWeight="SemiBold" Foreground="{DynamicResource SubTextBrush}" Margin="0,0,0,5"/>
                                    <StackPanel Orientation="Horizontal" Margin="0,0,0,7">
                                        <Ellipse x:Name="profileReportGraphStatus" Width="8" Height="8" Fill="{DynamicResource DangerBrush}" VerticalAlignment="Center" Margin="0,0,6,0"/>
                                        <TextBlock x:Name="txtProfileReportTenantInfo" Text="Not connected" FontSize="11" Foreground="{DynamicResource TextBrush}" VerticalAlignment="Center"/>
                                    </StackPanel>
                                    <TextBlock x:Name="txtProfileReportTenantName" Text="" FontSize="11" FontWeight="SemiBold" Margin="0,0,0,5" Foreground="{DynamicResource TextBrush}" TextTrimming="CharacterEllipsis" ToolTip=""/>
                                    <Button x:Name="btnProfileReportGraphConnect" Content="Connect" MinWidth="76" Height="28" HorizontalAlignment="Left" Style="{StaticResource SecondaryBtn}"/>
                                </StackPanel>
                            </Border>
                            <Border x:Name="assessmentGraphPanel" Background="{DynamicResource SurfaceBrush}" BorderBrush="{DynamicResource BorderBrush2}" BorderThickness="1" CornerRadius="4" Padding="8,7" Margin="0" Visibility="Collapsed">
                                <StackPanel>
                                    <TextBlock Text="Microsoft Graph" FontSize="10" FontWeight="SemiBold" Foreground="{DynamicResource SubTextBrush}" Margin="0,0,0,5"/>
                                    <StackPanel Orientation="Horizontal" Margin="0,0,0,7">
                                        <Ellipse x:Name="assessmentGraphStatus" Width="8" Height="8" Fill="{DynamicResource DangerBrush}" VerticalAlignment="Center" Margin="0,0,6,0"/>
                                        <TextBlock x:Name="txtAssessmentTenantInfo" Text="Not connected" FontSize="11" Foreground="{DynamicResource TextBrush}" VerticalAlignment="Center"/>
                                    </StackPanel>
                                    <TextBlock x:Name="txtAssessmentTenantName" Text="" FontSize="11" FontWeight="SemiBold" Margin="0,0,0,5" Foreground="{DynamicResource TextBrush}" TextTrimming="CharacterEllipsis" ToolTip=""/>
                                    <Button x:Name="btnAssessmentGraphConnect" Content="Connect" MinWidth="76" Height="28" HorizontalAlignment="Left" Style="{StaticResource SecondaryBtn}"/>
                                </StackPanel>
                            </Border>
                        </StackPanel>
                    </Border>
                </Grid>
            </Border>
            <Grid Grid.Column="2">
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="*"/>
                </Grid.RowDefinitions>
                <Grid Grid.Row="0" Margin="4,0,4,10">
                    <StackPanel VerticalAlignment="Center">
                        <TextBlock x:Name="txtCurrentPageTitle" FontSize="20" FontWeight="SemiBold"/>
                        <TextBlock x:Name="txtCurrentPageDescription" FontSize="12" Foreground="{DynamicResource SubTextBrush}" Margin="0,3,0,0"/>
                    </StackPanel>
                </Grid>
                <TabControl x:Name="MainTabs" Grid.Row="1" Padding="0" Style="{StaticResource PageHostTabControl}">
                <!-- Tab 1: PDF to Markdown -->
                <TabItem Header="  PDF to Markdown  ">
                    <Grid Margin="12">
                        <Grid.RowDefinitions>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="Auto"/>
                        </Grid.RowDefinitions>

                    <TextBlock Grid.Row="0" Text="Source" Style="{StaticResource SectionTitle}"/>
                    <TextBlock Grid.Row="1" Text="CIS Benchmark PDF *" FontWeight="SemiBold" Margin="0,0,0,4"/>
                    <Grid Grid.Row="2" Margin="0,0,0,4">
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="Auto"/>
                        </Grid.ColumnDefinitions>
                        <TextBox x:Name="txtPdfPath" Grid.Column="0" IsReadOnly="True" Padding="6,4" Background="{DynamicResource InputBgBrush}" BorderBrush="{DynamicResource BorderBrush2}"/>
                        <Button x:Name="btnBrowsePdf" Grid.Column="1" Content="Browse..." Width="90" Margin="8,0,0,0" Style="{StaticResource SecondaryBtn}"/>
                    </Grid>

                    <TextBlock Grid.Row="3" Text="Baseline Information" Style="{StaticResource SectionTitle}"/>
                    <Grid Grid.Row="4" Margin="0,0,0,4">
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="16"/>
                            <ColumnDefinition Width="*"/>
                        </Grid.ColumnDefinitions>
                        <Grid.RowDefinitions>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="Auto"/>
                        </Grid.RowDefinitions>
                        <TextBlock Grid.Column="0" Grid.Row="0" Text="Product Name *" FontWeight="SemiBold" Margin="0,0,0,4"/>
                        <TextBox x:Name="txtProductName" Grid.Column="0" Grid.Row="1" Padding="6,4" Background="{DynamicResource InputBgBrush}" BorderBrush="{DynamicResource BorderBrush2}"/>
                        <TextBlock Grid.Column="2" Grid.Row="0" Text="Intune Prefix (auto if blank)" FontWeight="SemiBold" Margin="0,0,0,4"/>
                        <TextBox x:Name="txtIntunePrefix" Grid.Column="2" Grid.Row="1" Padding="6,4" Foreground="Gray" Background="{DynamicResource InputBgBrush}" BorderBrush="{DynamicResource BorderBrush2}"/>
                    </Grid>

                    <Grid Grid.Row="5" Margin="0,0,0,4">
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="Auto"/>
                            <ColumnDefinition Width="Auto"/>
                        </Grid.ColumnDefinitions>
                        <TextBlock Grid.Column="0" Text="Conversion Options" Style="{StaticResource SectionTitle}" Margin="0,12,0,6"/>
                        <TextBlock x:Name="txtPdfDependencyStatus" Grid.Column="1" VerticalAlignment="Center" FontSize="11" Foreground="{DynamicResource SubTextBrush}" Margin="12,0,8,0"/>
                        <Button x:Name="btnInstallReq" Grid.Column="2" Content="Install" Visibility="Collapsed" Style="{StaticResource SecondaryBtn}" MinHeight="28" Padding="10,3"/>
                    </Grid>
                    <CheckBox x:Name="chkPrepareForIntune" Grid.Row="6" Content="Prepare output for Intune (include Intune-specific columns)" IsChecked="True" Margin="0,0,0,8"/>

                    <Grid Grid.Row="7" Margin="0,0,0,8">
                        <Grid.RowDefinitions>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="Auto"/>
                        </Grid.RowDefinitions>
                        <TextBlock Grid.Row="0" Text="Profile level" FontWeight="SemiBold" Margin="0,0,0,6"/>
                        <StackPanel Grid.Row="1" Orientation="Horizontal">
                            <RadioButton x:Name="rdoL1" Content="L1 + BL" GroupName="PdfProfileFilter" IsChecked="True" VerticalAlignment="Center" Margin="0,0,16,0"/>
                            <RadioButton x:Name="rdoL2" Content="L2 only" GroupName="PdfProfileFilter" VerticalAlignment="Center" Margin="0,0,16,0"/>
                            <RadioButton x:Name="rdoBoth" Content="Both" GroupName="PdfProfileFilter" VerticalAlignment="Center"/>
                        </StackPanel>
                    </Grid>

                    <TextBlock Grid.Row="8" Text="Output" Style="{StaticResource SectionTitle}"/>
                    <TextBlock Grid.Row="9" Text="Output folder (default: PDF folder)" Margin="0,0,0,4"/>
                    <Grid Grid.Row="10" Margin="0,0,0,4">
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="Auto"/>
                        </Grid.ColumnDefinitions>
                        <TextBox x:Name="txtOutputFolder" Grid.Column="0" IsReadOnly="True" Padding="6,4" Background="{DynamicResource InputBgBrush}" BorderBrush="{DynamicResource BorderBrush2}"/>
                        <Button x:Name="btnBrowseOutput" Grid.Column="1" Content="Browse..." Width="90" Margin="8,0,0,0" Style="{StaticResource SecondaryBtn}"/>
                    </Grid>

                    <TextBlock Grid.Row="11" Text="Conversion Action" Style="{StaticResource SectionTitle}"/>
                    <WrapPanel Grid.Row="12" Orientation="Horizontal" Margin="0,0,0,8">
                        <Button x:Name="btnConvert" Content="Convert" Style="{StaticResource PrimaryBtn}" Margin="0,0,8,0"/>
                        <Button x:Name="btnCopyClipboard" Content="Copy to Clipboard" Style="{StaticResource SecondaryBtn}" IsEnabled="False" Margin="0,0,8,0"/>
                        <Button x:Name="btnOpenHtml" Content="Open HTML" Style="{StaticResource SecondaryBtn}" IsEnabled="False"/>
                    </WrapPanel>

                    <ProgressBar x:Name="pdfProgressBar" Grid.Row="13" Height="8" Margin="0,0,0,4" IsIndeterminate="False" Minimum="0" Maximum="100"/>
                    <TextBlock x:Name="txtPdfStatus" Grid.Row="14" Text="" FontSize="11" Foreground="{DynamicResource SubTextBrush}" Margin="0,0,0,4"/>
                    <Expander x:Name="pdfActivityExpander" Grid.Row="15" IsExpanded="False" Background="{DynamicResource SurfaceBrush}" BorderThickness="0" Margin="0,4,0,0">
                        <Expander.Header>
                            <Grid>
                                <Grid.ColumnDefinitions><ColumnDefinition Width="Auto"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
                                <TextBlock Text="Activity" Style="{StaticResource ActivityHeading}"/>
                                <TextBlock x:Name="txtPdfActivitySummary" Grid.Column="1" Style="{StaticResource ActivitySummary}"/>
                            </Grid>
                        </Expander.Header>
                        <TextBox x:Name="txtPdfLog" IsReadOnly="True" TextWrapping="Wrap" VerticalScrollBarVisibility="Auto"
                                 Style="{StaticResource ActivityLog}" Height="120"/>
                    </Expander>
                </Grid>
            </TabItem>

            <!-- Tab 2: Export Change History -->
            <TabItem Header="  Export Change History  ">
                <Grid Margin="12">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                    </Grid.RowDefinitions>

                    <TextBlock Grid.Row="0" Text="Benchmark PDF *" FontWeight="SemiBold" Margin="0,0,0,4"/>
                    <Grid Grid.Row="1" Margin="0,0,0,8">
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="Auto"/>
                        </Grid.ColumnDefinitions>
                        <TextBox x:Name="txtChangeHistoryPdf" Grid.Column="0" IsReadOnly="True" Padding="6,4" Background="{DynamicResource InputBgBrush}" BorderBrush="{DynamicResource BorderBrush2}"/>
                        <Button x:Name="btnChangeHistoryBrowse" Grid.Column="1" Content="Browse..." Width="90" Margin="8,0,0,0" Style="{StaticResource SecondaryBtn}"/>
                    </Grid>

                    <TextBlock Grid.Row="2" Text="Output Folder (default: PDF folder)" Margin="0,0,0,4"/>
                    <Grid Grid.Row="3" Margin="0,0,0,8">
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="Auto"/>
                        </Grid.ColumnDefinitions>
                        <TextBox x:Name="txtChangeHistoryOutputFolder" Grid.Column="0" IsReadOnly="True" Padding="6,4" Background="{DynamicResource InputBgBrush}" BorderBrush="{DynamicResource BorderBrush2}"/>
                        <Button x:Name="btnChangeHistoryBrowseOutput" Grid.Column="1" Content="Browse..." Width="90" Margin="8,0,0,0" Style="{StaticResource SecondaryBtn}"/>
                    </Grid>

                    <WrapPanel Grid.Row="4" Orientation="Horizontal" Margin="0,0,0,8">
                        <Button x:Name="btnChangeHistoryExport" Content="Export" Style="{StaticResource PrimaryBtn}" Margin="0,0,8,0"/>
                        <Button x:Name="btnChangeHistoryOpen" Content="Open HTML" Style="{StaticResource SecondaryBtn}" IsEnabled="False"/>
                    </WrapPanel>

                    <Grid Grid.Row="5" Margin="0,0,0,8">
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="Auto"/>
                            <ColumnDefinition Width="120"/>
                            <ColumnDefinition Width="*"/>
                        </Grid.ColumnDefinitions>
                        <TextBlock Grid.Column="0" Text="Similarity Threshold (% or 0-1):" VerticalAlignment="Center" Margin="0,0,8,0"/>
                        <TextBox x:Name="txtChangeHistorySimilarityThreshold" Grid.Column="1" Text="0.80" Padding="6,4" Background="{DynamicResource InputBgBrush}" BorderBrush="{DynamicResource BorderBrush2}"/>
                    </Grid>

                    <Expander x:Name="changeHistoryActivityExpander" Grid.Row="6" IsExpanded="False" Background="{DynamicResource SurfaceBrush}" BorderThickness="0" VerticalAlignment="Top" VerticalContentAlignment="Top">
                        <Expander.Header>
                            <Grid><Grid.ColumnDefinitions><ColumnDefinition Width="Auto"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
                                <TextBlock Text="Activity" Style="{StaticResource ActivityHeading}"/>
                                <TextBlock x:Name="txtChangeHistoryActivitySummary" Grid.Column="1" Style="{StaticResource ActivitySummary}"/>
                            </Grid>
                        </Expander.Header>
                        <TextBox x:Name="txtChangeHistoryLog" IsReadOnly="True" TextWrapping="Wrap" VerticalScrollBarVisibility="Auto"
                                 Style="{StaticResource ActivityLog}" Height="130"/>
                    </Expander>
                </Grid>
            </TabItem>

            <!-- Tab 3: Markdown to HTML -->
            <TabItem Header="  Markdown to HTML  ">
                <Grid Margin="12">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                    </Grid.RowDefinitions>

                    <TextBlock Grid.Row="0" Text="Markdown File *" FontWeight="SemiBold" Margin="0,0,0,4"/>
                    <Grid Grid.Row="1" Margin="0,0,0,8">
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="Auto"/>
                        </Grid.ColumnDefinitions>
                        <TextBox x:Name="txtMdToHtmlInput" Grid.Column="0" IsReadOnly="True" Padding="6,4" Background="{DynamicResource InputBgBrush}" BorderBrush="{DynamicResource BorderBrush2}"/>
                        <Button x:Name="btnMdToHtmlBrowse" Grid.Column="1" Content="Browse..." Width="90" Margin="8,0,0,0" Style="{StaticResource SecondaryBtn}"/>
                    </Grid>

                    <TextBlock Grid.Row="2" Text="Output Folder (default: Markdown folder)" Margin="0,0,0,4"/>
                    <Grid Grid.Row="3" Margin="0,0,0,8">
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="Auto"/>
                        </Grid.ColumnDefinitions>
                        <TextBox x:Name="txtMdToHtmlOutputFolder" Grid.Column="0" IsReadOnly="True" Padding="6,4" Background="{DynamicResource InputBgBrush}" BorderBrush="{DynamicResource BorderBrush2}"/>
                        <Button x:Name="btnMdToHtmlBrowseOutput" Grid.Column="1" Content="Browse..." Width="90" Margin="8,0,0,0" Style="{StaticResource SecondaryBtn}"/>
                    </Grid>

                    <WrapPanel Grid.Row="4" Orientation="Horizontal" Margin="0,0,0,8">
                        <Button x:Name="btnMdToHtmlConvert" Content="Convert" Style="{StaticResource PrimaryBtn}" Margin="0,0,8,0"/>
                        <Button x:Name="btnMdToHtmlOpen" Content="Open HTML" Style="{StaticResource SecondaryBtn}" IsEnabled="False"/>
                    </WrapPanel>

                    <Expander x:Name="mdToHtmlActivityExpander" Grid.Row="5" IsExpanded="False" Background="{DynamicResource SurfaceBrush}" BorderThickness="0">
                        <Expander.Header>
                            <Grid><Grid.ColumnDefinitions><ColumnDefinition Width="Auto"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
                                <TextBlock Text="Activity" Style="{StaticResource ActivityHeading}"/>
                                <TextBlock x:Name="txtMdToHtmlActivitySummary" Grid.Column="1" Style="{StaticResource ActivitySummary}"/>
                            </Grid>
                        </Expander.Header>
                        <TextBox x:Name="txtMdToHtmlLog" IsReadOnly="True" TextWrapping="Wrap" VerticalScrollBarVisibility="Auto"
                                 Style="{StaticResource ActivityLog}" Height="130"/>
                    </Expander>
                </Grid>
            </TabItem>

            <!-- Tab 2: Baseline Merger -->
            <TabItem Header="  Baseline Merger  ">
                <Grid Margin="12">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                    </Grid.RowDefinitions>

                    <!-- New Baseline -->
                    <Grid Grid.Row="0" Margin="0,0,0,8">
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="130"/>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="Auto"/>
                        </Grid.ColumnDefinitions>
                        <TextBlock Grid.Column="0" Text="New Baseline:" VerticalAlignment="Center" FontWeight="SemiBold"/>
                        <TextBox x:Name="txtMergeNew" Grid.Column="1" Padding="6,4" Background="{DynamicResource InputBgBrush}" BorderBrush="{DynamicResource BorderBrush2}"/>
                        <Button x:Name="btnMergeNewBrowse" Grid.Column="2" Content="Browse..." Width="90" Margin="8,0,0,0" Style="{StaticResource SecondaryBtn}"/>
                    </Grid>

                    <!-- Old Baseline -->
                    <Grid Grid.Row="1" Margin="0,0,0,8">
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="130"/>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="Auto"/>
                        </Grid.ColumnDefinitions>
                        <TextBlock Grid.Column="0" Text="Old Baseline:" VerticalAlignment="Center" FontWeight="SemiBold"/>
                        <TextBox x:Name="txtMergeOld" Grid.Column="1" Padding="6,4" Background="{DynamicResource InputBgBrush}" BorderBrush="{DynamicResource BorderBrush2}"/>
                        <Button x:Name="btnMergeOldBrowse" Grid.Column="2" Content="Browse..." Width="90" Margin="8,0,0,0" Style="{StaticResource SecondaryBtn}"/>
                    </Grid>

                    <!-- Output File -->
                    <Grid Grid.Row="2" Margin="0,0,0,8">
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="130"/>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="Auto"/>
                        </Grid.ColumnDefinitions>
                        <TextBlock Grid.Column="0" Text="Output File (.txt):" VerticalAlignment="Center" FontWeight="SemiBold"/>
                        <TextBox x:Name="txtMergeOut" Grid.Column="1" Padding="6,4" Background="{DynamicResource InputBgBrush}" BorderBrush="{DynamicResource BorderBrush2}"/>
                        <Button x:Name="btnMergeOutBrowse" Grid.Column="2" Content="Browse..." Width="90" Margin="8,0,0,0" Style="{StaticResource SecondaryBtn}"/>
                    </Grid>

                    <!-- Options -->
                    <StackPanel Grid.Row="3" Margin="0,0,0,8">
                        <Grid Margin="0,0,0,8">
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="130"/>
                                <ColumnDefinition Width="*"/>
                            </Grid.ColumnDefinitions>
                            <TextBlock Text="Report Title:" VerticalAlignment="Center" FontWeight="SemiBold"/>
                            <TextBox x:Name="txtMergeTitle" Grid.Column="1" Padding="6,4" Background="{DynamicResource InputBgBrush}" BorderBrush="{DynamicResource BorderBrush2}"/>
                        </Grid>
                        <StackPanel Orientation="Horizontal" Margin="130,0,0,8">
                            <TextBlock Text="Profile:" VerticalAlignment="Center" FontWeight="SemiBold" Margin="0,0,12,0"/>
                            <RadioButton x:Name="rdoMergeL1" Content="L1 + BL" GroupName="MergeProfileFilter" IsChecked="True" VerticalAlignment="Center" Margin="0,0,16,0"/>
                            <RadioButton x:Name="rdoMergeL2" Content="L2 only" GroupName="MergeProfileFilter" VerticalAlignment="Center" Margin="0,0,16,0"/>
                            <RadioButton x:Name="rdoMergeAll" Content="All" GroupName="MergeProfileFilter" VerticalAlignment="Center"/>
                        </StackPanel>
                        <StackPanel Orientation="Horizontal" Margin="130,0,0,8">
                            <TextBlock Text="Text Match Threshold (% or 0-1):" VerticalAlignment="Center" FontWeight="SemiBold" Margin="0,0,12,0"/>
                            <TextBox x:Name="txtMergeSimilarityThreshold" Width="120" Text="0.90" Padding="6,4" Background="{DynamicResource InputBgBrush}" BorderBrush="{DynamicResource BorderBrush2}"/>
                        </StackPanel>
                        <CheckBox x:Name="chkMergeOpenBrowser" Content="Open HTML report in browser after generation" Margin="130,0,0,0"/>
                    </StackPanel>

                    <!-- Action Buttons -->
                    <WrapPanel Grid.Row="4" Orientation="Horizontal" Margin="0,0,0,8">
                        <Button x:Name="btnMergeOnly" Content="Merge Files" Style="{StaticResource SecondaryBtn}" Margin="0,0,8,0"/>
                        <Button x:Name="btnMergeAndReport" Content="Merge + Generate Report" Style="{StaticResource PrimaryBtn}" Margin="0,0,8,0"/>
                        <Button x:Name="btnReportOnly" Content="Generate Report from File..." Style="{StaticResource SecondaryBtn}" Margin="0,0,8,0"/>
                        <Button x:Name="btnCopyWiki" Content="Copy to DevOps Wiki" Style="{StaticResource SecondaryBtn}"/>
                    </WrapPanel>

                    <!-- Log -->
                    <Expander x:Name="mergeActivityExpander" Grid.Row="5" IsExpanded="False" Background="{DynamicResource SurfaceBrush}" BorderThickness="0">
                        <Expander.Header>
                            <Grid><Grid.ColumnDefinitions><ColumnDefinition Width="Auto"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
                                <TextBlock Text="Activity" Style="{StaticResource ActivityHeading}"/>
                                <TextBlock x:Name="txtMergeActivitySummary" Grid.Column="1" Style="{StaticResource ActivitySummary}"/>
                            </Grid>
                        </Expander.Header>
                        <TextBox x:Name="txtMergeLog" IsReadOnly="True" TextWrapping="Wrap" VerticalScrollBarVisibility="Auto"
                                 Style="{StaticResource ActivityLog}" Height="130"/>
                    </Expander>
                </Grid>
            </TabItem>

            <!-- Tab 3: Intune Policy Creator -->
            <TabItem Header="  Intune Policy Creator  ">
                <Grid Margin="12">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>

                    <!-- MD File Path -->
                    <Grid Grid.Row="1" Margin="0,0,0,8">
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="Auto"/>
                        </Grid.ColumnDefinitions>
                        <Grid.RowDefinitions>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="Auto"/>
                        </Grid.RowDefinitions>
                        <TextBlock Grid.Row="0" Grid.ColumnSpan="2" Text="CIS Benchmark Markdown File *" FontWeight="SemiBold" Margin="0,0,0,4"/>
                        <TextBox x:Name="txtPolicyMdPath" Grid.Row="1" Grid.Column="0" IsReadOnly="True" Padding="6,4" Background="{DynamicResource InputBgBrush}" BorderBrush="{DynamicResource BorderBrush2}"/>
                        <Button x:Name="btnPolicyBrowseMd" Grid.Row="1" Grid.Column="1" Content="Browse..." Width="90" Margin="8,0,0,0" Style="{StaticResource SecondaryBtn}"/>
                    </Grid>

                    <!-- Profile Level + Origin + Platform -->
                    <Grid Grid.Row="2" Margin="0,0,0,8">
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="16"/>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="16"/>
                            <ColumnDefinition Width="*"/>
                        </Grid.ColumnDefinitions>
                        <Grid.RowDefinitions>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="Auto"/>
                        </Grid.RowDefinitions>
                        <TextBlock Grid.Column="0" Grid.Row="0" Text="Profile Level" FontWeight="SemiBold" Margin="0,0,0,4"/>
                        <ComboBox x:Name="cmbPolicyProfile" Grid.Column="0" Grid.Row="1" Padding="6,4" Style="{StaticResource PolicyCreatorHighContrastComboBox}">
                            <ComboBoxItem Content="L1 - Level 1 only" IsSelected="True" Tag="L1"/>
                            <ComboBoxItem Content="L2 - Level 1 + Level 2" Tag="L2"/>
                            <ComboBoxItem Content="All - All recommendations" Tag="All"/>
                        </ComboBox>
                        <TextBlock Grid.Column="2" Grid.Row="0" Text="Origin" FontWeight="SemiBold" Margin="0,0,0,4"/>
                        <ComboBox x:Name="cmbPolicyOrigin" Grid.Column="2" Grid.Row="1" Padding="6,4" Style="{StaticResource PolicyCreatorHighContrastComboBox}">
                            <ComboBoxItem Content="All origins" IsSelected="True" Tag="All"/>
                            <ComboBoxItem Content="CIS only" Tag="CIS"/>
                            <ComboBoxItem Content="Twoday only" Tag="Twoday"/>
                        </ComboBox>
                        <StackPanel Grid.Column="4" Grid.Row="0" Grid.RowSpan="2">
                            <TextBlock Text="Platform" FontWeight="SemiBold" Margin="0,0,0,4"/>
                            <ComboBox x:Name="cmbPolicyPlatform" Padding="6,4" Margin="0,0,0,6" Style="{StaticResource PolicyCreatorHighContrastComboBox}">
                                <ComboBoxItem Content="windows10" IsSelected="True" Tag="windows10"/>
                                <ComboBoxItem Content="macOS" Tag="macOS"/>
                                <ComboBoxItem Content="iOS" Tag="iOS"/>
                                <ComboBoxItem Content="android" Tag="android"/>
                                <ComboBoxItem Content="linux" Tag="linux"/>
                            </ComboBox>
                            <TextBlock x:Name="txtPolicyProfileTypeLabel" Text="Profile Type" FontWeight="SemiBold" Margin="0,0,0,4" Visibility="Collapsed"/>
                            <ComboBox x:Name="cmbPolicyProfileType" Padding="6,4" Visibility="Collapsed" Style="{StaticResource PolicyCreatorHighContrastComboBox}">
                                <ComboBoxItem Content="Settings Catalog" IsSelected="True" Tag="SettingsCatalog"/>
                                <ComboBoxItem Content="Template" Tag="Template"/>
                            </ComboBox>
                        </StackPanel>
                    </Grid>

                    <!-- Info -->
                    <Border Grid.Row="3" Background="#1a2430" CornerRadius="4" Padding="10,8" Margin="0,0,0,8" BorderBrush="#2f4358" BorderThickness="1">
                        <TextBlock x:Name="txtPolicyInfo" Text="Select a Markdown file to see policy summary." FontSize="12" Foreground="{DynamicResource TextBrush}" TextWrapping="Wrap"/>
                    </Border>

                    <!-- Sub-TabControl: Selection + Output -->
                    <TabControl x:Name="policySubTabs" Grid.Row="4" SelectedIndex="0">
                        <TabItem Header="  Selection  ">
                            <Grid Margin="4">
                                <Grid.RowDefinitions>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="*"/>
                                </Grid.RowDefinitions>

                                <!-- Selection header: count + buttons -->
                                <Grid Grid.Row="0" Margin="0,4,0,4">
                                    <Grid.ColumnDefinitions>
                                        <ColumnDefinition Width="*"/>
                                        <ColumnDefinition Width="Auto"/>
                                        <ColumnDefinition Width="8"/>
                                        <ColumnDefinition Width="Auto"/>
                                    </Grid.ColumnDefinitions>
                                    <TextBlock x:Name="txtSelectionCount" Grid.Column="0" VerticalAlignment="Center" FontSize="12" Foreground="{DynamicResource SubTextBrush}"/>
                                    <Button x:Name="btnSelectAll" Grid.Column="1" Content="Select All" Width="100" Height="26" FontSize="11" Style="{StaticResource SecondaryBtn}"/>
                                    <Button x:Name="btnSelectNone" Grid.Column="3" Content="Select None" Width="100" Height="26" FontSize="11" Style="{StaticResource SecondaryBtn}"/>
                                </Grid>

                                <!-- Filter TextBox -->
                                <Grid Grid.Row="1" Margin="0,0,0,4">
                                    <TextBox x:Name="txtPolicyFilter" FontSize="12" Padding="4,3"
                                             ToolTip="Filter settings by name, CIS ref, value, or policy name"/>
                                    <TextBlock x:Name="txtPolicyFilterPlaceholder" Text="Filter settings..." FontSize="12" Foreground="{DynamicResource SubTextBrush}"
                                               Padding="6,5" IsHitTestVisible="False" Visibility="Visible"/>
                                </Grid>

                                <!-- DataGrid -->
                                <DataGrid x:Name="dgPolicies" Grid.Row="2"
                                          AutoGenerateColumns="False" IsReadOnly="False"
                                          HeadersVisibility="Column" GridLinesVisibility="Horizontal"
                                          CanUserSortColumns="True" CanUserReorderColumns="False"
                                          FontSize="11.5" SelectionMode="Extended"
                                          VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Auto"
                                          BorderBrush="#CCC" BorderThickness="1">
                                    <DataGrid.GroupStyle>
                                        <GroupStyle>
                                            <GroupStyle.ContainerStyle>
                                                <Style TargetType="{x:Type GroupItem}">
                                                    <Setter Property="Template">
                                                        <Setter.Value>
                                                            <ControlTemplate TargetType="{x:Type GroupItem}">
                                                                <Expander IsExpanded="False" Background="#1B2734" BorderBrush="{DynamicResource BorderBrush2}" BorderThickness="0,0,0,1">
                                                                    <Expander.Header>
                                                                        <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
                                                                            <CheckBox Tag="{Binding Path=Name}" IsThreeState="True" VerticalAlignment="Center" Margin="0,0,6,0"/>
                                                                            <TextBlock Text="{Binding Path=Name}" FontWeight="SemiBold" FontSize="12" Foreground="{DynamicResource TextBrush}" Margin="0,0,8,0" VerticalAlignment="Center"/>
                                                                            <TextBlock Text="{Binding Path=ItemCount, StringFormat='({0} settings)'}" FontSize="11" Foreground="{DynamicResource SubTextBrush}" VerticalAlignment="Center"/>
                                                                        </StackPanel>
                                                                    </Expander.Header>
                                                                    <ItemsPresenter/>
                                                                </Expander>
                                                            </ControlTemplate>
                                                        </Setter.Value>
                                                    </Setter>
                                                </Style>
                                            </GroupStyle.ContainerStyle>
                                        </GroupStyle>
                                    </DataGrid.GroupStyle>
                                    <DataGrid.RowStyle>
                                        <Style TargetType="DataGridRow">
                                            <Setter Property="Foreground" Value="{DynamicResource TextBrush}"/>
                                            <Style.Triggers>
                                                <Trigger Property="IsMouseOver" Value="True"><Setter Property="Background" Value="#1C2A38"/></Trigger>
                                                <Trigger Property="IsSelected" Value="True"><Setter Property="Background" Value="#1D352F"/></Trigger>
                                                <DataTrigger Binding="{Binding IsAutoCreationUnsupported}" Value="True">
                                                    <Setter Property="Foreground" Value="#F2B95F"/>
                                                    <Setter Property="FontStyle" Value="Italic"/>
                                                    <Setter Property="ToolTip" Value="Skipped: cannot be automatically created in Intune (from Notes)"/>
                                                </DataTrigger>
                                                <DataTrigger Binding="{Binding IsManual}" Value="True">
                                                    <Setter Property="Foreground" Value="#999999"/>
                                                    <Setter Property="FontStyle" Value="Italic"/>
                                                    <Setter Property="ToolTip" Value="Requires manual creation - company specific data needed"/>
                                                </DataTrigger>
                                            </Style.Triggers>
                                        </Style>
                                    </DataGrid.RowStyle>
                                    <DataGrid.Columns>
                                        <DataGridCheckBoxColumn Header="" Binding="{Binding Selected, UpdateSourceTrigger=PropertyChanged}" Width="30">
                                            <DataGridCheckBoxColumn.ElementStyle>
                                                <Style TargetType="CheckBox">
                                                    <Style.Triggers>
                                                        <DataTrigger Binding="{Binding IsManual}" Value="True">
                                                            <Setter Property="IsEnabled" Value="False"/>
                                                        </DataTrigger>
                                                    </Style.Triggers>
                                                </Style>
                                            </DataGridCheckBoxColumn.ElementStyle>
                                            <DataGridCheckBoxColumn.EditingElementStyle>
                                                <Style TargetType="CheckBox">
                                                    <Style.Triggers>
                                                        <DataTrigger Binding="{Binding IsManual}" Value="True">
                                                            <Setter Property="IsEnabled" Value="False"/>
                                                        </DataTrigger>
                                                    </Style.Triggers>
                                                </Style>
                                            </DataGridCheckBoxColumn.EditingElementStyle>
                                        </DataGridCheckBoxColumn>
                                        <DataGridTextColumn Header="CIS" Binding="{Binding CISRef}" Width="55" IsReadOnly="True"/>
                                        <DataGridTextColumn Header="Setting" Binding="{Binding SettingName}" Width="*" MinWidth="180" IsReadOnly="True"/>
                                        <DataGridTextColumn Header="Value" Binding="{Binding SettingValue}" Width="120" IsReadOnly="True"/>
                                        <DataGridTextColumn Header="L" Binding="{Binding Profile}" Width="30" IsReadOnly="True"/>
                                    </DataGrid.Columns>
                                </DataGrid>
                            </Grid>
                        </TabItem>
                        <TabItem Header="  Output  ">
                            <Expander x:Name="policyActivityExpander" IsExpanded="False" Background="{DynamicResource SurfaceBrush}" BorderThickness="0">
                                <Expander.Header>
                                    <Grid><Grid.ColumnDefinitions><ColumnDefinition Width="Auto"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
                                        <TextBlock Text="Activity" Style="{StaticResource ActivityHeading}"/>
                                        <TextBlock x:Name="txtPolicyActivitySummary" Grid.Column="1" Style="{StaticResource ActivitySummary}"/>
                                    </Grid>
                                </Expander.Header>
                                <TextBox x:Name="txtPolicyLog" IsReadOnly="True" TextWrapping="Wrap" VerticalScrollBarVisibility="Auto"
                                         Style="{StaticResource ActivityLog}" Height="130"/>
                            </Expander>
                        </TabItem>
                    </TabControl>

                    <!-- Action Buttons -->
                    <WrapPanel Grid.Row="5" Orientation="Horizontal" Margin="0,8,0,0">
                        <Button x:Name="btnPolicyExport" Content="Export JSON" Style="{StaticResource SecondaryBtn}" Margin="0,0,8,0"/>
                        <Button x:Name="btnPolicyCreate" Content="Create Policies" Style="{StaticResource PrimaryBtn}" Margin="0,0,8,0"/>
                        <Button x:Name="btnOpenReport" Content="Open Report" Style="{StaticResource SecondaryBtn}" Margin="0,0,8,0" IsEnabled="False" ToolTip="Open the interactive HTML creation report"/>
                    </WrapPanel>

                    <!-- Progress -->
                    <Grid Grid.Row="6" Margin="0,8,0,0">
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="Auto"/>
                        </Grid.ColumnDefinitions>
                        <ProgressBar x:Name="policyProgressBar" Grid.Column="0" Height="8" Minimum="0" Maximum="100"/>
                        <TextBlock x:Name="txtPolicyProgress" Grid.Column="1" Margin="10,0,0,0" VerticalAlignment="Center" FontSize="12" Foreground="{DynamicResource SubTextBrush}"/>
                    </Grid>

                    <!-- Status -->
                    <TextBlock x:Name="txtPolicyStatus" Grid.Row="7" Text="" FontSize="11" Foreground="{DynamicResource SubTextBrush}" Margin="0,4,0,0"/>
                </Grid>
            </TabItem>

            <!-- Tab 4: Intune Baseline Profile Report -->
            <TabItem Header="  Intune Policy Report  ">
                <Grid Margin="12">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                    </Grid.RowDefinitions>

                    <!-- Load and filter -->
                    <Grid Grid.Row="1" Margin="0,0,0,8">
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="Auto"/>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="Auto"/>
                        </Grid.ColumnDefinitions>
                        <TextBlock Grid.Column="0" Text="Filter profiles" VerticalAlignment="Center" Margin="0,0,8,0" FontWeight="SemiBold"/>
                        <TextBox x:Name="txtProfileReportFilter" Grid.Column="1" Padding="6,4" Text="" Background="{DynamicResource InputBgBrush}" BorderBrush="{DynamicResource BorderBrush2}"/>
                        <Button x:Name="btnProfileReportLoad" Grid.Column="2" Content="Load profiles" MinWidth="120" Margin="8,0,0,0" Style="{StaticResource PrimaryBtn}"/>
                    </Grid>

                    <!-- Selection controls -->
                    <Grid Grid.Row="2" Margin="0,0,0,8" MinHeight="34">
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="Auto"/>
                            <ColumnDefinition Width="Auto"/>
                        </Grid.ColumnDefinitions>
                        <TextBlock x:Name="txtProfileReportSelectionCount" Grid.Column="0" VerticalAlignment="Center" FontSize="12" Foreground="{DynamicResource SubTextBrush}"/>
                        <Button x:Name="btnProfileReportSelectAll" Grid.Column="1" Content="Select all" MinWidth="84" Style="{StaticResource SecondaryBtn}" Margin="0,0,8,0"/>
                        <Button x:Name="btnProfileReportSelectNone" Grid.Column="2" Content="Clear selection" MinWidth="110" Style="{StaticResource SecondaryBtn}"/>
                    </Grid>

                    <!-- Profile list -->
                    <Grid Grid.Row="3" Margin="0,0,0,8">
                        <DataGrid x:Name="dgProfileReportProfiles"
                                  AutoGenerateColumns="False" IsReadOnly="False"
                                  HeadersVisibility="Column" GridLinesVisibility="Horizontal"
                                  CanUserSortColumns="True" CanUserReorderColumns="False"
                                  FontSize="11.5" SelectionMode="Extended" SelectionUnit="FullRow"
                                  VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Auto"
                                  BorderBrush="{DynamicResource BorderBrush2}" BorderThickness="1"
                                  CanUserAddRows="False">
                            <DataGrid.Columns>
                                <DataGridCheckBoxColumn Header="" Binding="{Binding Selected, UpdateSourceTrigger=PropertyChanged}" Width="30"/>
                                <DataGridTextColumn Header="Name" Binding="{Binding Name}" Width="*" MinWidth="220" IsReadOnly="True"/>
                                <DataGridTextColumn Header="Type" Binding="{Binding PolicyKind}" Width="150" IsReadOnly="True"/>
                                <DataGridTextColumn Header="Last Modified" Binding="{Binding LastModifiedDateTime}" Width="160" IsReadOnly="True"/>
                            </DataGrid.Columns>
                        </DataGrid>
                        <Border x:Name="profileReportEmptyState" Background="#111922" BorderBrush="{DynamicResource BorderBrush2}" BorderThickness="1" Visibility="Visible" IsHitTestVisible="False">
                            <StackPanel HorizontalAlignment="Center" VerticalAlignment="Center" Margin="16" MaxWidth="520">
                                <TextBlock Text="No Intune profiles loaded" FontSize="13" FontWeight="SemiBold" TextAlignment="Center" TextWrapping="Wrap"/>
                                <TextBlock Text="Click Load profiles and select the policies to include." FontSize="12" Foreground="{DynamicResource SubTextBrush}" Margin="0,5,0,0" TextAlignment="Center" TextWrapping="Wrap"/>
                                <TextBlock Text="Leave Compare Folder blank for a policy report, or choose a JSON backup folder to compare settings." FontSize="12" Foreground="{DynamicResource SubTextBrush}" Margin="0,5,0,0" TextAlignment="Center" TextWrapping="Wrap"/>
                            </StackPanel>
                        </Border>
                    </Grid>

                    <!-- Output path -->
                    <Grid Grid.Row="4" Margin="0,0,0,8">
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="130"/>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="Auto"/>
                        </Grid.ColumnDefinitions>
                        <TextBlock Grid.Column="0" Text="Report location" VerticalAlignment="Center" FontWeight="SemiBold"/>
                        <TextBox x:Name="txtProfileReportOutputPath" Grid.Column="1" Padding="6,4" Background="{DynamicResource InputBgBrush}" BorderBrush="{DynamicResource BorderBrush2}"/>
                        <Button x:Name="btnProfileReportBrowseOutput" Grid.Column="2" Content="Browse..." Width="90" Margin="8,0,0,0" Style="{StaticResource SecondaryBtn}"/>
                    </Grid>

                    <!-- Baseline JSON folder for compliance comparison (optional) -->
                    <Grid Grid.Row="5" Margin="0,0,0,8">
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="130"/>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="Auto"/>
                        </Grid.ColumnDefinitions>
                        <TextBlock Grid.Column="0" Text="Compare Folder:" VerticalAlignment="Center" FontWeight="SemiBold" ToolTip="Optional folder containing JSON policy backups for settings comparison"/>
                        <TextBox x:Name="txtProfileReportBaselineFolder" Grid.Column="1" Padding="6,4" Background="{DynamicResource InputBgBrush}" BorderBrush="{DynamicResource BorderBrush2}" ToolTip="Leave blank for policy inventory. Otherwise match backups by policy ID, then unique name and type, and compare each policy's settings."/>
                        <Button x:Name="btnProfileReportBrowseBaseline" Grid.Column="2" Content="Browse..." Width="90" Margin="8,0,0,0" Style="{StaticResource SecondaryBtn}"/>
                    </Grid>

                    <!-- Actions -->
                    <WrapPanel Grid.Row="6" Orientation="Horizontal" Margin="0,0,0,8">
                        <Button x:Name="btnProfileReportGenerate" Content="Generate report" Style="{StaticResource PrimaryBtn}" Margin="0,0,8,0"/>
                        <Button x:Name="btnProfileReportWarmCache" Content="Warm Cache" Style="{StaticResource SecondaryBtn}" Margin="0,0,8,0" ToolTip="Resolve setting names for loaded policies and store to tenant cache"/>
                        <Button x:Name="btnProfileReportOpenHtml" Content="Open report" Style="{StaticResource SecondaryBtn}" IsEnabled="False" Margin="0,0,8,0"/>
                        <TextBlock Text="Cache mode:" VerticalAlignment="Center" Margin="0,0,6,0" FontWeight="SemiBold"/>
                        <ComboBox x:Name="cmbProfileReportCacheMode" Width="210" Margin="0,0,12,0" SelectedIndex="0" Foreground="#E7EDF3" Background="#111922" BorderBrush="#263849" ToolTip="Choose how name resolution should use cache and Graph lookups">
                            <ComboBox.Template>
                                <ControlTemplate TargetType="ComboBox">
                                    <Grid>
                                        <Border x:Name="SelectionBorder" Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="1" CornerRadius="2"/>
                                        <Grid>
                                            <Grid.ColumnDefinitions>
                                                <ColumnDefinition Width="*"/>
                                                <ColumnDefinition Width="28"/>
                                            </Grid.ColumnDefinitions>
                                            <ContentPresenter Grid.Column="0" IsHitTestVisible="False" Content="{TemplateBinding SelectionBoxItem}" ContentTemplate="{TemplateBinding SelectionBoxItemTemplate}" ContentTemplateSelector="{TemplateBinding ItemTemplateSelector}" TextElement.Foreground="{TemplateBinding Foreground}" Margin="8,0,4,0" VerticalAlignment="Center" HorizontalAlignment="Left"/>
                                            <ToggleButton Grid.Column="1" Focusable="False" Background="Transparent" BorderThickness="0" Foreground="{TemplateBinding Foreground}" IsChecked="{Binding IsDropDownOpen, Mode=TwoWay, RelativeSource={RelativeSource TemplatedParent}}" ClickMode="Press">
                                                <ToggleButton.Template>
                                                    <ControlTemplate TargetType="ToggleButton">
                                                        <Border Background="{TemplateBinding Background}">
                                                            <Path Data="M 0 0 L 4 4 L 8 0 Z" Fill="{TemplateBinding Foreground}" HorizontalAlignment="Center" VerticalAlignment="Center"/>
                                                        </Border>
                                                    </ControlTemplate>
                                                </ToggleButton.Template>
                                            </ToggleButton>
                                        </Grid>
                                        <Popup x:Name="Popup" Placement="Bottom" PlacementTarget="{Binding RelativeSource={RelativeSource TemplatedParent}}" IsOpen="{Binding IsDropDownOpen, Mode=TwoWay, RelativeSource={RelativeSource TemplatedParent}}" AllowsTransparency="True" Focusable="False" StaysOpen="False" PopupAnimation="Fade">
                                            <Border Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="1" MinWidth="210">
                                                <ScrollViewer MaxHeight="{TemplateBinding MaxDropDownHeight}" CanContentScroll="True">
                                                    <ItemsPresenter KeyboardNavigation.DirectionalNavigation="Contained"/>
                                                </ScrollViewer>
                                            </Border>
                                        </Popup>
                                    </Grid>
                                    <ControlTemplate.Triggers>
                                        <Trigger Property="IsKeyboardFocusWithin" Value="True">
                                            <Setter TargetName="SelectionBorder" Property="BorderBrush" Value="#36C18F"/>
                                        </Trigger>
                                    </ControlTemplate.Triggers>
                                </ControlTemplate>
                            </ComboBox.Template>
                            <ComboBox.ItemContainerStyle>
                                <Style TargetType="ComboBoxItem">
                                    <Setter Property="Foreground" Value="#E7EDF3"/>
                                    <Setter Property="Background" Value="#111922"/>
                                    <Setter Property="Padding" Value="8,5"/>
                                    <Setter Property="Template">
                                        <Setter.Value>
                                            <ControlTemplate TargetType="ComboBoxItem">
                                                <Border x:Name="ItemBorder" Background="{TemplateBinding Background}" Padding="{TemplateBinding Padding}">
                                                    <ContentPresenter VerticalAlignment="Center" RecognizesAccessKey="True"/>
                                                </Border>
                                                <ControlTemplate.Triggers>
                                                    <Trigger Property="IsHighlighted" Value="True">
                                                        <Setter TargetName="ItemBorder" Property="Background" Value="#1C2A38"/>
                                                        <Setter Property="Foreground" Value="#E7EDF3"/>
                                                    </Trigger>
                                                    <Trigger Property="IsSelected" Value="True">
                                                        <Setter TargetName="ItemBorder" Property="Background" Value="#36C18F"/>
                                                        <Setter Property="Foreground" Value="#0D151D"/>
                                                    </Trigger>
                                                </ControlTemplate.Triggers>
                                            </ControlTemplate>
                                        </Setter.Value>
                                    </Setter>
                                </Style>
                            </ComboBox.ItemContainerStyle>
                            <ComboBoxItem Content="Cache + Graph fallback" Tag="fallback"/>
                            <ComboBoxItem Content="Strict cache-only" Tag="cacheonly"/>
                            <ComboBoxItem Content="Force fresh (ignore cache)" Tag="forcefresh"/>
                        </ComboBox>
                        <CheckBox x:Name="chkProfileReportAutoOpen" Content="Open report when completed" VerticalAlignment="Center" IsChecked="True"/>
                    </WrapPanel>

                    <!-- Progress -->
                    <Grid Grid.Row="7" Margin="0,0,0,8">
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="Auto"/>
                        </Grid.ColumnDefinitions>
                        <ProgressBar x:Name="profileReportProgressBar" Grid.Column="0" Height="8" Minimum="0" Maximum="100"/>
                        <TextBlock x:Name="txtProfileReportProgress" Grid.Column="1" Margin="10,0,0,0" VerticalAlignment="Center" FontSize="12" Foreground="{DynamicResource SubTextBrush}"/>
                    </Grid>

                    <!-- Log -->
                    <Expander x:Name="profileReportActivityExpander" Grid.Row="8" IsExpanded="False" Background="{DynamicResource SurfaceBrush}" BorderThickness="0">
                        <Expander.Header>
                            <Grid><Grid.ColumnDefinitions><ColumnDefinition Width="Auto"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
                                <TextBlock Text="Activity" Style="{StaticResource ActivityHeading}"/>
                                <TextBlock x:Name="txtProfileReportActivitySummary" Grid.Column="1" Style="{StaticResource ActivitySummary}"/>
                            </Grid>
                        </Expander.Header>
                        <TextBox x:Name="txtProfileReportLog" IsReadOnly="True" TextWrapping="Wrap" VerticalScrollBarVisibility="Auto"
                                 Style="{StaticResource ActivityLog}" Height="130"/>
                    </Expander>
                </Grid>
            </TabItem>

            <!-- Tab 5: Baseline Assessment -->
            <TabItem Header="  Baseline Assessment  ">
                <Grid Margin="12">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>

                    <TextBlock Grid.Row="0" Text="Baseline Profiles" Style="{StaticResource SectionTitle}"/>

                    <!-- Baseline folder -->
                    <Grid Grid.Row="1" Margin="0,0,0,8">
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="130"/>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="Auto"/>
                        </Grid.ColumnDefinitions>
                        <StackPanel Grid.Column="0" VerticalAlignment="Center" ToolTip="Folder containing JSON Intune policy backup files to assess for import readiness">
                            <TextBlock Text="Baseline folder" FontSize="12" FontWeight="SemiBold"/>
                        </StackPanel>
                        <TextBox x:Name="txtAssessmentBaselineFolder" Grid.Column="1" Padding="8,5" ToolTip="Folder containing JSON Intune policy backup files"/>
                        <Button x:Name="btnAssessmentBrowseBaseline" Grid.Column="2" Content="Browse" Width="90" Margin="8,0,0,0" Style="{StaticResource SecondaryBtn}"/>
                    </Grid>

                    <!-- Load and filter -->
                    <Grid Grid.Row="2" Margin="0,0,0,8">
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="Auto"/>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="Auto"/>
                        </Grid.ColumnDefinitions>
                        <TextBlock Grid.Column="0" Text="Filter profiles" VerticalAlignment="Center" Margin="0,0,8,0" FontWeight="SemiBold"/>
                        <Grid Grid.Column="1">
                            <TextBox x:Name="txtAssessmentFilter" Padding="8,5" Text="" ToolTip="Search loaded intune profiles by name. Cleared automatically on reload."/>
                            <TextBlock x:Name="txtAssessmentFilterPlaceholder" Text="Search loaded intune profiles..." FontSize="12" Foreground="{DynamicResource SubTextBrush}" IsHitTestVisible="False" VerticalAlignment="Center" Margin="9,0"/>
                        </Grid>
                            <Button x:Name="btnAssessmentLoad" Grid.Column="2" Content="Load profiles" MinWidth="120" Margin="8,0,0,0" Style="{StaticResource PrimaryBtn}"/>
                    </Grid>

                    <!-- Selection controls -->
                    <Grid Grid.Row="3" Margin="0,0,0,8" MinHeight="34">
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="Auto"/>
                            <ColumnDefinition Width="Auto"/>
                        </Grid.ColumnDefinitions>
                        <TextBlock x:Name="txtAssessmentSelectionCount" Grid.Column="0" VerticalAlignment="Center" FontSize="12" Foreground="{DynamicResource SubTextBrush}"/>
                        <Button x:Name="btnAssessmentSelectAll" Grid.Column="1" Content="Select all" MinWidth="84" Style="{StaticResource SecondaryBtn}" Margin="0,0,8,0"/>
                        <Button x:Name="btnAssessmentSelectNone" Grid.Column="2" Content="Clear selection" MinWidth="110" Style="{StaticResource SecondaryBtn}"/>
                    </Grid>

                    <!-- Profile list -->
                    <Grid Grid.Row="4" Margin="0,0,0,8">
                        <DataGrid x:Name="dgAssessmentProfiles"
                                  AutoGenerateColumns="False" IsReadOnly="False"
                                  HeadersVisibility="Column" GridLinesVisibility="Horizontal"
                                  CanUserSortColumns="True" CanUserReorderColumns="False"
                                  FontSize="12" SelectionMode="Extended" SelectionUnit="FullRow"
                                  VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Auto"
                                  BorderBrush="{DynamicResource BorderBrush2}" BorderThickness="1"
                                  Background="#111922" Foreground="{DynamicResource TextBrush}"
                                  RowBackground="#111922" AlternatingRowBackground="#151F2A"
                                  HorizontalGridLinesBrush="{DynamicResource BorderBrush2}" VerticalGridLinesBrush="{DynamicResource BorderBrush2}"
                                  AlternationCount="2" CanUserAddRows="False">
                            <DataGrid.Resources>
                                <Style TargetType="DataGridColumnHeader">
                                    <Setter Property="Background" Value="#1B2734"/>
                                    <Setter Property="Foreground" Value="{DynamicResource TextBrush}"/>
                                    <Setter Property="BorderBrush" Value="{DynamicResource BorderBrush2}"/>
                                    <Setter Property="BorderThickness" Value="0,0,0,1"/>
                                    <Setter Property="Padding" Value="8,7"/>
                                    <Setter Property="FontWeight" Value="SemiBold"/>
                                </Style>
                                <Style TargetType="DataGridRow">
                                    <Setter Property="Foreground" Value="{DynamicResource ProfileRowTextBrush}"/>
                                    <Setter Property="Background" Value="#111922"/>
                                    <Style.Triggers>
                                        <Trigger Property="IsMouseOver" Value="True"><Setter Property="Background" Value="#1C2A38"/></Trigger>
                                        <Trigger Property="IsSelected" Value="True"><Setter Property="Background" Value="#1D352F"/></Trigger>
                                    </Style.Triggers>
                                </Style>
                                <Style TargetType="DataGridCell">
                                    <Setter Property="Foreground" Value="{DynamicResource ProfileRowTextBrush}"/>
                                    <Setter Property="Background" Value="Transparent"/>
                                    <Setter Property="BorderBrush" Value="Transparent"/>
                                    <Setter Property="Padding" Value="6,4"/>
                                    <Style.Triggers>
                                        <Trigger Property="IsSelected" Value="True">
                                            <Setter Property="Background" Value="#1D352F"/>
                                            <Setter Property="Foreground" Value="{DynamicResource TextBrush}"/>
                                        </Trigger>
                                        <Trigger Property="IsKeyboardFocusWithin" Value="True"><Setter Property="BorderBrush" Value="{DynamicResource AccentBrush}"/></Trigger>
                                    </Style.Triggers>
                                </Style>
                            </DataGrid.Resources>
                            <DataGrid.Columns>
                                <DataGridCheckBoxColumn Header="" Binding="{Binding Selected, UpdateSourceTrigger=PropertyChanged}" Width="40" MinWidth="40"/>
                                <DataGridTextColumn Header="Name" Binding="{Binding Name}" Width="*" MinWidth="220" IsReadOnly="True"/>
                                <DataGridTextColumn Header="Type" Binding="{Binding PolicyKind}" Width="150" IsReadOnly="True"/>
                                <DataGridTextColumn Header="Last Modified" Binding="{Binding LastModifiedDateTime}" Width="165" IsReadOnly="True"/>
                            </DataGrid.Columns>
                        </DataGrid>
                        <Border x:Name="assessmentEmptyState" Background="#111922" BorderBrush="{DynamicResource BorderBrush2}" BorderThickness="1" Visibility="Visible">
                            <StackPanel HorizontalAlignment="Center" VerticalAlignment="Center" Margin="16">
                                <TextBlock Text="No Intune profiles loaded" FontSize="13" FontWeight="SemiBold" HorizontalAlignment="Center"/>
                                <TextBlock Text="Select a baseline folder and load profiles to begin." FontSize="12" Foreground="{DynamicResource SubTextBrush}" Margin="0,5,0,0" HorizontalAlignment="Center"/>
                            </StackPanel>
                        </Border>
                    </Grid>

                    <!-- Assessment options -->
                    <StackPanel Grid.Row="5" Margin="0,0,0,10">
                        <TextBlock Text="Assessment Options" Style="{StaticResource SectionTitle}"/>
                        <Grid Margin="0,0,0,8">
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="Auto"/>
                                <ColumnDefinition Width="260"/>
                                <ColumnDefinition Width="*"/>
                            </Grid.ColumnDefinitions>
                            <TextBlock Grid.Column="0" Text="Cache mode" VerticalAlignment="Center" Margin="0,0,10,0" FontSize="12"/>
                            <ComboBox x:Name="cmbAssessmentCacheMode" Grid.Column="1" SelectedIndex="0" Foreground="#E7EDF3" Background="#111922" BorderBrush="#263849" ToolTip="Choose how name resolution uses cache and Graph lookups">
                                <ComboBox.Template>
                                    <ControlTemplate TargetType="ComboBox">
                                        <Grid>
                                            <Border x:Name="SelectionBorder" Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="1" CornerRadius="2"/>
                                            <Grid>
                                                <Grid.ColumnDefinitions>
                                                    <ColumnDefinition Width="*"/>
                                                    <ColumnDefinition Width="28"/>
                                                </Grid.ColumnDefinitions>
                                                <ContentPresenter Grid.Column="0" IsHitTestVisible="False" Content="{TemplateBinding SelectionBoxItem}" ContentTemplate="{TemplateBinding SelectionBoxItemTemplate}" ContentTemplateSelector="{TemplateBinding ItemTemplateSelector}" TextElement.Foreground="{TemplateBinding Foreground}" Margin="8,0,4,0" VerticalAlignment="Center" HorizontalAlignment="Left"/>
                                                <ToggleButton Grid.Column="1" Focusable="False" Background="Transparent" BorderThickness="0" Foreground="{TemplateBinding Foreground}" IsChecked="{Binding IsDropDownOpen, Mode=TwoWay, RelativeSource={RelativeSource TemplatedParent}}" ClickMode="Press">
                                                    <ToggleButton.Template>
                                                        <ControlTemplate TargetType="ToggleButton">
                                                            <Border Background="{TemplateBinding Background}">
                                                                <Path Data="M 0 0 L 4 4 L 8 0 Z" Fill="{TemplateBinding Foreground}" HorizontalAlignment="Center" VerticalAlignment="Center"/>
                                                            </Border>
                                                        </ControlTemplate>
                                                    </ToggleButton.Template>
                                                </ToggleButton>
                                            </Grid>
                                            <Popup x:Name="Popup" Placement="Bottom" PlacementTarget="{Binding RelativeSource={RelativeSource TemplatedParent}}" IsOpen="{Binding IsDropDownOpen, Mode=TwoWay, RelativeSource={RelativeSource TemplatedParent}}" AllowsTransparency="True" Focusable="False" StaysOpen="False" PopupAnimation="Fade">
                                                <Border Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="1" MinWidth="260">
                                                    <ScrollViewer MaxHeight="{TemplateBinding MaxDropDownHeight}" CanContentScroll="True">
                                                        <ItemsPresenter KeyboardNavigation.DirectionalNavigation="Contained"/>
                                                    </ScrollViewer>
                                                </Border>
                                            </Popup>
                                        </Grid>
                                        <ControlTemplate.Triggers>
                                            <Trigger Property="IsKeyboardFocusWithin" Value="True">
                                                <Setter TargetName="SelectionBorder" Property="BorderBrush" Value="#36C18F"/>
                                            </Trigger>
                                        </ControlTemplate.Triggers>
                                    </ControlTemplate>
                                </ComboBox.Template>
                                <ComboBox.ItemContainerStyle>
                                    <Style TargetType="ComboBoxItem">
                                        <Setter Property="Foreground" Value="#E7EDF3"/>
                                        <Setter Property="Background" Value="#111922"/>
                                        <Setter Property="Padding" Value="8,5"/>
                                        <Setter Property="Template">
                                            <Setter.Value>
                                                <ControlTemplate TargetType="ComboBoxItem">
                                                    <Border x:Name="ItemBorder" Background="{TemplateBinding Background}" Padding="{TemplateBinding Padding}">
                                                        <ContentPresenter VerticalAlignment="Center" RecognizesAccessKey="True"/>
                                                    </Border>
                                                    <ControlTemplate.Triggers>
                                                        <Trigger Property="IsHighlighted" Value="True">
                                                            <Setter TargetName="ItemBorder" Property="Background" Value="#1C2A38"/>
                                                            <Setter Property="Foreground" Value="#E7EDF3"/>
                                                        </Trigger>
                                                        <Trigger Property="IsSelected" Value="True">
                                                            <Setter TargetName="ItemBorder" Property="Background" Value="#36C18F"/>
                                                            <Setter Property="Foreground" Value="#0D151D"/>
                                                        </Trigger>
                                                    </ControlTemplate.Triggers>
                                                </ControlTemplate>
                                            </Setter.Value>
                                        </Setter>
                                    </Style>
                                </ComboBox.ItemContainerStyle>
                                <ComboBoxItem Content="Cache + Graph fallback" Tag="fallback"/>
                                <ComboBoxItem Content="Strict cache-only" Tag="cacheonly"/>
                                <ComboBoxItem Content="Force fresh (ignore cache)" Tag="forcefresh"/>
                            </ComboBox>
                        </Grid>
                        <WrapPanel VerticalAlignment="Center">
                            <CheckBox x:Name="chkAssessmentSkipUnassigned" Content="Skip unassigned policies" VerticalAlignment="Center" IsChecked="False" Margin="0,0,16,0" ToolTip="Requires assignment lookups to be enabled."/>
                            <CheckBox x:Name="chkAssessmentSkipAssignments" Content="Skip assignment lookups" VerticalAlignment="Center" IsChecked="False" ToolTip="Avoids retrieving assignment information for loaded policies."/>
                        </WrapPanel>
                    </StackPanel>

                    <!-- Report -->
                    <StackPanel Grid.Row="6" Margin="0,0,0,10">
                        <TextBlock Text="Report" Style="{StaticResource SectionTitle}"/>
                        <TextBlock Text="Report location" FontSize="12" FontWeight="SemiBold" Margin="0,0,0,4"/>
                        <Grid Margin="0,0,0,6">
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="Auto"/>
                            </Grid.ColumnDefinitions>
                            <TextBox x:Name="txtAssessmentOutputPath" Grid.Column="0" Padding="8,5"/>
                            <Button x:Name="btnAssessmentBrowseOutput" Grid.Column="1" Content="Browse" Width="90" Margin="8,0,0,0" Style="{StaticResource SecondaryBtn}"/>
                        </Grid>
                        <CheckBox x:Name="chkAssessmentAutoOpen" Content="Open report when completed" VerticalAlignment="Center" IsChecked="True"/>
                    </StackPanel>

                    <!-- Run Assessment -->
                    <Grid Grid.Row="7" Margin="0,0,0,8">
                        <Grid.RowDefinitions>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="Auto"/>
                        </Grid.RowDefinitions>
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="Auto"/>
                        </Grid.ColumnDefinitions>
                        <TextBlock Grid.Row="0" Text="Assessment" Style="{StaticResource SectionTitle}" VerticalAlignment="Center" Margin="0"/>
                        <WrapPanel Grid.Row="0" Grid.Column="1" Margin="0,0,0,6">
                            <Button x:Name="btnAssessmentRun" Content="Run assessment" MinWidth="145" Style="{StaticResource SecondaryBtn}" Margin="0,0,8,0"/>
                            <Button x:Name="btnAssessmentOpenHtml" Content="Open report" Style="{StaticResource SecondaryBtn}" IsEnabled="False" Margin="0,0,8,0"/>
                            <Button x:Name="btnAssessmentCreateSplitPolicy" Content="Create Split Policy" Style="{StaticResource SecondaryBtn}" IsEnabled="False"/>
                        </WrapPanel>
                        <Grid Grid.Row="1" Grid.ColumnSpan="2">
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="Auto"/>
                                <ColumnDefinition Width="*"/>
                            </Grid.ColumnDefinitions>
                            <TextBlock x:Name="txtAssessmentProgress" Grid.Column="0" Margin="0,0,12,0" VerticalAlignment="Center" FontSize="11" Foreground="{DynamicResource SubTextBrush}"/>
                            <ProgressBar x:Name="assessmentProgressBar" Grid.Column="1" Height="8" Minimum="0" Maximum="100" Visibility="Collapsed"/>
                        </Grid>
                    </Grid>

                    <!-- Log -->
                    <Expander x:Name="assessmentActivityExpander" Grid.Row="8" IsExpanded="False" Background="{DynamicResource SurfaceBrush}" BorderThickness="0">
                        <Expander.Header>
                            <Grid><Grid.ColumnDefinitions><ColumnDefinition Width="Auto"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
                                <TextBlock Text="Activity" Style="{StaticResource ActivityHeading}"/>
                                <TextBlock x:Name="txtAssessmentActivitySummary" Grid.Column="1" Style="{StaticResource ActivitySummary}"/>
                            </Grid>
                        </Expander.Header>
                        <TextBox x:Name="txtAssessmentLog" IsReadOnly="True" TextWrapping="Wrap" VerticalScrollBarVisibility="Auto"
                                 FontFamily="Cascadia Mono, Consolas" FontSize="11" Background="#111922" Foreground="{DynamicResource TextBrush}" Padding="8" BorderThickness="0" Height="120"/>
                    </Expander>
                </Grid>
            </TabItem>
            <TabItem Header="  Baseline Workbook  ">
                <Grid Margin="12">
                    <ScrollViewer VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Disabled">
                        <StackPanel>
                            <TextBlock Text="Baseline Workbook" Style="{StaticResource SectionTitle}" Margin="0,0,0,8"/>
                            <TextBlock Text="Details CSV from the interactive Baseline Assessment report" FontWeight="SemiBold" Margin="0,0,0,4"/>
                            <Grid Margin="0,0,0,12">
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="*"/>
                                    <ColumnDefinition Width="Auto"/>
                                </Grid.ColumnDefinitions>
                                <TextBox x:Name="txtAssessmentWorkbookCsv" Grid.Column="0" IsReadOnly="True" Padding="8,5"/>
                                <Button x:Name="btnAssessmentWorkbookBrowseCsv" Grid.Column="1" Content="Browse..." Width="90" Margin="8,0,0,0" Style="{StaticResource SecondaryBtn}"/>
                            </Grid>

                            <TextBlock Text="Folder containing source Markdown files" FontWeight="SemiBold" Margin="0,0,0,4"/>
                            <Grid Margin="0,0,0,12">
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="*"/>
                                    <ColumnDefinition Width="Auto"/>
                                </Grid.ColumnDefinitions>
                                <TextBox x:Name="txtAssessmentWorkbookMarkdownFolder" Grid.Column="0" IsReadOnly="True" Padding="8,5"/>
                                <Button x:Name="btnAssessmentWorkbookBrowseMarkdown" Grid.Column="1" Content="Browse..." Width="90" Margin="8,0,0,0" Style="{StaticResource SecondaryBtn}"/>
                            </Grid>

                            <Grid Margin="0,0,0,12">
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="Auto"/>
                                    <ColumnDefinition Width="120"/>
                                    <ColumnDefinition Width="*"/>
                                </Grid.ColumnDefinitions>
                                <TextBlock Grid.Column="0" Text="Text match threshold (% or 0-1):" VerticalAlignment="Center" Margin="0,0,8,0" FontWeight="SemiBold"/>
                                <TextBox x:Name="txtAssessmentWorkbookSimilarityThreshold" Grid.Column="1" Text="0.90" Padding="6,4"/>
                            </Grid>

                            <TextBlock Text="Output workbook (.xlsx)" FontWeight="SemiBold" Margin="0,0,0,4"/>
                            <Grid Margin="0,0,0,12">
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="*"/>
                                    <ColumnDefinition Width="Auto"/>
                                </Grid.ColumnDefinitions>
                                <TextBox x:Name="txtAssessmentWorkbookOutput" Grid.Column="0" IsReadOnly="True" Padding="8,5"/>
                                <Button x:Name="btnAssessmentWorkbookBrowseOutput" Grid.Column="1" Content="Browse..." Width="90" Margin="8,0,0,0" Style="{StaticResource SecondaryBtn}"/>
                            </Grid>

                            <WrapPanel Margin="0,0,0,8">
                                <Button x:Name="btnAssessmentWorkbookExport" Content="Create Workbook" Style="{StaticResource PrimaryBtn}" Margin="0,0,8,0"/>
                                <Button x:Name="btnAssessmentWorkbookOpen" Content="Open Workbook" Style="{StaticResource SecondaryBtn}" IsEnabled="False"/>
                            </WrapPanel>
                            <ProgressBar x:Name="assessmentWorkbookProgressBar" Height="8" Margin="0,0,0,6" Minimum="0" Maximum="100"/>
                            <TextBlock x:Name="txtAssessmentWorkbookStatus" Text="" FontSize="11" Foreground="{DynamicResource SubTextBrush}" Margin="0,0,0,8" TextWrapping="Wrap"/>
                            <Expander x:Name="assessmentWorkbookActivityExpander" IsExpanded="False" Background="{DynamicResource SurfaceBrush}" BorderThickness="0">
                                <Expander.Header>
                                    <Grid><Grid.ColumnDefinitions><ColumnDefinition Width="Auto"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
                                        <TextBlock Text="Activity" Style="{StaticResource ActivityHeading}"/>
                                        <TextBlock x:Name="txtAssessmentWorkbookActivitySummary" Grid.Column="1" Style="{StaticResource ActivitySummary}"/>
                                    </Grid>
                                </Expander.Header>
                                <TextBox x:Name="txtAssessmentWorkbookLog" IsReadOnly="True" TextWrapping="Wrap" VerticalScrollBarVisibility="Auto" Style="{StaticResource ActivityLog}" Height="150"/>
                            </Expander>
                        </StackPanel>
                    </ScrollViewer>
                </Grid>
            </TabItem>
                </TabControl>
            </Grid>
        </Grid>
    </Grid>
</Window>
"@

# ══════════════════════════════════════════════════════════════════════════════
#  BUILD WINDOW
# ══════════════════════════════════════════════════════════════════════════════

$reader = [System.Xml.XmlNodeReader]::new($xaml)
try { $window = [System.Windows.Markup.XamlReader]::Load($reader) }
catch {
    [System.Windows.MessageBox]::Show("Failed to load XAML: $($_.Exception.Message)", 'Error', 'OK', 'Error')
    exit
}

$TxtVersion = $window.FindName('TxtVersion')
$MainTabs   = $window.FindName('MainTabs')
$txtCurrentPageTitle = $window.FindName('txtCurrentPageTitle')
$txtCurrentPageDescription = $window.FindName('txtCurrentPageDescription')
$script:graphStatusPanelByIndex = @{
    4 = $window.FindName('policyGraphPanel')
    5 = $window.FindName('profileReportGraphPanel')
    6 = $window.FindName('assessmentGraphPanel')
}

$script:toolPageInfoByIndex = @{
    0 = @{ Title = 'PDF to Markdown'; Description = 'Convert CIS Benchmark PDFs into structured Markdown content.' }
    1 = @{ Title = 'Export Change History'; Description = 'Compare benchmark revisions and export a change history report.' }
    2 = @{ Title = 'Markdown to HTML'; Description = 'Render Markdown files as browsable HTML output.' }
    3 = @{ Title = 'Baseline Merger'; Description = 'Merge baseline revisions and generate comparison reports.' }
    4 = @{ Title = 'Intune Policy Creator'; Description = 'Review recommendations and create Intune configuration policies.' }
    5 = @{ Title = 'Intune Policy Report'; Description = 'Inventory selected policies or compare their settings with JSON policy backups.' }
    6 = @{ Title = 'Baseline Assessment'; Description = 'Compare CIS baseline profiles and generate an HTML assessment report.' }
    7 = @{ Title = 'Baseline Workbook'; Description = 'Add Markdown CIS references to an assessment Details CSV workbook.' }
}
$script:updateToolPageHeader = {
    $pageInfo = $script:toolPageInfoByIndex[[int]$MainTabs.SelectedIndex]
    if ($pageInfo) {
        $txtCurrentPageTitle.Text = $pageInfo.Title
        $txtCurrentPageDescription.Text = $pageInfo.Description
    }
    foreach ($pageIndex in $script:graphStatusPanelByIndex.Keys) {
        $panelVisibility = if ([int]$pageIndex -eq [int]$MainTabs.SelectedIndex) { 'Visible' } else { 'Collapsed' }
        $script:graphStatusPanelByIndex[$pageIndex].Visibility = $panelVisibility
    }
}

$script:toolNavByIndex = @{}
$navigationNames = @(
    'navPdfToMarkdown',
    'navMarkdownToHtml',
    'navBaselineMerger',
    'navBaselineAssessment',
    'navAssessmentWorkbook',
    'navIntuneBaselineReport',
    'navIntunePolicyCreator',
    'navExportChangeHistory'
)
foreach ($navigationName in $navigationNames) {
    $navigationItem = $window.FindName($navigationName)
    if (-not $navigationItem) { continue }
    $script:toolNavByIndex[[int]$navigationItem.Tag] = $navigationItem
    $navigationItem.Add_Checked({
        param($navigationControl, $checkedEvent)
        if ($navigationControl.IsChecked) { $MainTabs.SelectedIndex = [int]$navigationControl.Tag }
    })
}
$MainTabs.Add_SelectionChanged({
    param($tabControl, $selectionEvent)
    $navigationItem = $script:toolNavByIndex[[int]$tabControl.SelectedIndex]
    if ($navigationItem -and -not $navigationItem.IsChecked) {
        $navigationItem.IsChecked = $true
    }
    & $script:updateToolPageHeader
})
if ($MainTabs.SelectedIndex -lt 0) { $MainTabs.SelectedIndex = 0 }
& $script:updateToolPageHeader

# Tab 1 — Markdown to HTML
$txtMdToHtmlInput   = $window.FindName('txtMdToHtmlInput')
$btnMdToHtmlBrowse  = $window.FindName('btnMdToHtmlBrowse')
$txtMdToHtmlOutputFolder = $window.FindName('txtMdToHtmlOutputFolder')
$btnMdToHtmlBrowseOutput = $window.FindName('btnMdToHtmlBrowseOutput')
$btnMdToHtmlConvert = $window.FindName('btnMdToHtmlConvert')
$btnMdToHtmlOpen    = $window.FindName('btnMdToHtmlOpen')
$txtMdToHtmlLog     = $window.FindName('txtMdToHtmlLog')

# Tab 2 — Export Change History
$txtChangeHistoryPdf   = $window.FindName('txtChangeHistoryPdf')
$btnChangeHistoryBrowse = $window.FindName('btnChangeHistoryBrowse')
$txtChangeHistoryOutputFolder = $window.FindName('txtChangeHistoryOutputFolder')
$btnChangeHistoryBrowseOutput = $window.FindName('btnChangeHistoryBrowseOutput')
$txtChangeHistorySimilarityThreshold = $window.FindName('txtChangeHistorySimilarityThreshold')
$btnChangeHistoryExport = $window.FindName('btnChangeHistoryExport')
$btnChangeHistoryOpen   = $window.FindName('btnChangeHistoryOpen')
$txtChangeHistoryLog    = $window.FindName('txtChangeHistoryLog')

# Tab 3 — PDF to Markdown
$txtPdfPath        = $window.FindName('txtPdfPath')
$btnBrowsePdf      = $window.FindName('btnBrowsePdf')
$txtProductName    = $window.FindName('txtProductName')
$txtIntunePrefix   = $window.FindName('txtIntunePrefix')
$chkPrepareForIntune = $window.FindName('chkPrepareForIntune')
$txtOutputFolder   = $window.FindName('txtOutputFolder')
$btnBrowseOutput   = $window.FindName('btnBrowseOutput')
$rdoL1             = $window.FindName('rdoL1')
$rdoL2             = $window.FindName('rdoL2')
$rdoBoth           = $window.FindName('rdoBoth')
$txtPdfDependencyStatus = $window.FindName('txtPdfDependencyStatus')
$btnInstallReq     = $window.FindName('btnInstallReq')
$btnConvert        = $window.FindName('btnConvert')
$btnCopyClipboard  = $window.FindName('btnCopyClipboard')
$btnOpenHtml       = $window.FindName('btnOpenHtml')
$txtPdfLog         = $window.FindName('txtPdfLog')
$pdfProgressBar    = $window.FindName('pdfProgressBar')
$txtPdfStatus      = $window.FindName('txtPdfStatus')

# Tab 2 — Merger
$txtMergeNew         = $window.FindName('txtMergeNew')
$btnMergeNewBrowse   = $window.FindName('btnMergeNewBrowse')
$txtMergeOld         = $window.FindName('txtMergeOld')
$btnMergeOldBrowse   = $window.FindName('btnMergeOldBrowse')
$txtMergeOut         = $window.FindName('txtMergeOut')
$btnMergeOutBrowse   = $window.FindName('btnMergeOutBrowse')
$txtMergeTitle       = $window.FindName('txtMergeTitle')
$rdoMergeL1          = $window.FindName('rdoMergeL1')
$rdoMergeL2          = $window.FindName('rdoMergeL2')
$rdoMergeAll         = $window.FindName('rdoMergeAll')
$txtMergeSimilarityThreshold = $window.FindName('txtMergeSimilarityThreshold')
$chkMergeOpenBrowser = $window.FindName('chkMergeOpenBrowser')
$btnMergeOnly        = $window.FindName('btnMergeOnly')
$btnMergeAndReport   = $window.FindName('btnMergeAndReport')
$btnReportOnly       = $window.FindName('btnReportOnly')
$btnCopyWiki         = $window.FindName('btnCopyWiki')
$txtMergeLog         = $window.FindName('txtMergeLog')

# Tab 3 — Policy Creator
$txtTenantInfo      = $window.FindName('txtTenantInfo')
$policyGraphStatus  = $window.FindName('policyGraphStatus')
$txtPolicyTenantName = $window.FindName('txtPolicyTenantName')
$btnGraphConnect    = $window.FindName('btnGraphConnect')
$txtPolicyMdPath    = $window.FindName('txtPolicyMdPath')
$btnPolicyBrowseMd  = $window.FindName('btnPolicyBrowseMd')
$cmbPolicyProfile   = $window.FindName('cmbPolicyProfile')
$cmbPolicyOrigin    = $window.FindName('cmbPolicyOrigin')
$cmbPolicyPlatform  = $window.FindName('cmbPolicyPlatform')
$txtPolicyProfileTypeLabel = $window.FindName('txtPolicyProfileTypeLabel')
$cmbPolicyProfileType = $window.FindName('cmbPolicyProfileType')
$txtPolicyInfo      = $window.FindName('txtPolicyInfo')
$policySubTabs      = $window.FindName('policySubTabs')
$txtPolicyLog       = $window.FindName('txtPolicyLog')
$dgPolicies         = $window.FindName('dgPolicies')
$txtSelectionCount  = $window.FindName('txtSelectionCount')
$btnSelectAll       = $window.FindName('btnSelectAll')
$btnSelectNone      = $window.FindName('btnSelectNone')
$txtPolicyFilter    = $window.FindName('txtPolicyFilter')
$txtPolicyFilterPlaceholder = $window.FindName('txtPolicyFilterPlaceholder')
$btnPolicyExport    = $window.FindName('btnPolicyExport')
$btnPolicyCreate    = $window.FindName('btnPolicyCreate')
$btnOpenReport      = $window.FindName('btnOpenReport')
$policyProgressBar  = $window.FindName('policyProgressBar')
$txtPolicyProgress  = $window.FindName('txtPolicyProgress')
$txtPolicyStatus    = $window.FindName('txtPolicyStatus')

# Tab 4 - Intune Baseline Report
$txtProfileReportTenantInfo   = $window.FindName('txtProfileReportTenantInfo')
$profileReportGraphStatus     = $window.FindName('profileReportGraphStatus')
$txtProfileReportTenantName   = $window.FindName('txtProfileReportTenantName')
$btnProfileReportGraphConnect = $window.FindName('btnProfileReportGraphConnect')
$txtProfileReportFilter       = $window.FindName('txtProfileReportFilter')
$btnProfileReportLoad         = $window.FindName('btnProfileReportLoad')
$txtProfileReportSelectionCount = $window.FindName('txtProfileReportSelectionCount')
$btnProfileReportSelectAll    = $window.FindName('btnProfileReportSelectAll')
$btnProfileReportSelectNone   = $window.FindName('btnProfileReportSelectNone')
$dgProfileReportProfiles      = $window.FindName('dgProfileReportProfiles')
$profileReportEmptyState      = $window.FindName('profileReportEmptyState')
$txtProfileReportOutputPath   = $window.FindName('txtProfileReportOutputPath')
$btnProfileReportBrowseOutput = $window.FindName('btnProfileReportBrowseOutput')
$txtProfileReportBaselineFolder = $window.FindName('txtProfileReportBaselineFolder')
$btnProfileReportBrowseBaseline = $window.FindName('btnProfileReportBrowseBaseline')
$btnProfileReportGenerate     = $window.FindName('btnProfileReportGenerate')
$btnProfileReportWarmCache    = $window.FindName('btnProfileReportWarmCache')
$btnProfileReportOpenHtml     = $window.FindName('btnProfileReportOpenHtml')
$cmbProfileReportCacheMode    = $window.FindName('cmbProfileReportCacheMode')
$chkProfileReportAutoOpen     = $window.FindName('chkProfileReportAutoOpen')
$profileReportProgressBar     = $window.FindName('profileReportProgressBar')
$txtProfileReportProgress     = $window.FindName('txtProfileReportProgress')
$txtProfileReportLog          = $window.FindName('txtProfileReportLog')

$txtProfileReportFilter.IsEnabled = $true
$txtProfileReportFilter.ToolTip = 'Filter loaded profiles by name.'

# Tab 5 - Baseline Assessment
$txtAssessmentTenantInfo      = $window.FindName('txtAssessmentTenantInfo')
$assessmentGraphStatus        = $window.FindName('assessmentGraphStatus')
$txtAssessmentTenantName      = $window.FindName('txtAssessmentTenantName')
$btnAssessmentGraphConnect    = $window.FindName('btnAssessmentGraphConnect')
$txtAssessmentBaselineFolder  = $window.FindName('txtAssessmentBaselineFolder')
$btnAssessmentBrowseBaseline  = $window.FindName('btnAssessmentBrowseBaseline')
$txtAssessmentFilter            = $window.FindName('txtAssessmentFilter')
$txtAssessmentFilterPlaceholder = $window.FindName('txtAssessmentFilterPlaceholder')
$btnAssessmentLoad              = $window.FindName('btnAssessmentLoad')
$txtAssessmentSelectionCount  = $window.FindName('txtAssessmentSelectionCount')
$btnAssessmentSelectAll       = $window.FindName('btnAssessmentSelectAll')
$btnAssessmentSelectNone      = $window.FindName('btnAssessmentSelectNone')
$dgAssessmentProfiles         = $window.FindName('dgAssessmentProfiles')
$assessmentEmptyState          = $window.FindName('assessmentEmptyState')
$txtAssessmentOutputPath      = $window.FindName('txtAssessmentOutputPath')
$btnAssessmentBrowseOutput    = $window.FindName('btnAssessmentBrowseOutput')
$btnAssessmentRun             = $window.FindName('btnAssessmentRun')
$btnAssessmentOpenHtml        = $window.FindName('btnAssessmentOpenHtml')
$btnAssessmentCreateSplitPolicy = $window.FindName('btnAssessmentCreateSplitPolicy')
$cmbAssessmentCacheMode       = $window.FindName('cmbAssessmentCacheMode')
$chkAssessmentAutoOpen        = $window.FindName('chkAssessmentAutoOpen')
$chkAssessmentSkipUnassigned  = $window.FindName('chkAssessmentSkipUnassigned')
$chkAssessmentSkipAssignments = $window.FindName('chkAssessmentSkipAssignments')
$setAssessmentAssignmentOptions = {
    $chkAssessmentSkipUnassigned.IsEnabled = -not [bool]$chkAssessmentSkipAssignments.IsChecked
}
$chkAssessmentSkipAssignments.Add_Checked($setAssessmentAssignmentOptions)
$chkAssessmentSkipAssignments.Add_Unchecked($setAssessmentAssignmentOptions)
& $setAssessmentAssignmentOptions
$assessmentProgressBar        = $window.FindName('assessmentProgressBar')
$txtAssessmentProgress        = $window.FindName('txtAssessmentProgress')
$txtAssessmentLog             = $window.FindName('txtAssessmentLog')

$txtAssessmentWorkbookCsv = $window.FindName('txtAssessmentWorkbookCsv')
$btnAssessmentWorkbookBrowseCsv = $window.FindName('btnAssessmentWorkbookBrowseCsv')
$txtAssessmentWorkbookMarkdownFolder = $window.FindName('txtAssessmentWorkbookMarkdownFolder')
$btnAssessmentWorkbookBrowseMarkdown = $window.FindName('btnAssessmentWorkbookBrowseMarkdown')
$txtAssessmentWorkbookSimilarityThreshold = $window.FindName('txtAssessmentWorkbookSimilarityThreshold')
$txtAssessmentWorkbookOutput = $window.FindName('txtAssessmentWorkbookOutput')
$btnAssessmentWorkbookBrowseOutput = $window.FindName('btnAssessmentWorkbookBrowseOutput')
$btnAssessmentWorkbookExport = $window.FindName('btnAssessmentWorkbookExport')
$btnAssessmentWorkbookOpen = $window.FindName('btnAssessmentWorkbookOpen')
$assessmentWorkbookProgressBar = $window.FindName('assessmentWorkbookProgressBar')
$txtAssessmentWorkbookStatus = $window.FindName('txtAssessmentWorkbookStatus')
$txtAssessmentWorkbookLog = $window.FindName('txtAssessmentWorkbookLog')
$assessmentWorkbookActivityExpander = $window.FindName('assessmentWorkbookActivityExpander')
$script:lastAssessmentWorkbookPath = $null

$script:assessmentProfilesAll = @()
$script:assessmentProfilesCollection = [System.Collections.ObjectModel.ObservableCollection[object]]::new()
$script:assessmentUiState = @{ IsBusy = $false }
$dgAssessmentProfiles.ItemsSource = $script:assessmentProfilesCollection
$script:assessmentLastHtmlPath = $null
$script:assessmentLastData = $null
# Shared container: the assessment run completes inside a GetNewClosure timer whose
# $script: scope differs from the main script scope, so results are stored on this
# hashtable (captured by reference) that both the timer and the click handlers can see.
$script:assessmentResultStore = @{ Data = $null; HtmlPath = $null }

$script:policyPreviewItems = $null
$script:lastHtmlReportPath = $null
$script:groupHeaderCheckboxRefs = @{}
$script:intuneReportProfilesAll = @()
$script:intuneReportProfilesFiltered = @()
$script:intuneReportProfilesCollection = [System.Collections.ObjectModel.ObservableCollection[object]]::new()
$script:intuneReportProfilesView = [System.Windows.Data.CollectionViewSource]::GetDefaultView($script:intuneReportProfilesCollection)
$dgProfileReportProfiles.ItemsSource = $script:intuneReportProfilesCollection
$script:intuneReportLastHtmlPath = $null
$script:lastMdToHtmlOutputPath = $null
$script:lastChangeHistoryHtmlPath = $null

$TxtVersion.Text = "v$script:ToolVersion"

# ══════════════════════════════════════════════════════════════════════════════
#  HELPER FUNCTIONS
# ══════════════════════════════════════════════════════════════════════════════

function Write-PdfLog {
    param([string]$Message)
    $txtPdfLog.Dispatcher.Invoke([Action]{ $txtPdfLog.AppendText("$Message`r`n"); $txtPdfLog.ScrollToEnd() })
}

function Update-PdfDependencyUi {
    param($Prerequisites)

    if ($Prerequisites.IsOk) {
        $txtPdfDependencyStatus.Text = 'Dependencies ready'
        $txtPdfDependencyStatus.Foreground = $window.FindResource('SuccessBrush')
        $btnInstallReq.Visibility = 'Collapsed'
        $btnInstallReq.IsEnabled = $false
        $btnConvert.IsEnabled = $true
    } else {
        $txtPdfDependencyStatus.Text = 'Requirements missing'
        $txtPdfDependencyStatus.Foreground = [System.Windows.Media.Brushes]::Goldenrod
        $btnInstallReq.Visibility = 'Visible'
        $btnInstallReq.IsEnabled = $true
        $btnConvert.IsEnabled = $false
    }
}

function Write-MdToHtmlLog {
    param([string]$Message, [string]$Level = 'INFO')
    $ts = Get-Date -Format 'HH:mm:ss'
    $line = "[$ts][$Level] $Message"
    $txtMdToHtmlLog.Dispatcher.Invoke([Action]{ $txtMdToHtmlLog.AppendText("$line`r`n"); $txtMdToHtmlLog.ScrollToEnd() })
}

function Write-ChangeHistoryLog {
    param([string]$Message, [string]$Level = 'INFO')
    $ts = Get-Date -Format 'HH:mm:ss'
    $line = "[$ts][$Level] $Message"
    $txtChangeHistoryLog.Dispatcher.Invoke([Action]{ $txtChangeHistoryLog.AppendText("$line`r`n"); $txtChangeHistoryLog.ScrollToEnd() })
}

function Write-MergeLog {
    param([string]$Message, [string]$Level = 'INFO')
    $ts   = Get-Date -Format 'HH:mm:ss'
    $line = "[$ts][$Level] $Message"
    $txtMergeLog.Dispatcher.Invoke([Action]{ $txtMergeLog.AppendText("$line`r`n"); $txtMergeLog.ScrollToEnd() })
}

function Write-PolicyLog {
    param([string]$Message)
    $txtPolicyLog.Dispatcher.Invoke([Action]{ $txtPolicyLog.AppendText("$Message`r`n"); $txtPolicyLog.ScrollToEnd() })
}

function Write-ProfileReportLog {
    param([string]$Message)
    $ts = Get-Date -Format 'HH:mm:ss'
    $line = "[$ts] $Message"
    $txtProfileReportLog.Dispatcher.Invoke([Action]{ $txtProfileReportLog.AppendText("$line`r`n"); $txtProfileReportLog.ScrollToEnd() })
}

function Get-CacheModeFromComboBox {
    param($ComboBox)

    try {
        $selected = if ($ComboBox) { $ComboBox.SelectedItem } else { $null }
        if ($selected) {
            $tag = Get-SafeObjectPropertyValue -Object $selected -Name 'Tag'
            if ($tag) {
                $mode = ([string]$tag).ToLowerInvariant()
                if ($mode -in @('fallback','cacheonly','forcefresh')) { return $mode }
            }
        }
    } catch {}

    return 'fallback'
}

function Update-ProfileReportTenantInfo {
    $ctx = try { Get-MgContext -ErrorAction SilentlyContinue } catch { $null }
    return (Set-GraphPanelState -StatusText $txtProfileReportTenantInfo -StatusDot $profileReportGraphStatus `
        -TenantNameText $txtProfileReportTenantName -ConnectButton $btnProfileReportGraphConnect -Context $ctx)
}

function Ensure-ProfileReportGraphConnection {
    try {
        $requiredCmdlets = @('Connect-MgGraph','Get-MgContext','Invoke-MgGraphRequest')
        $missingCmdlets = @($requiredCmdlets | Where-Object { -not (Get-Command $_ -ErrorAction SilentlyContinue) })
        if ($missingCmdlets.Count -gt 0) {
            throw "Missing required Microsoft Graph cmdlets: $($missingCmdlets -join ', '). Install/Update Microsoft.Graph modules and restart PowerShell."
        }

        $requiredTenant = Get-ResolverTenantOverride
        if (Test-MgSessionIsValid) {
            if ($requiredTenant -and -not (Test-MgSessionTenantMatch -TenantId $requiredTenant)) {
                [System.Windows.MessageBox]::Show("Current Graph session is connected to a different tenant.`nRequired tenant: $requiredTenant`n`nClick Sign In to reconnect.", 'Wrong Tenant', 'OK', 'Warning') | Out-Null
                return $false
            }
            return $true
        }

        [System.Windows.MessageBox]::Show('Please sign in to Microsoft Graph before loading profile data.', 'Not Connected', 'OK', 'Warning') | Out-Null
        return $false
    } catch {
        [System.Windows.MessageBox]::Show("Graph error: $($_.Exception.Message)", 'Error', 'OK', 'Error') | Out-Null
        return $false
    }
}

function Update-ProfileReportSelectionCount {
    $gridSource = $dgProfileReportProfiles.ItemsSource
    $allItems = if ($null -ne $gridSource) { @($gridSource | ForEach-Object { $_ }) } else { @() }
    $allCount = [int]($allItems | Measure-Object).Count
    if ($profileReportEmptyState) {
        $profileReportEmptyState.Visibility = if ($allCount -eq 0) { 'Visible' } else { 'Collapsed' }
    }

    $view = if ($null -ne $gridSource) { [System.Windows.Data.CollectionViewSource]::GetDefaultView($gridSource) } else { $null }
    $filteredCount = if ($null -ne $view) {
        [int]($view | Measure-Object).Count
    } else {
        $allCount
    }

    $selectedCount = [int]($allItems | Where-Object { $_.Selected } | Measure-Object).Count

    $txtProfileReportSelectionCount.Text = "$selectedCount selected | $filteredCount filtered | $allCount loaded"
}

function Apply-ProfileReportFilter {
    param([string]$FilterText = $null)

    $termSource = if ($null -ne $FilterText) { [string]$FilterText } else { [string]$txtProfileReportFilter.Text }
    $term = $termSource.Trim()

    $gridSource = $dgProfileReportProfiles.ItemsSource
    if ($null -eq $gridSource) {
        $script:intuneReportProfilesFiltered = @()
        Update-ProfileReportSelectionCount
        return
    }

    $allItems = @($gridSource | ForEach-Object { $_ })
    if (($allItems | Measure-Object).Count -eq 0) {
        $script:intuneReportProfilesFiltered = @()
        Update-ProfileReportSelectionCount
        return
    }

    $view = [System.Windows.Data.CollectionViewSource]::GetDefaultView($gridSource)
    if ($null -eq $view) {
        $script:intuneReportProfilesFiltered = @()
        Update-ProfileReportSelectionCount
        return
    }

    if ($term) {
        $view.Filter = {
            param($item)
            if (-not $item) { return $false }
            $name = [string]$item.Name
            $name.IndexOf($term, [System.StringComparison]::OrdinalIgnoreCase) -ge 0
        }.GetNewClosure()
    } else {
        $view.Filter = $null
    }

    $view.Refresh()
    $script:intuneReportProfilesFiltered = @($view | ForEach-Object { $_ })
    Update-ProfileReportSelectionCount
    Write-ProfileReportLog "Profiles shown in list: $(($script:intuneReportProfilesFiltered | Measure-Object).Count)"
}

# ═══════════════════════════════════════════════════════════════════════════════
#  BASELINE JSON COMPARISON HELPERS
# ═══════════════════════════════════════════════════════════════════════════════

# ═══════════════════════════════════════════════════════════════════════════════
#  BASELINE ASSESSMENT FUNCTIONS
# ═══════════════════════════════════════════════════════════════════════════════

function Show-SplitRequiredBaselinePicker {
    param([object[]]$Baselines)

    if (-not $Baselines -or $Baselines.Count -eq 0) { return @() }

    $pickerXaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Select Split Required Baselines" Height="440" Width="860" WindowStartupLocation="CenterOwner" ResizeMode="CanResizeWithGrip" Background="#F7FAFC">
    <Grid Margin="12">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        <TextBlock Text="Tick one or more split-required baselines to create policies from:" FontWeight="SemiBold" Margin="0,0,0,8"/>
        <StackPanel Grid.Row="1" Orientation="Horizontal" Margin="0,0,0,8">
            <Button x:Name="btnPickerSelectAll" Content="Select All" Width="90" Margin="0,0,8,0"/>
            <Button x:Name="btnPickerSelectNone" Content="Select None" Width="90"/>
            <TextBlock x:Name="txtPickerCount" VerticalAlignment="Center" Margin="12,0,0,0" Foreground="{DynamicResource SubTextBrush}"/>
        </StackPanel>
        <DataGrid x:Name="dgPicker" Grid.Row="2" AutoGenerateColumns="False" IsReadOnly="False" SelectionMode="Extended" SelectionUnit="FullRow" HeadersVisibility="Column" GridLinesVisibility="Horizontal" CanUserAddRows="False">
            <DataGrid.Columns>
                <DataGridCheckBoxColumn Header="" Binding="{Binding Selected, UpdateSourceTrigger=PropertyChanged}" Width="34"/>
                <DataGridTextColumn Header="Baseline" Binding="{Binding PolicyName}" Width="2*" IsReadOnly="True"/>
                <DataGridTextColumn Header="Customer Policy" Binding="{Binding MatchedProfileName}" Width="2*" IsReadOnly="True"/>
                <DataGridTextColumn Header="Detail" Binding="{Binding Detail}" Width="3*" IsReadOnly="True"/>
            </DataGrid.Columns>
        </DataGrid>
        <StackPanel Grid.Row="3" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,10,0,0">
            <Button x:Name="btnOk" Content="Create Selected" Width="120" Margin="0,0,8,0" IsDefault="True"/>
            <Button x:Name="btnCancel" Content="Cancel" Width="90" IsCancel="True"/>
        </StackPanel>
    </Grid>
</Window>
"@

    # Wrap baselines in selectable rows so the checkbox column can bind two-way.
    $rows = [System.Collections.ObjectModel.ObservableCollection[object]]::new()
    foreach ($b in $Baselines) {
        $rows.Add([pscustomobject]@{
            Selected           = $false
            PolicyName         = [string]$b.PolicyName
            MatchedProfileName = [string]$b.MatchedProfileName
            Detail             = [string]$b.Detail
            Baseline           = $b
        })
    }

    $reader = [System.Xml.XmlNodeReader]::new(([xml]$pickerXaml))
    $dialog = [System.Windows.Markup.XamlReader]::Load($reader)
    $dialog.Owner = $window
    $dgPicker = $dialog.FindName('dgPicker')
    $btnOk = $dialog.FindName('btnOk')
    $btnPickerSelectAll = $dialog.FindName('btnPickerSelectAll')
    $btnPickerSelectNone = $dialog.FindName('btnPickerSelectNone')
    $txtPickerCount = $dialog.FindName('txtPickerCount')
    $dgPicker.ItemsSource = $rows

    $updatePickerCount = {
        try { $dgPicker.CommitEdit([System.Windows.Controls.DataGridEditingUnit]::Cell, $true) } catch {}
        $sel = @($rows | Where-Object { $_.Selected }).Count
        $txtPickerCount.Text = "$sel of $($rows.Count) selected"
    }.GetNewClosure()

    $btnPickerSelectAll.Add_Click({
        foreach ($r in $rows) { $r.Selected = $true }
        $dgPicker.Items.Refresh()
        & $updatePickerCount
    }.GetNewClosure())

    $btnPickerSelectNone.Add_Click({
        foreach ($r in $rows) { $r.Selected = $false }
        $dgPicker.Items.Refresh()
        & $updatePickerCount
    }.GetNewClosure())

    $dgPicker.Add_CellEditEnding({
        $dgPicker.Dispatcher.BeginInvoke([Action]{ & $updatePickerCount }) | Out-Null
    }.GetNewClosure())

    $btnOk.Add_Click({
        try { $dgPicker.CommitEdit([System.Windows.Controls.DataGridEditingUnit]::Cell, $true) } catch {}
        try { $dgPicker.CommitEdit([System.Windows.Controls.DataGridEditingUnit]::Row, $true) } catch {}
        $tickedCount = @($rows | Where-Object { $_.Selected }).Count
        if ($tickedCount -eq 0 -and $dgPicker.SelectedItems.Count -eq 0) {
            [System.Windows.MessageBox]::Show($dialog, 'Tick at least one baseline to continue.', 'No Selection', 'OK', 'Warning') | Out-Null
            return
        }
        $dialog.DialogResult = $true
    }.GetNewClosure())

    & $updatePickerCount
    $result = $dialog.ShowDialog()
    if ($result -ne $true) { return @() }

    try { $dgPicker.CommitEdit([System.Windows.Controls.DataGridEditingUnit]::Cell, $true) } catch {}
    # Collect selections from the ORIGINAL row objects so the underlying baseline
    # hashtables are returned intact (avoids WPF Tag/closure marshaling issues).
    $selected = @($rows | Where-Object { $_.Selected } | ForEach-Object { $_.Baseline })
    if ($selected.Count -eq 0 -and $dgPicker.SelectedItems.Count -gt 0) {
        $selected = @($dgPicker.SelectedItems | ForEach-Object { $_.Baseline })
    }
    return @($selected)
}

function Write-AssessmentLog {
    param([string]$Message)
    $ts = Get-Date -Format 'HH:mm:ss'
    $line = "[$ts] $Message"
    $txtAssessmentLog.Dispatcher.Invoke([Action]{ $txtAssessmentLog.AppendText("$line`r`n"); $txtAssessmentLog.ScrollToEnd() })
}

function Write-AssessmentWorkbookLog {
    param([string]$Message)
    $ts = Get-Date -Format 'HH:mm:ss'
    $line = "[$ts] $Message"
    $txtAssessmentWorkbookLog.Dispatcher.Invoke([Action]{ $txtAssessmentWorkbookLog.AppendText("$line`r`n"); $txtAssessmentWorkbookLog.ScrollToEnd() })
}

function Initialize-ActivityPanel {
    param(
        [System.Windows.Controls.TextBox]$LogTextBox,
        [System.Windows.Controls.Expander]$Expander,
        [System.Windows.Controls.TextBlock]$SummaryTextBlock
    )

    if (-not $LogTextBox -or -not $Expander -or -not $SummaryTextBlock) { return }
    $Expander.Tag = $LogTextBox.Text.Length

    $updateSummary = {
        $currentText = [string]$LogTextBox.Text
        $previousLength = [int]$Expander.Tag
        $newText = if ($currentText.Length -ge $previousLength) {
            $currentText.Substring($previousLength)
        } else {
            $currentText
        }
        $Expander.Tag = $currentText.Length

        $latestLine = $currentText.TrimEnd()
        $lineBreak = $latestLine.LastIndexOf("`n")
        if ($lineBreak -ge 0) { $latestLine = $latestLine.Substring($lineBreak + 1) }
        $latestLine = $latestLine.TrimEnd([char]13)
        if ($latestLine.Length -gt 160) { $latestLine = $latestLine.Substring(0, 157) + '...' }
        $fullLatestLine = $latestLine
        $summaryMessage = $latestLine
        $timestamp = $null
        $timestampMatch = [regex]::Match($latestLine, '^\[(?<Timestamp>\d{2}:\d{2}:\d{2})\]\s*(?<Message>.*)$')
        if ($timestampMatch.Success) {
            $timestamp = $timestampMatch.Groups['Timestamp'].Value
            $summaryMessage = $timestampMatch.Groups['Message'].Value
        }

        $statusKind = if ($summaryMessage -match '(?i)\b(warn(?:ing)?|error|exception)\b|\bfailed\s*:') {
            if ($summaryMessage -match '(?i)\b(warn(?:ing)?)\b') { 'warning' } else { 'error' }
        } elseif ($summaryMessage -match '(?i)\bno assigned profiles\b') {
            'warning'
        } elseif ($summaryMessage -match '(?i)\b(completed|created|succeeded|successfully)\b') {
            'success'
        } else {
            'normal'
        }

        $SummaryTextBlock.Inlines.Clear()
        $SummaryTextBlock.ToolTip = $fullLatestLine
        if ($statusKind -ne 'normal') {
            $statusBrushKey = switch ($statusKind) {
                'success' { 'SuccessBrush' }
                'warning' { 'AttentionBrush' }
                'error' { 'DangerBrush' }
            }
            $statusRun = [System.Windows.Documents.Run]::new('● ')
            $statusRun.Foreground = $SummaryTextBlock.FindResource($statusBrushKey)
            [void]$SummaryTextBlock.Inlines.Add($statusRun)
        }
        if ($timestamp) {
            $timestampRun = [System.Windows.Documents.Run]::new("$timestamp  ")
            $timestampRun.Foreground = $SummaryTextBlock.FindResource('SubTextBrush')
            [void]$SummaryTextBlock.Inlines.Add($timestampRun)
        }
        $messageRun = [System.Windows.Documents.Run]::new($summaryMessage)
        $messageRun.Foreground = $SummaryTextBlock.FindResource('TextBrush')
        [void]$SummaryTextBlock.Inlines.Add($messageRun)

        if ($newText -match '(?i)\b(warn(?:ing)?|error|exception)\b|\bfailed\s*:') {
            $Expander.IsExpanded = $true
        }
    }.GetNewClosure()

    $LogTextBox.Add_TextChanged($updateSummary)
    & $updateSummary
}

Initialize-ActivityPanel -LogTextBox $txtPdfLog -Expander ($window.FindName('pdfActivityExpander')) -SummaryTextBlock ($window.FindName('txtPdfActivitySummary'))
Initialize-ActivityPanel -LogTextBox $txtChangeHistoryLog -Expander ($window.FindName('changeHistoryActivityExpander')) -SummaryTextBlock ($window.FindName('txtChangeHistoryActivitySummary'))
Initialize-ActivityPanel -LogTextBox $txtMdToHtmlLog -Expander ($window.FindName('mdToHtmlActivityExpander')) -SummaryTextBlock ($window.FindName('txtMdToHtmlActivitySummary'))
Initialize-ActivityPanel -LogTextBox $txtMergeLog -Expander ($window.FindName('mergeActivityExpander')) -SummaryTextBlock ($window.FindName('txtMergeActivitySummary'))
Initialize-ActivityPanel -LogTextBox $txtPolicyLog -Expander ($window.FindName('policyActivityExpander')) -SummaryTextBlock ($window.FindName('txtPolicyActivitySummary'))
Initialize-ActivityPanel -LogTextBox $txtProfileReportLog -Expander ($window.FindName('profileReportActivityExpander')) -SummaryTextBlock ($window.FindName('txtProfileReportActivitySummary'))
Initialize-ActivityPanel -LogTextBox $txtAssessmentLog -Expander ($window.FindName('assessmentActivityExpander')) -SummaryTextBlock ($window.FindName('txtAssessmentActivitySummary'))
Initialize-ActivityPanel -LogTextBox $txtAssessmentWorkbookLog -Expander ($window.FindName('assessmentWorkbookActivityExpander')) -SummaryTextBlock ($window.FindName('txtAssessmentWorkbookActivitySummary'))

$script:graphTenantNameCache = @{}

function Set-GraphPanelState {
    param(
        [System.Windows.Controls.TextBlock]$StatusText,
        [System.Windows.Shapes.Ellipse]$StatusDot,
        [System.Windows.Controls.TextBlock]$TenantNameText,
        [System.Windows.Controls.Button]$ConnectButton,
        $Context
    )

    $tenantId = if ($Context -and $Context.TenantId) { [string]$Context.TenantId } else { '' }
    if ($tenantId) {
        if (-not $script:graphTenantNameCache.ContainsKey($tenantId)) {
            $tenantName = ''
            try {
                $organization = Invoke-MgGraphRequest -Uri 'https://graph.microsoft.com/v1.0/organization' -Method Get -ErrorAction Stop
                $tenantName = [string]$organization.value[0].displayName
            } catch { }
            if (-not $tenantName) { $tenantName = if ($Context.Account) { [string]$Context.Account } else { 'Connected' } }
            $script:graphTenantNameCache[$tenantId] = $tenantName
        }

        $account = if ($Context.Account) { [string]$Context.Account } else { '' }
        $StatusText.Text = 'Connected'
        $StatusText.Foreground = $window.FindResource('TextBrush')
        $StatusDot.SetResourceReference([System.Windows.Shapes.Shape]::FillProperty, 'SuccessBrush')
        $TenantNameText.Text = $script:graphTenantNameCache[$tenantId]
        $TenantNameText.ToolTip = "User: $account`nTenant ID: $tenantId"
        $ConnectButton.Content = 'Disconnect'
        $ConnectButton.Style = $window.FindResource('SecondaryBtn')
        return $true
    }

    $StatusText.Text = 'Not connected'
    $StatusText.Foreground = $window.FindResource('TextBrush')
    $StatusDot.SetResourceReference([System.Windows.Shapes.Shape]::FillProperty, 'DangerBrush')
    $TenantNameText.Text = ''
    $TenantNameText.ToolTip = ''
    $ConnectButton.Content = 'Connect'
    $ConnectButton.Style = $window.FindResource('SecondaryBtn')
    return $false
}

function Update-AssessmentTenantInfo {
    $ctx = try { Get-MgContext -ErrorAction SilentlyContinue } catch { $null }
    $connected = Set-GraphPanelState -StatusText $txtAssessmentTenantInfo -StatusDot $assessmentGraphStatus `
        -TenantNameText $txtAssessmentTenantName -ConnectButton $btnAssessmentGraphConnect -Context $ctx
    Update-AssessmentActionState
    return $connected
}

function Update-AssessmentSelectionCount {
    $gridSource = $dgAssessmentProfiles.ItemsSource
    $allItems = if ($null -ne $gridSource) { @($gridSource | ForEach-Object { $_ }) } else { @() }
    $allCount      = [int]($allItems | Measure-Object).Count
    $selectedCount = [int]($allItems | Where-Object { $_.Selected } | Measure-Object).Count
    if ($assessmentEmptyState) {
        $assessmentEmptyState.Visibility = if ($allCount -eq 0) { 'Visible' } else { 'Collapsed' }
    }
    if ($allCount -eq 0) {
        $txtAssessmentSelectionCount.Text = ''
        Update-AssessmentActionState
        return
    }
    $filterTerm = $txtAssessmentFilter.Text.Trim()
    if ($filterTerm -and $null -ne $gridSource) {
        $view = [System.Windows.Data.CollectionViewSource]::GetDefaultView($gridSource)
        $filteredCount = if ($view) { [int]@($view | ForEach-Object { $_ }).Count } else { $allCount }
        $txtAssessmentSelectionCount.Text = "$allCount profiles loaded • $selectedCount selected • $filteredCount shown"
    } else {
        $txtAssessmentSelectionCount.Text = "$allCount profiles loaded • $selectedCount selected"
    }
    Update-AssessmentActionState
}

function Update-AssessmentActionState {
    $gridSource = $dgAssessmentProfiles.ItemsSource
    $assessmentProfiles = @()
    if ($null -ne $gridSource) {
        $assessmentProfiles = @($gridSource | ForEach-Object { $_ })
    }
    $profilesLoaded = $assessmentProfiles.Count -gt 0
    $selectedCount = @($assessmentProfiles | Where-Object { $_.Selected }).Count
    $isConnected = Test-MgSessionIsValid

    $btnAssessmentLoad.Style = if ($profilesLoaded) { $window.FindResource('SecondaryBtn') } else { $window.FindResource('PrimaryBtn') }
    $btnAssessmentRun.Style = if ($profilesLoaded) { $window.FindResource('PrimaryBtn') } else { $window.FindResource('SecondaryBtn') }
    if (-not $script:assessmentUiState.IsBusy) {
        $btnAssessmentLoad.IsEnabled = [bool]$isConnected
        $btnAssessmentRun.IsEnabled = [bool]($profilesLoaded -and $selectedCount -gt 0 -and $isConnected)
    }
}

function Apply-AssessmentFilter {
    param([string]$FilterText = $null)
    $term = if ($null -ne $FilterText) { [string]$FilterText } else { [string]$txtAssessmentFilter.Text }
    $term = $term.Trim()
    $gridSource = $dgAssessmentProfiles.ItemsSource
    if ($null -eq $gridSource) { Update-AssessmentSelectionCount; return }
    $view = [System.Windows.Data.CollectionViewSource]::GetDefaultView($gridSource)
    if ($null -eq $view) { Update-AssessmentSelectionCount; return }
    if ($term) {
        $view.Filter = {
            param($item)
            if (-not $item) { return $false }
            $name = [string]$item.Name
            $name.IndexOf($term, [System.StringComparison]::OrdinalIgnoreCase) -ge 0
        }.GetNewClosure()
    } else {
        $view.Filter = $null
    }
    $view.Refresh()
    Update-AssessmentSelectionCount
}

function Update-IntunePrefixPlaceholder {
    if (-not $chkPrepareForIntune.IsChecked) { return }
    $name = $txtProductName.Text.Trim()
    if ($name -and -not $txtIntunePrefix.Tag) {
        $clean = $name -replace '\s', ''
        $txtIntunePrefix.Text       = "Prod-Win-All-SecurityBaseline-$clean-"
        $txtIntunePrefix.Foreground = [System.Windows.Media.Brushes]::Gray
    }
}

function Update-PrepareForIntuneUi {
    if ($chkPrepareForIntune.IsChecked) {
        $txtIntunePrefix.IsEnabled = $true
        if (-not $txtIntunePrefix.Text.Trim()) {
            $txtIntunePrefix.Tag = $null
            Update-IntunePrefixPlaceholder
        }
    } else {
        $txtIntunePrefix.IsEnabled  = $false
        $txtIntunePrefix.Tag        = $null
        $txtIntunePrefix.Text       = ''
        $txtIntunePrefix.Foreground = [System.Windows.Media.Brushes]::Gray
    }
}

function Try-GuessProductName {
    param([string]$FileName)
    $base = [System.IO.Path]::GetFileNameWithoutExtension($FileName)
    $base = $base -replace '^CIS_', '' -replace '_Benchmark.*$', '' -replace '_GPO', ''
    if ($base -match 'Intune_for_(.+)') { $base = "Microsoft $($Matches[1])" }
    return ($base -replace '_', ' ').Trim()
}

function Open-MarkdownFile {
    param([string]$Title = 'Select File')
    $dlg = [Microsoft.Win32.OpenFileDialog]::new()
    $dlg.Title  = $Title
    $dlg.Filter = 'Markdown / Text files (*.md;*.txt)|*.md;*.txt|All files (*.*)|*.*'
    if ($dlg.ShowDialog($window)) { return $dlg.FileName }
    return $null
}

function Update-PolicyTenantInfo {
    $ctx = try { Get-MgContext -ErrorAction SilentlyContinue } catch { $null }
    $connected = Set-GraphPanelState -StatusText $txtTenantInfo -StatusDot $policyGraphStatus `
        -TenantNameText $txtPolicyTenantName -ConnectButton $btnGraphConnect -Context $ctx
    Update-ProfileReportTenantInfo | Out-Null
    Update-AssessmentTenantInfo | Out-Null
    return $connected
}

# ══════════════════════════════════════════════════════════════════════════════
#  TAB 1 — MARKDOWN TO HTML EVENT HANDLERS
# ══════════════════════════════════════════════════════════════════════════════

$btnMdToHtmlBrowse.Add_Click({
    $path = Open-MarkdownFile 'Select Markdown File'
    if ($path) {
        $txtMdToHtmlInput.Text = $path
        $btnMdToHtmlOpen.IsEnabled = $false
        $script:lastMdToHtmlOutputPath = $null
        Write-MdToHtmlLog "Markdown selected: $path"
    }
})

$btnMdToHtmlBrowseOutput.Add_Click({
    $dlg = [System.Windows.Forms.FolderBrowserDialog]::new()
    $dlg.Description = 'Select output folder for HTML file'
    $dlg.ShowNewFolderButton = $true
    if ($txtMdToHtmlOutputFolder.Text) { $dlg.SelectedPath = $txtMdToHtmlOutputFolder.Text }
    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $txtMdToHtmlOutputFolder.Text = $dlg.SelectedPath
    }
})

$btnMdToHtmlOpen.Add_Click({
    if ($script:lastMdToHtmlOutputPath -and (Test-Path -Path $script:lastMdToHtmlOutputPath -PathType Leaf)) {
        Start-Process -FilePath $script:lastMdToHtmlOutputPath | Out-Null
    } else {
        [System.Windows.MessageBox]::Show('No HTML report found. Run a conversion first.', 'No File', 'OK', 'Warning') | Out-Null
    }
})

$btnMdToHtmlConvert.Add_Click({
    $inputPath = $txtMdToHtmlInput.Text.Trim()
    if (-not $inputPath) {
        [System.Windows.MessageBox]::Show('Please select a markdown file.', 'Missing', 'OK', 'Warning') | Out-Null
        return
    }
    if (-not (Test-Path -Path $inputPath -PathType Leaf)) {
        [System.Windows.MessageBox]::Show("Markdown file not found: $inputPath", 'Error', 'OK', 'Error') | Out-Null
        return
    }

    $converterScriptPath = Join-Path -Path (Split-Path -Path $script:thisScriptPath -Parent) -ChildPath 'Features\Convert-MarkdownTextToInteractiveHtmlReport.ps1'
    if (-not (Test-Path -Path $converterScriptPath -PathType Leaf)) {
        [System.Windows.MessageBox]::Show("Converter script not found: $converterScriptPath", 'Missing Script', 'OK', 'Error') | Out-Null
        return
    }

    $outputFolder = $txtMdToHtmlOutputFolder.Text.Trim()
    if (-not $outputFolder) { $outputFolder = Split-Path -Path $inputPath -Parent }
    if (-not (Test-Path -Path $outputFolder)) {
        try {
            New-Item -Path $outputFolder -ItemType Directory -Force | Out-Null
        } catch {
            [System.Windows.MessageBox]::Show("Unable to create output folder: $outputFolder`n$($_.Exception.Message)", 'Output Folder Error', 'OK', 'Error') | Out-Null
            return
        }
    }

    $outputPath = Join-Path -Path $outputFolder -ChildPath ([System.IO.Path]::GetFileNameWithoutExtension($inputPath) + '.html')
    $title = [System.IO.Path]::GetFileNameWithoutExtension($inputPath)

    $btnMdToHtmlConvert.IsEnabled = $false
    $btnMdToHtmlOpen.IsEnabled = $false
    Write-MdToHtmlLog "Starting conversion..."
    Write-MdToHtmlLog "Input: $inputPath"
    Write-MdToHtmlLog "Output folder: $outputFolder"

    try {
        $outputLines = & $converterScriptPath -InputFile $inputPath -OutputFile $outputPath -Title $title 2>&1
        foreach ($line in $outputLines) {
            if ($null -ne $line) { Write-MdToHtmlLog ([string]$line) }
        }

        if (Test-Path -Path $outputPath -PathType Leaf) {
            $script:lastMdToHtmlOutputPath = $outputPath
            $btnMdToHtmlOpen.IsEnabled = $true
            Write-MdToHtmlLog "Created: $outputPath"
        } else {
            Write-MdToHtmlLog 'Converter completed but output HTML was not found.' 'WARN'
        }
    }
    catch {
        Write-MdToHtmlLog $_.Exception.Message 'ERROR'
        [System.Windows.MessageBox]::Show($_.Exception.Message, 'Conversion Error', 'OK', 'Error') | Out-Null
    }
    finally {
        $btnMdToHtmlConvert.IsEnabled = $true
    }
})

# ══════════════════════════════════════════════════════════════════════════════
#  TAB 2 — EXPORT CHANGE HISTORY EVENT HANDLERS
# ══════════════════════════════════════════════════════════════════════════════

$btnChangeHistoryBrowse.Add_Click({
    $dlg = [Microsoft.Win32.OpenFileDialog]::new()
    $dlg.Filter = 'PDF Files (*.pdf)|*.pdf|All Files (*.*)|*.*'
    $dlg.Title  = 'Select Benchmark PDF'
    if ($dlg.ShowDialog($window)) {
        $txtChangeHistoryPdf.Text = $dlg.FileName
        $btnChangeHistoryOpen.IsEnabled = $false
        $script:lastChangeHistoryHtmlPath = $null
        Write-ChangeHistoryLog "PDF selected: $($dlg.FileName)"
    }
})

$btnChangeHistoryBrowseOutput.Add_Click({
    $dlg = [System.Windows.Forms.FolderBrowserDialog]::new()
    $dlg.Description = 'Select output folder for Change History HTML'
    $dlg.ShowNewFolderButton = $true
    if ($txtChangeHistoryOutputFolder.Text) { $dlg.SelectedPath = $txtChangeHistoryOutputFolder.Text }
    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $txtChangeHistoryOutputFolder.Text = $dlg.SelectedPath
    }
})

$btnChangeHistoryOpen.Add_Click({
    if ($script:lastChangeHistoryHtmlPath -and (Test-Path -Path $script:lastChangeHistoryHtmlPath -PathType Leaf)) {
        Start-Process -FilePath $script:lastChangeHistoryHtmlPath | Out-Null
    } else {
        [System.Windows.MessageBox]::Show('No HTML report found. Run export first.', 'No File', 'OK', 'Warning') | Out-Null
    }
})

$btnChangeHistoryExport.Add_Click({
    $pdfPath = $txtChangeHistoryPdf.Text.Trim()
    if (-not $pdfPath) {
        [System.Windows.MessageBox]::Show('Please select a PDF file.', 'Missing', 'OK', 'Warning') | Out-Null
        return
    }
    if (-not (Test-Path -Path $pdfPath -PathType Leaf)) {
        [System.Windows.MessageBox]::Show("PDF file not found: $pdfPath", 'Error', 'OK', 'Error') | Out-Null
        return
    }

    $changeHistorySimilarityThresholdRaw = $txtChangeHistorySimilarityThreshold.Text.Trim()
    $changeHistorySimilarityThreshold = 0.8
    if (-not [double]::TryParse($changeHistorySimilarityThresholdRaw, [System.Globalization.NumberStyles]::Float, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$changeHistorySimilarityThreshold)) {
        [System.Windows.MessageBox]::Show('Similarity Threshold must be a number (for example 0.80 or 80).', 'Invalid Threshold', 'OK', 'Warning') | Out-Null
        return
    }
    if ($changeHistorySimilarityThreshold -gt 1.0 -and $changeHistorySimilarityThreshold -le 100.0) { $changeHistorySimilarityThreshold = $changeHistorySimilarityThreshold / 100.0 }
    if ($changeHistorySimilarityThreshold -lt 0.0 -or $changeHistorySimilarityThreshold -gt 1.0) {
        [System.Windows.MessageBox]::Show('Similarity Threshold must be between 0 and 1, or 0 and 100.', 'Invalid Threshold', 'OK', 'Warning') | Out-Null
        return
    }

    $exportScriptPath = Join-Path -Path (Split-Path -Path $script:thisScriptPath -Parent) -ChildPath 'Features\Export-ChangeHistory.ps1'
    if (-not (Test-Path -Path $exportScriptPath -PathType Leaf)) {
        [System.Windows.MessageBox]::Show("Export script not found: $exportScriptPath", 'Missing Script', 'OK', 'Error') | Out-Null
        return
    }

    $outputDir = $txtChangeHistoryOutputFolder.Text.Trim()
    if (-not $outputDir) { $outputDir = Split-Path -Path $pdfPath -Parent }
    if (-not (Test-Path -Path $outputDir)) {
        try {
            New-Item -Path $outputDir -ItemType Directory -Force | Out-Null
        } catch {
            [System.Windows.MessageBox]::Show("Unable to create output folder: $outputDir`n$($_.Exception.Message)", 'Output Folder Error', 'OK', 'Error') | Out-Null
            return
        }
    }

    $pdfBase = [System.IO.Path]::GetFileNameWithoutExtension($pdfPath)
    $pdfBase = [regex]::Replace($pdfBase, '[_\-\s]+[vV]?\d+(?:\.\d+)*$', '')
    if ([string]::IsNullOrWhiteSpace($pdfBase)) { $pdfBase = 'ChangeHistory' }
    $htmlPattern = "$pdfBase`_ChangeHistory_v*.html"

    $btnChangeHistoryExport.IsEnabled = $false
    $btnChangeHistoryOpen.IsEnabled = $false
    Write-ChangeHistoryLog "Starting export..."
    Write-ChangeHistoryLog "Input: $pdfPath"
    Write-ChangeHistoryLog "Output folder: $outputDir"
    Write-ChangeHistoryLog "Similarity threshold: $changeHistorySimilarityThreshold"

    $startTime = Get-Date
    try {
        $outputLines = & $exportScriptPath -PdfPath $pdfPath -OutputDir $outputDir -SimilarityThreshold $changeHistorySimilarityThreshold *>&1
        foreach ($line in @($outputLines)) {
            if ($null -eq $line) { continue }
            Write-ChangeHistoryLog ([string]$line)
        }

        $explicitHtmlPath = $null
        foreach ($line in @($outputLines)) {
            $text = [string]$line
            $m = [regex]::Match($text, 'HTML\s*:\s*(.+)$')
            if ($m.Success) {
                $candidate = $m.Groups[1].Value.Trim()
                if ($candidate -and (Test-Path -Path $candidate -PathType Leaf)) {
                    $explicitHtmlPath = $candidate
                    break
                }
            }
        }

        $latestHtml = $null
        $htmlFiles = @(Get-ChildItem -Path $outputDir -Filter $htmlPattern -File -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending)
        if ($htmlFiles.Count -gt 0) {
            $latestHtml = ($htmlFiles | Where-Object { $_.LastWriteTime -ge $startTime.AddSeconds(-2) } | Select-Object -First 1)
            if (-not $latestHtml) { $latestHtml = $htmlFiles[0] }
        }

        $resolvedHtmlPath = if ($explicitHtmlPath) { $explicitHtmlPath } elseif ($latestHtml) { $latestHtml.FullName } else { $null }
        if ($resolvedHtmlPath -and (Test-Path -Path $resolvedHtmlPath -PathType Leaf)) {
            $script:lastChangeHistoryHtmlPath = $resolvedHtmlPath
            $btnChangeHistoryOpen.IsEnabled = $true
            Write-ChangeHistoryLog "HTML created: $resolvedHtmlPath"
        } else {
            Write-ChangeHistoryLog 'Export completed, but no HTML output could be detected.' 'WARN'
        }
    }
    catch {
        Write-ChangeHistoryLog $_.Exception.Message 'ERROR'
        [System.Windows.MessageBox]::Show($_.Exception.Message, 'Export Error', 'OK', 'Error') | Out-Null
    }
    finally {
        $btnChangeHistoryExport.IsEnabled = $true
    }
})

# ══════════════════════════════════════════════════════════════════════════════
#  TAB 3 — PDF TO MARKDOWN EVENT HANDLERS
# ══════════════════════════════════════════════════════════════════════════════

$btnBrowsePdf.Add_Click({
    $dlg = [Microsoft.Win32.OpenFileDialog]::new()
    $dlg.Filter = 'PDF Files (*.pdf)|*.pdf|All Files (*.*)|*.*'
    $dlg.Title  = 'Select CIS Benchmark PDF'
    if ($dlg.ShowDialog()) {
        $txtPdfPath.Text = $dlg.FileName
        if (-not $txtProductName.Text.Trim()) {
            $guess = Try-GuessProductName -FileName $dlg.SafeFileName
            if ($guess) { $txtProductName.Text = $guess; Update-IntunePrefixPlaceholder }
        }
    }
})

$btnBrowseOutput.Add_Click({
    $dlg = [System.Windows.Forms.FolderBrowserDialog]::new()
    $dlg.Description = 'Select output folder'
    $dlg.ShowNewFolderButton = $true
    if ($txtOutputFolder.Text) { $dlg.SelectedPath = $txtOutputFolder.Text }
    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { $txtOutputFolder.Text = $dlg.SelectedPath }
})

$txtProductName.Add_TextChanged({ Update-IntunePrefixPlaceholder })

$chkPrepareForIntune.Add_Checked({ Update-PrepareForIntuneUi })
$chkPrepareForIntune.Add_Unchecked({ Update-PrepareForIntuneUi })
Update-PrepareForIntuneUi

$txtIntunePrefix.Add_PreviewKeyDown({
    if (-not $chkPrepareForIntune.IsChecked) { return }
    $txtIntunePrefix.Tag        = 'manual'
    $txtIntunePrefix.Foreground = [System.Windows.Media.Brushes]::White
})
$txtIntunePrefix.Add_LostFocus({
    if (-not $txtIntunePrefix.Text.Trim()) { $txtIntunePrefix.Tag = $null; Update-IntunePrefixPlaceholder }
})

$btnInstallReq.Add_Click({
    $choice = [System.Windows.MessageBox]::Show("Install Python 3 dependencies (PyPDF2)?", 'Confirm', 'YesNo', 'Question')
    if ($choice -ne [System.Windows.MessageBoxResult]::Yes) { return }
    $pythonPath = $null
    $pythonKind = $null
    $cmd1 = Get-Command python -ErrorAction SilentlyContinue
    $cmd2 = Get-Command py     -ErrorAction SilentlyContinue
    if ($cmd1 -and (Test-Path $cmd1.Source)) {
        $pythonPath = $cmd1.Source
        $pythonKind = 'python'
    } elseif ($cmd2 -and (Test-Path $cmd2.Source)) {
        $pythonPath = $cmd2.Source
        $pythonKind = 'py'
    }
    if (-not $pythonPath) {
        [System.Windows.MessageBox]::Show('Python 3 was not found.', 'Not Found', 'OK', 'Error')
        return
    }
    $installCmd = if ($pythonKind -eq 'py') {
        "`$pp = `"$pythonPath`"`n& `$pp -3 -m pip install --upgrade pip 2>&1 | ForEach-Object { Write-Host `$_ }`n& `$pp -3 -m pip install PyPDF2 2>&1 | ForEach-Object { Write-Host `$_ }`n& `$pp -3 -c `"import PyPDF2`" 2>&1 | ForEach-Object { Write-Host `$_ }`nWrite-Host 'Done.' -ForegroundColor Green`nRead-Host 'Press Enter to close'"
    } else {
        "`$pp = `"$pythonPath`"`n& `$pp -m pip install --upgrade pip 2>&1 | ForEach-Object { Write-Host `$_ }`n& `$pp -m pip install PyPDF2 2>&1 | ForEach-Object { Write-Host `$_ }`n& `$pp -c `"import PyPDF2`" 2>&1 | ForEach-Object { Write-Host `$_ }`nWrite-Host 'Done.' -ForegroundColor Green`nRead-Host 'Press Enter to close'"
    }
    $tmp = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), [System.IO.Path]::GetRandomFileName() + '.ps1')
    [System.IO.File]::WriteAllText($tmp, $installCmd)
    try {
        $pi = [System.Diagnostics.ProcessStartInfo]::new()
        $pi.FileName = 'powershell.exe'
        $pi.Arguments = "-NoProfile -ExecutionPolicy Bypass -File `"$tmp`""
        $pi.Verb = 'runas'; $pi.UseShellExecute = $true
        $proc = [System.Diagnostics.Process]::Start($pi)
        $proc.WaitForExit()
        $prereq2 = Test-Prerequisites
        Update-PdfDependencyUi -Prerequisites $prereq2
        if ($prereq2.IsOk) {
            Write-PdfLog 'Requirements installed.'
        } else {
            Write-PdfLog ("Requirements check failed: " + ($prereq2.Issues -join '; '))
        }
    } catch { Write-PdfLog "Error: $($_.Exception.Message)" }
    finally { if (Test-Path $tmp) { Remove-Item $tmp -Force -ErrorAction SilentlyContinue } }
})

$btnCopyClipboard.Add_Click({
    $info = $window.Tag
    if ($info -and $info.TxtPath -and (Test-Path $info.TxtPath)) {
        [System.Windows.Clipboard]::SetText([System.IO.File]::ReadAllText($info.TxtPath, [System.Text.UTF8Encoding]::new($false)))
        Write-PdfLog 'Copied to clipboard.'
    } else {
        [System.Windows.MessageBox]::Show('No markdown file found. Run a conversion first.', 'No File', 'OK', 'Warning')
    }
})

$btnOpenHtml.Add_Click({
    $info = $window.Tag
    if ($info -and $info.HtmlPath -and (Test-Path $info.HtmlPath)) { Start-Process -FilePath $info.HtmlPath }
    else { [System.Windows.MessageBox]::Show('HTML file not found.', 'No File', 'OK', 'Warning') }
})

$btnConvert.Add_Click({
    $pdfPath     = $txtPdfPath.Text.Trim()
    $productName = $txtProductName.Text.Trim()
    if (-not $pdfPath)     { [System.Windows.MessageBox]::Show('Please select a PDF file.', 'Missing', 'OK', 'Warning'); return }
    if (-not (Test-Path $pdfPath -PathType Leaf)) { [System.Windows.MessageBox]::Show("PDF not found: $pdfPath", 'Error', 'OK', 'Error'); return }
    if (-not $productName) { [System.Windows.MessageBox]::Show('Enter a Product Name.', 'Missing', 'OK', 'Warning'); return }

    $intunePrefix   = $txtIntunePrefix.Text.Trim()
    $prepareForIntune = [bool]$chkPrepareForIntune.IsChecked
    $outputFolder   = $txtOutputFolder.Text.Trim()
    $profileFilter  = if ($rdoL1.IsChecked) { 'L1' } elseif ($rdoL2.IsChecked) { 'L2' } else { 'Both' }

    # Resolve defaults
    if ($prepareForIntune -and -not $intunePrefix) {
        $cleanName    = $productName -replace '\s', ''
        $intunePrefix = "Prod-Win-All-SecurityBaseline-$cleanName-"
    } elseif (-not $prepareForIntune) {
        $intunePrefix = ''
    }
    if (-not $outputFolder) { $outputFolder = Split-Path $pdfPath -Parent }
    if (-not (Test-Path $outputFolder)) { New-Item -Path $outputFolder -ItemType Directory -Force | Out-Null }

    # Find python executable
    $pythonExe = $null
    $pythonCmd = Get-Command python -ErrorAction SilentlyContinue
    if ($pythonCmd) { $pythonExe = $pythonCmd.Source }
    else {
        $pyCmd = Get-Command py -ErrorAction SilentlyContinue
        if ($pyCmd) { $pythonExe = $pyCmd.Source }
    }
    if (-not $pythonExe) {
        [System.Windows.MessageBox]::Show('Python 3 not found. Install Python and PyPDF2.', 'Error', 'OK', 'Error')
        return
    }

    $expectedTxtPath  = Join-Path $outputFolder "$productName.txt"
    $expectedHtmlPath = Join-Path $outputFolder "$productName.html"

    $txtPdfLog.Text = ''
    $btnConvert.IsEnabled        = $false
    $btnCopyClipboard.IsEnabled  = $false
    $pdfProgressBar.IsIndeterminate = $true
    Write-PdfLog "Starting conversion: $productName ($profileFilter)..."

    $runspace = [runspacefactory]::CreateRunspace()
    $runspace.ApartmentState = 'STA'
    $runspace.Open()
    $runspace.SessionStateProxy.SetVariable('pdfPath', $pdfPath)
    $runspace.SessionStateProxy.SetVariable('productName', $productName)
    $runspace.SessionStateProxy.SetVariable('intunePrefix', $intunePrefix)
    $runspace.SessionStateProxy.SetVariable('prepareForIntune', $prepareForIntune)
    $runspace.SessionStateProxy.SetVariable('outputFolder', $outputFolder)
    $runspace.SessionStateProxy.SetVariable('profileFilter', $profileFilter)
    $runspace.SessionStateProxy.SetVariable('pythonExe', $pythonExe)
    $runspace.SessionStateProxy.SetVariable('txtPdfLog', $txtPdfLog)
    $runspace.SessionStateProxy.SetVariable('pdfProgressBar', $pdfProgressBar)
    $runspace.SessionStateProxy.SetVariable('btnConvert', $btnConvert)
    $runspace.SessionStateProxy.SetVariable('btnCopyClipboard', $btnCopyClipboard)
    $runspace.SessionStateProxy.SetVariable('btnOpenHtml', $btnOpenHtml)
    $runspace.SessionStateProxy.SetVariable('expectedTxtPath', $expectedTxtPath)
    $runspace.SessionStateProxy.SetVariable('expectedHtmlPath', $expectedHtmlPath)
    $runspace.SessionStateProxy.SetVariable('thisScriptPath', $script:thisScriptPath)
    $runspace.SessionStateProxy.SetVariable('window', $window)

    $ps = [powershell]::Create()
    $ps.Runspace = $runspace
    $ps.AddScript({
        try {
            $converterPath = Join-Path (Split-Path $thisScriptPath -Parent) 'Features\Invoke-BaselinePdfConversion.ps1'
            $result = & $converterPath -PdfPath $pdfPath -ProductName $productName -IntunePrefix $intunePrefix -PrepareForIntune:$prepareForIntune -OutputFolder $outputFolder -ProfileFilter $profileFilter -LogCallback {
                param($message)
                $txtPdfLog.Dispatcher.Invoke([Action]{ $txtPdfLog.AppendText("$message`r`n"); $txtPdfLog.ScrollToEnd() })
            }
            $pdfProgressBar.Dispatcher.Invoke([Action]{ $pdfProgressBar.IsIndeterminate = $false; $pdfProgressBar.Value = 100 })
            $btnCopyClipboard.Dispatcher.Invoke([Action]{ $btnCopyClipboard.IsEnabled = $true })
            $btnOpenHtml.Dispatcher.Invoke([Action]{ $btnOpenHtml.IsEnabled = $true })
            $window.Dispatcher.Invoke([Action]{ $window.Tag = @{ TxtPath = $result.TextPath; HtmlPath = $result.HtmlPath } })
        } catch {
            $message = "ERROR: $($_.Exception.Message)"
            $txtPdfLog.Dispatcher.Invoke([Action]{ $txtPdfLog.AppendText("$message`r`n"); $txtPdfLog.ScrollToEnd() })
            $pdfProgressBar.Dispatcher.Invoke([Action]{ $pdfProgressBar.IsIndeterminate = $false; $pdfProgressBar.Value = 0 })
        } finally {
            $btnConvert.Dispatcher.Invoke([Action]{ $btnConvert.IsEnabled = $true })
        }
    }) | Out-Null
    $ps.BeginInvoke() | Out-Null
})

# ══════════════════════════════════════════════════════════════════════════════
#  TAB 2 — BASELINE MERGER EVENT HANDLERS
# ══════════════════════════════════════════════════════════════════════════════

$script:lastMismatches = @()

$btnMergeNewBrowse.Add_Click({
    $path = Open-MarkdownFile 'Select New Baseline File'
    if ($path) {
        $txtMergeNew.Text = $path
        Write-MergeLog "New Baseline: $path"
        if (-not $txtMergeOut.Text.Trim()) {
            $dir  = Split-Path $path -Parent
            $name = [System.IO.Path]::GetFileNameWithoutExtension($path)
            $txtMergeOut.Text = Join-Path $dir "$name-Merged.txt"
        }
        if (-not $txtMergeTitle.Text.Trim()) { $txtMergeTitle.Text = [System.IO.Path]::GetFileNameWithoutExtension($path) + ' (Merged)' }
    }
})

$btnMergeOldBrowse.Add_Click({
    $path = Open-MarkdownFile 'Select Old Baseline File'
    if ($path) { $txtMergeOld.Text = $path; Write-MergeLog "Old Baseline: $path" }
})

$btnMergeOutBrowse.Add_Click({
    $dlg = [Microsoft.Win32.SaveFileDialog]::new()
    $dlg.Title    = 'Save Merged File As'
    $dlg.Filter   = 'Text files (*.txt)|*.txt|All files (*.*)|*.*'
    $dlg.FileName = 'Merged-Baseline.txt'
    if ($dlg.ShowDialog($window)) { $txtMergeOut.Text = $dlg.FileName; Write-MergeLog "Output: $($dlg.FileName)" }
})

function Start-MergeOperation {
    $newPath = $txtMergeNew.Text.Trim()
    $oldPath = $txtMergeOld.Text.Trim()
    $outPath = $txtMergeOut.Text.Trim()
    if (-not $newPath) { [System.Windows.MessageBox]::Show('Select the New Baseline file.', 'Missing', 'OK', 'Warning'); return $false }
    if (-not $oldPath) { [System.Windows.MessageBox]::Show('Select the Old Baseline file.', 'Missing', 'OK', 'Warning'); return $false }
    if (-not $outPath) { [System.Windows.MessageBox]::Show('Specify an Output path.', 'Missing', 'OK', 'Warning'); return $false }
    if (-not (Test-Path $newPath)) { [System.Windows.MessageBox]::Show("New Baseline not found.", 'Error', 'OK', 'Error'); return $false }
    if (-not (Test-Path $oldPath)) { [System.Windows.MessageBox]::Show("Old Baseline not found.", 'Error', 'OK', 'Error'); return $false }

    $mergeProfileFilter = if ($rdoMergeL2.IsChecked) { 'L2' } elseif ($rdoMergeAll.IsChecked) { 'All' } else { 'L1' }
    $mergeSimilarityThresholdRaw = $txtMergeSimilarityThreshold.Text.Trim()
    $mergeSimilarityThreshold = 0.9
    if (-not [double]::TryParse($mergeSimilarityThresholdRaw, [System.Globalization.NumberStyles]::Float, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$mergeSimilarityThreshold)) {
        [System.Windows.MessageBox]::Show('Text Match Threshold must be a number (for example 0.90 or 90).', 'Invalid Threshold', 'OK', 'Warning') | Out-Null
        return $false
    }
    if ($mergeSimilarityThreshold -gt 1.0 -and $mergeSimilarityThreshold -le 100.0) { $mergeSimilarityThreshold = $mergeSimilarityThreshold / 100.0 }
    if ($mergeSimilarityThreshold -lt 0.0 -or $mergeSimilarityThreshold -gt 1.0) {
        [System.Windows.MessageBox]::Show('Text Match Threshold must be between 0 and 1, or 0 and 100.', 'Invalid Threshold', 'OK', 'Warning') | Out-Null
        return $false
    }

    Write-MergeLog "Starting merge... (Profile: $mergeProfileFilter, Similarity: $mergeSimilarityThreshold)"
    try {
        $stats = Invoke-MergeBaselines -NewBaselinePath $newPath -OldBaselinePath $oldPath -OutputPath $outPath -ProfileFilter $mergeProfileFilter -SimilarityThreshold $mergeSimilarityThreshold
        $script:lastMismatches = $stats.Mismatches
        Write-MergeLog "Merge complete. Overrides: $($stats.Overrides) | Notes copied: $($stats.NotesCopied) | Profile excluded: $($stats.ProfileFiltered) | Twoday: $($stats.TwodayCount) | Mismatches: $($stats.Mismatches.Count)"
        Write-MergeLog "Output: $outPath"
        return $true
    } catch {
        Write-MergeLog "ERROR: $($_.Exception.Message)" 'ERROR'
        [System.Windows.MessageBox]::Show("Merge failed: $($_.Exception.Message)", 'Error', 'OK', 'Error')
        return $false
    }
}

function Start-MergeHtmlReport {
    param([string]$MarkdownPath)
    $htmlPath = [System.IO.Path]::ChangeExtension($MarkdownPath, '.html')
    $title    = $txtMergeTitle.Text.Trim()
    if (-not $title) { $title = [System.IO.Path]::GetFileNameWithoutExtension($MarkdownPath) }
    Write-MergeLog "Generating HTML: $htmlPath"
    $mismatchesToPass = if ($null -ne $script:lastMismatches) { [hashtable[]]$script:lastMismatches } else { [hashtable[]]@() }
    try {
        Invoke-GenerateHtmlReport -InputMarkdownPath $MarkdownPath -OutputHtmlPath $htmlPath -ReportTitle $title -Mismatches $mismatchesToPass | Out-Null
        Write-MergeLog "Report created: $htmlPath"
        if ($chkMergeOpenBrowser.IsChecked) { Start-Process $htmlPath }
    } catch {
        Write-MergeLog "ERROR: $($_.Exception.Message)" 'ERROR'
        [System.Windows.MessageBox]::Show("HTML failed: $($_.Exception.Message)", 'Error', 'OK', 'Error')
    }
}

$btnMergeOnly.Add_Click({ Start-MergeOperation | Out-Null })

$btnMergeAndReport.Add_Click({
    if (Start-MergeOperation) { Start-MergeHtmlReport -MarkdownPath $txtMergeOut.Text.Trim() }
})

$btnReportOnly.Add_Click({
    $path = Open-MarkdownFile 'Select Markdown for Report'
    if (-not $path) { return }
    Write-MergeLog "Report source: $path"
    if (-not $txtMergeTitle.Text.Trim()) { $txtMergeTitle.Text = [System.IO.Path]::GetFileNameWithoutExtension($path) }
    $script:lastMismatches = @()
    Start-MergeHtmlReport -MarkdownPath $path
})

$btnCopyWiki.Add_Click({
    $outPath = $txtMergeOut.Text.Trim()
    if (-not $outPath -or -not (Test-Path $outPath -PathType Leaf)) {
        [System.Windows.MessageBox]::Show('No output file. Merge first.', 'No File', 'OK', 'Warning')
        return
    }
    try {
        $content = Get-Content -Path $outPath -Raw -Encoding UTF8
        [System.Windows.Clipboard]::SetText($content)
        Write-MergeLog "Copied $(($content -split "`n").Count) lines to clipboard."
    } catch {
        Write-MergeLog "ERROR: $($_.Exception.Message)" 'ERROR'
    }
})

# ══════════════════════════════════════════════════════════════════════════════
#  TAB 3 — INTUNE POLICY CREATOR EVENT HANDLERS
# ══════════════════════════════════════════════════════════════════════════════

#region Selection Tab Helpers

function Set-PolicyGridSource {
    param($Items)
    if (-not $Items -or @($Items).Count -eq 0) {
        $dgPolicies.ItemsSource = $null
        return
    }
    $view = [System.Windows.Data.CollectionViewSource]::GetDefaultView($Items)
    $view.GroupDescriptions.Clear()
    $view.GroupDescriptions.Add([System.Windows.Data.PropertyGroupDescription]::new('PolicyName'))
    $view.SortDescriptions.Clear()
    $view.SortDescriptions.Add([System.ComponentModel.SortDescription]::new('PolicyName', [System.ComponentModel.ListSortDirection]::Ascending))
    $view.SortDescriptions.Add([System.ComponentModel.SortDescription]::new('CISRef', [System.ComponentModel.ListSortDirection]::Ascending))
    $dgPolicies.ItemsSource = $view
}

function Update-PolicySelectionCount {
    $items = $script:policyPreviewItems
    if (-not $items) { $txtSelectionCount.Text = ''; return }
    $sel      = @($items | Where-Object { $_.Selected }).Count
    $tot      = @($items).Count
    $policies = @($items | Where-Object { $_.Selected } | ForEach-Object { $_.PolicyName } | Sort-Object -Unique).Count
    $filterText = $txtPolicyFilter.Text.Trim()
    $showCount = if ($dgPolicies.Items) { $dgPolicies.Items.Count } else { 0 }
    $filterInfo = if ($filterText) { " | showing $showCount matching '$filterText'" } else { '' }
    $txtSelectionCount.Text = "$sel of $tot settings selected ($policies policies)$filterInfo"
}

function Sync-PolicyGroupCheckboxes {
    $items = $script:policyPreviewItems
    if (-not $items) { return }
    $groupHeaderCheckboxes = @(Find-VisualChildren -Parent $dgPolicies -ChildType ([System.Windows.Controls.CheckBox]) | Where-Object { $_.Tag })
    foreach ($cb in $groupHeaderCheckboxes) {
        $policyName = [string]$cb.Tag
        $groupItems = @($items | Where-Object { $_.PolicyName -eq $policyName -and -not $_.IsManual })
        if ($groupItems.Count -eq 0) { $cb.IsChecked = $false; continue }
        $selectedCount = @($groupItems | Where-Object { $_.Selected }).Count
        if ($selectedCount -eq 0)                     { $cb.IsChecked = $false }
        elseif ($selectedCount -eq $groupItems.Count) { $cb.IsChecked = $true  }
        else                                          { $cb.IsChecked = $null  }
    }
}

function Find-VisualChildren {
    param(
        [Parameter(Mandatory = $true)]
        [System.Windows.DependencyObject]$Parent,
        [Parameter(Mandatory = $true)]
        [Type]$ChildType
    )

    if (-not $Parent) { return }
    $count = [System.Windows.Media.VisualTreeHelper]::GetChildrenCount($Parent)
    for ($i = 0; $i -lt $count; $i++) {
        $child = [System.Windows.Media.VisualTreeHelper]::GetChild($Parent, $i)
        if ($ChildType.IsAssignableFrom($child.GetType())) {
            $child
        }
        Find-VisualChildren -Parent $child -ChildType $ChildType
    }
}

function Get-CheckBoxFromEventSource {
    param([object]$Source)

    $current = $Source
    while ($current -and -not ($current -is [System.Windows.Controls.CheckBox])) {
        if ($current -is [System.Windows.DependencyObject]) {
            $parent = [System.Windows.Media.VisualTreeHelper]::GetParent($current)
            if (-not $parent) {
                $parent = [System.Windows.LogicalTreeHelper]::GetParent($current)
            }
            $current = $parent
        } else {
            $current = $null
        }
    }

    return $current
}

function Apply-PolicyTextFilter {
    $filterText = $txtPolicyFilter.Text.Trim()
    if (-not $script:policyPreviewItems) { return }
    if (-not $filterText) {
        Set-PolicyGridSource $script:policyPreviewItems
    } else {
        $filtered = [System.Collections.ObjectModel.ObservableCollection[RecItem]]::new()
        foreach ($item in $script:policyPreviewItems) {
            if ($item.CISRef       -like "*$filterText*" -or
                $item.SettingName  -like "*$filterText*" -or
                $item.SettingValue -like "*$filterText*" -or
                $item.PolicyName   -like "*$filterText*" -or
                $item.Recommendation -like "*$filterText*") {
                $filtered.Add($item)
            }
        }
        Set-PolicyGridSource $filtered
    }
    Update-PolicySelectionCount
    $dgPolicies.Dispatcher.BeginInvoke([Action]{ Sync-PolicyGroupCheckboxes }) | Out-Null
}

function Update-PolicyPreview {
    $mdPath = $txtPolicyMdPath.Text.Trim()
    if (-not $mdPath -or -not (Test-Path $mdPath)) {
        Set-PolicyGridSource $null
        $script:policyPreviewItems = $null
        $txtPolicyInfo.Text = 'Select a Markdown file to see policy summary.'
        Update-PolicySelectionCount
        return
    }
    try {
        $profileLevel = ($cmbPolicyProfile.SelectedItem).Tag
        $originFilter = ($cmbPolicyOrigin.SelectedItem).Tag
        $preview = Get-BaselinePolicyPreview -MdFilePath $mdPath -ProfileLevel $profileLevel -OriginFilter $originFilter
        $filteredRecs = @($preview.Recommendations)
        $script:groupHeaderCheckboxRefs = @{}
        $script:policyPreviewItems = [System.Collections.ObjectModel.ObservableCollection[RecItem]]::new()
        foreach ($rec in $filteredRecs) {
            $item               = [RecItem]::new()
            $item.CISRef        = $rec.CISReference
            $item.SettingName   = if ($rec.SettingName) { $rec.SettingName } else { '' }
            $item.SettingValue  = if ($rec.SettingValue) { $rec.SettingValue } else { '-' }
            $item.PolicyName    = if ($rec.IntunePolicy) { $rec.IntunePolicy } else { '' }
            $item.Profile       = if ($rec.Profile) { $rec.Profile } else { '' }
            $item.Recommendation = if ($rec.Recommendation) { $rec.Recommendation } else { '' }
            $item.Notes         = if ($rec.Notes) { $rec.Notes } else { '-' }
            $item.IsAutoCreationUnsupported = ($rec.Notes -and $rec.Notes -match 'Cannot\s+be\s+automatically\s+created\s+in\s+Intune')
            $item.IsCustomerInfoRequired = (-not $item.IsAutoCreationUnsupported) -and ($rec.Notes -and $rec.Notes -match 'Customer\s+data')
            $item.IsManual      = $false
            $item.Selected      = $false
            $script:policyPreviewItems.Add($item)
        }
        Set-PolicyGridSource $script:policyPreviewItems

        $totalRecs     = $preview.TotalCount
        $filteredCount = @($filteredRecs).Count
        $policyCount   = $preview.PolicyCount
        $autoSkipCount = @($script:policyPreviewItems | Where-Object { $_.IsAutoCreationUnsupported }).Count
        $txtPolicyInfo.Text = "$totalRecs total recommendations | $filteredCount after $profileLevel/$originFilter filter | $policyCount policies | $autoSkipCount not auto-creatable"
        Update-PolicySelectionCount
        Apply-PolicyTextFilter
    } catch {
        $txtPolicyInfo.Text = "Error: $($_.Exception.Message) at line $($_.InvocationInfo.ScriptLineNumber)"
        $script:policyPreviewItems = $null
        Set-PolicyGridSource $null
    }
}

function Get-PolicyProfileType {
    if (-not $cmbPolicyProfileType -or -not $cmbPolicyProfileType.SelectedItem) {
        return 'SettingsCatalog'
    }
    return [string](($cmbPolicyProfileType.SelectedItem).Tag)
}

function Update-PolicyProfileTypeUi {
    $platform = ''
    if ($cmbPolicyPlatform -and $cmbPolicyPlatform.SelectedItem) {
        $platform = [string](($cmbPolicyPlatform.SelectedItem).Tag)
    }

    $isIos = $platform -eq 'iOS'
    if ($txtPolicyProfileTypeLabel) {
        $txtPolicyProfileTypeLabel.Visibility = if ($isIos) { 'Visible' } else { 'Collapsed' }
    }
    if ($cmbPolicyProfileType) {
        $cmbPolicyProfileType.Visibility = if ($isIos) { 'Visible' } else { 'Collapsed' }
        if (-not $isIos) {
            foreach ($it in $cmbPolicyProfileType.Items) {
                if ($it.Tag -eq 'SettingsCatalog') {
                    $cmbPolicyProfileType.SelectedItem = $it
                    break
                }
            }
        }
    }
}

#endregion Selection Tab Helpers

$btnSelectAll.Add_Click({
    $items = $script:policyPreviewItems
    if ($items) { foreach ($i in $items) { if (-not $i.IsManual) { $i.Selected = $true } } }
    $dgPolicies.Dispatcher.BeginInvoke([Action]{
        Update-PolicySelectionCount
        Sync-PolicyGroupCheckboxes
    })
})

$btnSelectNone.Add_Click({
    $items = $script:policyPreviewItems
    if ($items) { foreach ($i in $items) { $i.Selected = $false } }
    $dgPolicies.Dispatcher.BeginInvoke([Action]{
        Update-PolicySelectionCount
        Sync-PolicyGroupCheckboxes
    })
})

# Handle group-level checkbox click via routed ButtonBase.Click (fires after toggle)
$dgPolicies.AddHandler(
    [System.Windows.Controls.Primitives.ButtonBase]::ClickEvent,
    [System.Windows.RoutedEventHandler]{
        param($eventSource, $e)
        $cb = Get-CheckBoxFromEventSource -Source $e.OriginalSource
        if ($cb -is [System.Windows.Controls.CheckBox] -and $cb.Tag) {
            $policyName = $cb.Tag
            $isChecked  = $cb.IsChecked
            $items = $script:policyPreviewItems
            if ($items) {
                foreach ($i in $items) {
                    if ($i.PolicyName -eq $policyName) {
                        if ($isChecked -and -not $i.IsManual) { $i.Selected = $true }
                        elseif (-not $isChecked) { $i.Selected = $false }
                    }
                }
                $dgPolicies.Dispatcher.BeginInvoke([Action]{
                    try { $dgPolicies.CommitEdit([System.Windows.Controls.DataGridEditingUnit]::Row, $true) } catch {}
                    Update-PolicySelectionCount
                })
            }
        }
    }
)

$toggleSyncHandler = [System.Windows.RoutedEventHandler]{
    param($eventSource, $e)
    $cb = Get-CheckBoxFromEventSource -Source $e.OriginalSource
    if ($cb -and -not $cb.Tag) {
        $dgPolicies.Dispatcher.BeginInvoke([Action]{
            try { $dgPolicies.CommitEdit([System.Windows.Controls.DataGridEditingUnit]::Row, $true) } catch {}
            Update-PolicySelectionCount
            Sync-PolicyGroupCheckboxes
        }) | Out-Null
    }
}
$dgPolicies.AddHandler([System.Windows.Controls.Primitives.ToggleButton]::CheckedEvent, $toggleSyncHandler)
$dgPolicies.AddHandler([System.Windows.Controls.Primitives.ToggleButton]::UncheckedEvent, $toggleSyncHandler)

$txtPolicyFilter.Add_TextChanged({
    if ($txtPolicyFilter.Text.Length -gt 0) {
        $txtPolicyFilterPlaceholder.Visibility = 'Collapsed'
    } else {
        $txtPolicyFilterPlaceholder.Visibility = 'Visible'
    }
    Apply-PolicyTextFilter
})

$dgPolicies.Add_CellEditEnding({
    Update-PolicySelectionCount
    $dgPolicies.Dispatcher.BeginInvoke([Action]{ Sync-PolicyGroupCheckboxes }) | Out-Null
})

$cmbPolicyProfile.Add_SelectionChanged({
    if ($txtPolicyMdPath.Text.Trim() -and (Test-Path $txtPolicyMdPath.Text.Trim())) {
        Update-PolicyPreview
    }
})

$cmbPolicyOrigin.Add_SelectionChanged({
    if ($txtPolicyMdPath.Text.Trim() -and (Test-Path $txtPolicyMdPath.Text.Trim())) {
        Update-PolicyPreview
    }
})

$cmbPolicyPlatform.Add_SelectionChanged({
    Update-PolicyProfileTypeUi
})

Update-PolicyProfileTypeUi

$btnGraphConnect.Add_Click({
    try {
        $ctx = Get-MgContext -ErrorAction SilentlyContinue
        if ($ctx -and $ctx.TenantId) {
            Disconnect-MgGraph -ErrorAction SilentlyContinue
            Update-PolicyTenantInfo
        } else {
            Connect-MgGraph -Scopes 'DeviceManagementConfiguration.ReadWrite.All','Organization.Read.All' -NoWelcome -ErrorAction Stop | Out-Null
            Update-PolicyTenantInfo
        }
    } catch {
        [System.Windows.MessageBox]::Show("Graph error: $($_.Exception.Message)", 'Error', 'OK', 'Error')
    }
})

$btnPolicyBrowseMd.Add_Click({
    $dlg = [Microsoft.Win32.OpenFileDialog]::new()
    $dlg.Filter = 'Markdown Files (*.md)|*.md|All Files (*.*)|*.*'
    $dlg.Title  = 'Select CIS Benchmark Markdown File'
    if ($dlg.ShowDialog()) {
        $txtPolicyMdPath.Text = $dlg.FileName
        Write-PolicyLog "Selected: $($dlg.FileName)"
        Update-PolicyPreview
    }
})

$btnPolicyExport.Add_Click({
    $mdPath = $txtPolicyMdPath.Text.Trim()
    if (-not $mdPath) { [System.Windows.MessageBox]::Show('Select a Markdown file.', 'Missing', 'OK', 'Warning'); return }
    if (-not (Test-Path $mdPath -PathType Leaf)) { [System.Windows.MessageBox]::Show("File not found.", 'Error', 'OK', 'Error'); return }

    $profileLevel = ($cmbPolicyProfile.SelectedItem).Tag
    $platform     = ($cmbPolicyPlatform.SelectedItem).Tag
    $profileType  = Get-PolicyProfileType

    # Locate the original policy creator script
    $creatorScript = Join-Path (Split-Path $script:thisScriptPath -Parent) 'Features\Invoke-IntuneCISPolicyCreator.ps1'
    if (-not (Test-Path $creatorScript)) {
        [System.Windows.MessageBox]::Show("Policy Creator script not found at:`n$creatorScript", 'Missing Script', 'OK', 'Error')
        return
    }

    $arguments = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $creatorScript, '-MdFilePath', $mdPath, '-ProfileLevel', $profileLevel, '-Platform', $platform, '-ProfileType', $profileType, '-ExportOnly')

    # Apply selection filters from the DataGrid if available
    if ($script:policyPreviewItems -and @($script:policyPreviewItems).Count -gt 0) {
        $selectedItems = @($script:policyPreviewItems | Where-Object { $_.Selected -and -not $_.IsManual -and -not $_.IsAutoCreationUnsupported })
        if ($selectedItems.Count -eq 0) {
            [System.Windows.MessageBox]::Show('No settings selected in the Selection tab.', 'No Selection', 'OK', 'Warning'); return
        }
        $selectedPolicies = @($selectedItems | ForEach-Object { $_.PolicyName } | Where-Object { $_ -and $_ -ne '-' } | Sort-Object -Unique)
        $selectedRefs     = @($selectedItems | Where-Object { $_.CISRef -and $_.CISRef -ne '-' } | ForEach-Object { $_.CISRef })
        $selectedRecTexts = @($selectedItems | Where-Object { -not $_.CISRef -or $_.CISRef -eq '-' } | ForEach-Object { $_.Recommendation } | Where-Object { $_ })
        $policyFilterStr = ($selectedPolicies -join '|||')
        $settingFilterStr = ($selectedRefs -join '|||')
        $recFilterStr = ($selectedRecTexts -join '|||')
        if ($policyFilterStr) {
            $arguments += '-PolicyFilter'
            $arguments += $policyFilterStr
        }
        if ($settingFilterStr) {
            $arguments += '-SettingFilter'
            $arguments += $settingFilterStr
        }
        if ($recFilterStr) {
            $arguments += '-RecommendationFilter'
            $arguments += $recFilterStr
        }
    }

    $expectedJsonPath = [System.IO.Path]::ChangeExtension($mdPath, '.policies.json')

    $txtPolicyLog.Text = ''
    $btnPolicyExport.IsEnabled = $false
    $btnPolicyCreate.IsEnabled = $false
    $policyProgressBar.IsIndeterminate = $true
    $txtPolicyProgress.Text = 'Exporting...'
    $policySubTabs.SelectedIndex = 1  # Switch to Output tab
    Write-PolicyLog "Exporting... (Profile: $profileLevel, Platform: $platform, ProfileType: $profileType)"

    $runspace = [runspacefactory]::CreateRunspace()
    $runspace.ApartmentState = 'STA'
    $runspace.Open()
    $runspace.SessionStateProxy.SetVariable('arguments', $arguments)
    $runspace.SessionStateProxy.SetVariable('txtPolicyLog', $txtPolicyLog)
    $runspace.SessionStateProxy.SetVariable('policyProgressBar', $policyProgressBar)
    $runspace.SessionStateProxy.SetVariable('txtPolicyProgress', $txtPolicyProgress)
    $runspace.SessionStateProxy.SetVariable('btnPolicyExport', $btnPolicyExport)
    $runspace.SessionStateProxy.SetVariable('btnPolicyCreate', $btnPolicyCreate)
    $runspace.SessionStateProxy.SetVariable('expectedJsonPath', $expectedJsonPath)

    $ps = [powershell]::Create()
    $ps.Runspace = $runspace
    $ps.AddScript({
        function Write-Log { param([string]$M); $txtPolicyLog.Dispatcher.Invoke([Action]{ $txtPolicyLog.AppendText("$M`r`n"); $txtPolicyLog.ScrollToEnd() }) }
        try {
            $process = [System.Diagnostics.Process]::new()
            $process.StartInfo.FileName               = 'powershell.exe'
            $process.StartInfo.Arguments              = ($arguments | ForEach-Object { if ($_ -match '\s') { "`"$_`"" } else { $_ } }) -join ' '
            $process.StartInfo.UseShellExecute        = $false
            $process.StartInfo.RedirectStandardOutput = $true
            $process.StartInfo.RedirectStandardError  = $true
            $process.StartInfo.CreateNoWindow         = $true
            $process.StartInfo.StandardOutputEncoding = [System.Text.Encoding]::UTF8
            $process.StartInfo.StandardErrorEncoding  = [System.Text.Encoding]::UTF8
            $process.Start() | Out-Null
            while (-not $process.StandardOutput.EndOfStream) { $line = $process.StandardOutput.ReadLine(); if ($line) { Write-Log $line } }
            $stderr = $process.StandardError.ReadToEnd()
            if ($stderr) { foreach ($el in ($stderr -split "`n")) { if ($el.Trim()) { Write-Log "ERROR: $($el.Trim())" } } }
            $process.WaitForExit()
            if ($process.ExitCode -eq 0) {
                Write-Log "`n=== Export completed ==="
                if (Test-Path $expectedJsonPath) { Write-Log "JSON: $expectedJsonPath" }
                $policyProgressBar.Dispatcher.Invoke([Action]{ $policyProgressBar.IsIndeterminate = $false; $policyProgressBar.Value = 100 })
                $txtPolicyProgress.Dispatcher.Invoke([Action]{ $txtPolicyProgress.Text = 'Done' })
            } else {
                Write-Log "`n=== Export failed ==="
                $policyProgressBar.Dispatcher.Invoke([Action]{ $policyProgressBar.IsIndeterminate = $false; $policyProgressBar.Value = 0 })
                $txtPolicyProgress.Dispatcher.Invoke([Action]{ $txtPolicyProgress.Text = 'Error' })
            }
        } catch { Write-Log "ERROR: $($_.Exception.Message)"; $policyProgressBar.Dispatcher.Invoke([Action]{ $policyProgressBar.IsIndeterminate = $false }) }
        finally { $btnPolicyExport.Dispatcher.Invoke([Action]{ $btnPolicyExport.IsEnabled = $true; $btnPolicyCreate.IsEnabled = $true }) }
    }) | Out-Null
    $ps.BeginInvoke() | Out-Null
})

$btnPolicyCreate.Add_Click({
    $mdPath = $txtPolicyMdPath.Text.Trim()
    if (-not $mdPath) { [System.Windows.MessageBox]::Show('Select a Markdown file.', 'Missing', 'OK', 'Warning'); return }
    if (-not (Test-Path $mdPath -PathType Leaf)) { [System.Windows.MessageBox]::Show("File not found.", 'Error', 'OK', 'Error'); return }

    $profileLevel   = ($cmbPolicyProfile.SelectedItem).Tag
    $platform       = ($cmbPolicyPlatform.SelectedItem).Tag
    $profileType    = Get-PolicyProfileType

    $selectedItemsAll = @()
    $selectedItemsCreatable = @()
    $selectedItemsAutoUnsupported = @()
    $selectedItemsCustomerInfoRequired = @()

    if ($script:policyPreviewItems -and @($script:policyPreviewItems).Count -gt 0) {
        $selectedItemsAll = @($script:policyPreviewItems | Where-Object { $_.Selected -and -not $_.IsManual })
        if ($selectedItemsAll.Count -eq 0) {
            [System.Windows.MessageBox]::Show('No settings selected in the Selection tab.', 'No Selection', 'OK', 'Warning'); return
        }

        $selectedItemsCreatable = @($selectedItemsAll | Where-Object { -not $_.IsAutoCreationUnsupported -and -not $_.IsCustomerInfoRequired })
        $selectedItemsAutoUnsupported = @($selectedItemsAll | Where-Object { $_.IsAutoCreationUnsupported })
        $selectedItemsCustomerInfoRequired = @($selectedItemsAll | Where-Object { $_.IsCustomerInfoRequired })

        if ($selectedItemsCreatable.Count -eq 0 -and ($selectedItemsAutoUnsupported.Count -gt 0 -or $selectedItemsCustomerInfoRequired.Count -gt 0)) {
            $txtPolicyLog.Text = ''
            $policySubTabs.SelectedIndex = 1
            $policyProgressBar.IsIndeterminate = $false
            $policyProgressBar.Value = 100
            $txtPolicyProgress.Text = 'Report generated (no creatable settings selected)'
            Write-PolicyLog "No creatable settings selected. Generating report for selected non-creatable settings..."

            try {
                $reportRows = [System.Collections.Generic.List[pscustomobject]]::new()
                foreach ($ai in $selectedItemsAutoUnsupported) {
                    $reportRows.Add([pscustomobject]@{
                        PolicyName   = $ai.PolicyName
                        CISRef       = $ai.CISRef
                        SettingName  = $ai.SettingName
                        SettingValue = $ai.SettingValue
                        Profile      = $ai.Profile
                        Outcome      = 'AutoCreationUnsupported'
                        Detail       = 'Cannot be automatically created in Intune (from Notes column)'
                    })
                }
                foreach ($ci in $selectedItemsCustomerInfoRequired) {
                    $reportRows.Add([pscustomobject]@{
                        PolicyName   = $ci.PolicyName
                        CISRef       = $ci.CISRef
                        SettingName  = $ci.SettingName
                        SettingValue = $ci.SettingValue
                        Profile      = $ci.Profile
                        Outcome      = 'CustomerInformationRequired'
                        Detail       = 'Customer information required (from Notes column)'
                    })
                }

                $htmlReportPath = [System.IO.Path]::ChangeExtension($mdPath, '.creation-report.html')
                Build-HtmlReport -OutputPath $htmlReportPath -MdFileName ([System.IO.Path]::GetFileName($mdPath)) -ProfileLevel $profileLevel -Platform $platform -ReportRows @($reportRows)
                $btnOpenReport.IsEnabled = $true
                $btnOpenReport.Tag = $htmlReportPath
                Write-PolicyLog "HTML report: $htmlReportPath"
            } catch {
                Write-PolicyLog "Could not generate HTML report: $($_.Exception.Message)"
                $txtPolicyProgress.Text = 'Error'
                $policyProgressBar.Value = 0
            }
            return
        }
    }

    $confirm = [System.Windows.MessageBox]::Show(
        "This will connect to Microsoft Graph and create Intune policies.`n`nProfile: $profileLevel`nPlatform: $platform`nProfile Type: $profileType`nSource: $(Split-Path $mdPath -Leaf)`n`nContinue?",
        'Confirm', 'YesNo', 'Question')
    if ($confirm -ne 'Yes') { return }

    # Locate the original policy creator script
    $creatorScript = Join-Path (Split-Path $script:thisScriptPath -Parent) 'Features\Invoke-IntuneCISPolicyCreator.ps1'
    if (-not (Test-Path $creatorScript)) {
        [System.Windows.MessageBox]::Show("Policy Creator script not found at:`n$creatorScript", 'Missing Script', 'OK', 'Error')
        return
    }

    $arguments = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $creatorScript, '-MdFilePath', $mdPath, '-ProfileLevel', $profileLevel, '-Platform', $platform, '-ProfileType', $profileType)

    # Apply selection filters from the DataGrid if available
    if ($script:policyPreviewItems -and @($script:policyPreviewItems).Count -gt 0) {
        $selectedPolicies = @($selectedItemsCreatable | ForEach-Object { $_.PolicyName } | Where-Object { $_ -and $_ -ne '-' } | Sort-Object -Unique)
        $selectedRefs     = @($selectedItemsCreatable | Where-Object { $_.CISRef -and $_.CISRef -ne '-' } | ForEach-Object { $_.CISRef })
        $selectedRecTexts = @($selectedItemsCreatable | Where-Object { -not $_.CISRef -or $_.CISRef -eq '-' } | ForEach-Object { $_.Recommendation } | Where-Object { $_ })
        $policyFilterStr = ($selectedPolicies -join '|||')
        $settingFilterStr = ($selectedRefs -join '|||')
        $recFilterStr = ($selectedRecTexts -join '|||')
        if ($policyFilterStr) {
            $arguments += '-PolicyFilter'
            $arguments += $policyFilterStr
        }
        if ($settingFilterStr) {
            $arguments += '-SettingFilter'
            $arguments += $settingFilterStr
        }
        if ($recFilterStr) {
            $arguments += '-RecommendationFilter'
            $arguments += $recFilterStr
        }
    }

    $txtPolicyLog.Text = ''
    $btnPolicyExport.IsEnabled = $false
    $btnPolicyCreate.IsEnabled = $false
    $policyProgressBar.IsIndeterminate = $true
    $txtPolicyProgress.Text = 'Creating policies...'
    $policySubTabs.SelectedIndex = 1  # Switch to Output tab
    Write-PolicyLog "Creating policies... (Profile: $profileLevel, Platform: $platform, ProfileType: $profileType)"

    $runspace = [runspacefactory]::CreateRunspace()
    $runspace.ApartmentState = 'STA'
    $runspace.Open()
    $runspace.SessionStateProxy.SetVariable('arguments', $arguments)
    $runspace.SessionStateProxy.SetVariable('txtPolicyLog', $txtPolicyLog)
    $runspace.SessionStateProxy.SetVariable('policyProgressBar', $policyProgressBar)
    $runspace.SessionStateProxy.SetVariable('txtPolicyProgress', $txtPolicyProgress)
    $runspace.SessionStateProxy.SetVariable('btnPolicyExport', $btnPolicyExport)
    $runspace.SessionStateProxy.SetVariable('btnPolicyCreate', $btnPolicyCreate)
    $runspace.SessionStateProxy.SetVariable('btnOpenReport', $btnOpenReport)
    $runspace.SessionStateProxy.SetVariable('window', $window)
    $runspace.SessionStateProxy.SetVariable('buildHtmlReportSb', (Get-Item Function:Build-HtmlReport).ScriptBlock)
    # Data needed for report generation
    $selectedItemsForReport = @($script:policyPreviewItems | Where-Object { $_.Selected -and -not $_.IsManual -and -not $_.IsAutoCreationUnsupported -and -not $_.IsCustomerInfoRequired })
    $manualItemsForReport   = @($script:policyPreviewItems | Where-Object { $_.IsManual })
    $autoUnsupportedItemsForReport = @($script:policyPreviewItems | Where-Object { $_.Selected -and $_.IsAutoCreationUnsupported })
    $customerInfoItemsForReport = @($script:policyPreviewItems | Where-Object { $_.Selected -and $_.IsCustomerInfoRequired })
    $runspace.SessionStateProxy.SetVariable('selectedItemsForReport', $selectedItemsForReport)
    $runspace.SessionStateProxy.SetVariable('manualItemsForReport', $manualItemsForReport)
    $runspace.SessionStateProxy.SetVariable('autoUnsupportedItemsForReport', $autoUnsupportedItemsForReport)
    $runspace.SessionStateProxy.SetVariable('customerInfoItemsForReport', $customerInfoItemsForReport)
    $runspace.SessionStateProxy.SetVariable('mdPathForReport', $mdPath)
    $runspace.SessionStateProxy.SetVariable('profileLevelForReport', $profileLevel)
    $runspace.SessionStateProxy.SetVariable('platformForReport', $platform)

    $ps = [powershell]::Create()
    $ps.Runspace = $runspace
    $ps.AddScript({
        function Write-Log { param([string]$M); $txtPolicyLog.Dispatcher.Invoke([Action]{ $txtPolicyLog.AppendText("$M`r`n"); $txtPolicyLog.ScrollToEnd() }) }
        try {
            $process = [System.Diagnostics.Process]::new()
            $process.StartInfo.FileName               = 'powershell.exe'
            $process.StartInfo.Arguments              = ($arguments | ForEach-Object { if ($_ -match '\s') { "`"$_`"" } else { $_ } }) -join ' '
            $process.StartInfo.UseShellExecute        = $false
            $process.StartInfo.RedirectStandardOutput = $true
            $process.StartInfo.RedirectStandardError  = $true
            $process.StartInfo.CreateNoWindow         = $true
            $process.StartInfo.StandardOutputEncoding = [System.Text.Encoding]::UTF8
            $process.StartInfo.StandardErrorEncoding  = [System.Text.Encoding]::UTF8
            $process.Start() | Out-Null
            $allOutputLines = [System.Collections.Generic.List[string]]::new()
            while (-not $process.StandardOutput.EndOfStream) {
                $line = $process.StandardOutput.ReadLine()
                if ($line) {
                    Write-Log $line
                    $allOutputLines.Add($line)
                    if ($line -match 'Connecting to Microsoft Graph') { $txtPolicyProgress.Dispatcher.Invoke([Action]{ $txtPolicyProgress.Text = 'Waiting for sign-in...' }) }
                    elseif ($line -match 'Connected to Microsoft Graph') { $txtPolicyProgress.Dispatcher.Invoke([Action]{ $txtPolicyProgress.Text = 'Creating...' }); $policyProgressBar.Dispatcher.Invoke([Action]{ $policyProgressBar.IsIndeterminate = $false; $policyProgressBar.Value = 0 }) }
                    elseif ($line -match 'CREATED:|FAILED:') { }
                }
            }
            $stderr = $process.StandardError.ReadToEnd()
            if ($stderr) { foreach ($el in ($stderr -split "`n")) { if ($el.Trim()) { Write-Log "ERROR: $($el.Trim())" } } }
            $process.WaitForExit()
            if ($process.ExitCode -eq 0) {
                Write-Log "`n=== Policy creation completed ==="
                $policyProgressBar.Dispatcher.Invoke([Action]{ $policyProgressBar.IsIndeterminate = $false; $policyProgressBar.Value = 100 })
                $txtPolicyProgress.Dispatcher.Invoke([Action]{ $txtPolicyProgress.Text = 'Completed' })

                # Build HTML report
                try {
                    $mappedSettings = [System.Collections.Generic.HashSet[string]]::new()
                    $notFoundSettings = [System.Collections.Generic.HashSet[string]]::new()
                    $mismatchSettings = [System.Collections.Generic.HashSet[string]]::new()
                    $customerInfoSettings = [System.Collections.Generic.HashSet[string]]::new()
                    foreach ($l in $allOutputLines) {
                        if ($l -match 'Mapped:\s+(.+?)\s+->') { $mappedSettings.Add($Matches[1].Trim()) | Out-Null }
                        elseif ($l -match 'NOT FOUND IN INTUNE:\s+(.+?)\s+\(CIS') { $notFoundSettings.Add($Matches[1].Trim()) | Out-Null }
                        elseif ($l -match 'VALUE MISMATCH:\s+(.+?)\s+-') { $mismatchSettings.Add($Matches[1].Trim()) | Out-Null }
                        elseif ($l -match 'CUSTOMER INFO MISSING:\s+(.+?)\s+\(CIS') { $customerInfoSettings.Add($Matches[1].Trim()) | Out-Null }
                    }
                    $reportRows = [System.Collections.Generic.List[pscustomobject]]::new()
                    foreach ($si in $selectedItemsForReport) {
                        $outcome = if ($customerInfoSettings.Contains($si.SettingName)) { 'CustomerInfoMissing' }
                                   elseif ($notFoundSettings.Contains($si.SettingName)) { 'NotFound' }
                                   elseif ($mismatchSettings.Contains($si.SettingName)) { 'ValueMismatch' }
                                   elseif ($mappedSettings.Contains($si.SettingName))   { 'Created' }
                                   else { 'Created' }
                        $detail = if ($outcome -eq 'CustomerInfoMissing') { 'Requires customer-provided values' } else { '' }
                        $reportRows.Add([pscustomobject]@{ PolicyName = $si.PolicyName; CISRef = $si.CISRef; SettingName = $si.SettingName; SettingValue = $si.SettingValue; Profile = $si.Profile; Outcome = $outcome; Detail = $detail })
                    }
                    foreach ($mi in $manualItemsForReport) {
                        $reportRows.Add([pscustomobject]@{ PolicyName = $mi.PolicyName; CISRef = $mi.CISRef; SettingName = $mi.SettingName; SettingValue = $mi.SettingValue; Profile = $mi.Profile; Outcome = 'Manual'; Detail = 'Customer-specific data required' })
                    }
                    foreach ($ai in $autoUnsupportedItemsForReport) {
                        $reportRows.Add([pscustomobject]@{ PolicyName = $ai.PolicyName; CISRef = $ai.CISRef; SettingName = $ai.SettingName; SettingValue = $ai.SettingValue; Profile = $ai.Profile; Outcome = 'AutoCreationUnsupported'; Detail = 'Cannot be automatically created in Intune (from Notes column)' })
                    }
                    foreach ($ci in $customerInfoItemsForReport) {
                        $reportRows.Add([pscustomobject]@{ PolicyName = $ci.PolicyName; CISRef = $ci.CISRef; SettingName = $ci.SettingName; SettingValue = $ci.SettingValue; Profile = $ci.Profile; Outcome = 'CustomerInformationRequired'; Detail = 'Customer information required (from Notes column)' })
                    }
                    $htmlReportPath = [System.IO.Path]::ChangeExtension($mdPathForReport, '.creation-report.html')
                    $reportRowsCopy = @($reportRows)
                    $mdFileNameLocal = [System.IO.Path]::GetFileName($mdPathForReport)
                    $htmlPathLocal = $htmlReportPath
                    $pfLocal = $profileLevelForReport
                    $platLocal = $platformForReport
                    $window.Dispatcher.Invoke([Action]{
                        & $buildHtmlReportSb -OutputPath $htmlPathLocal -MdFileName $mdFileNameLocal -ProfileLevel $pfLocal -Platform $platLocal -ReportRows $reportRowsCopy
                    })
                    Write-Log "HTML report: $htmlReportPath"
                    $btnOpenReport.Dispatcher.Invoke([Action]{ $btnOpenReport.IsEnabled = $true; $btnOpenReport.Tag = $htmlReportPath })
                } catch {
                    Write-Log "Could not generate HTML report: $($_.Exception.Message)"
                }
            } else {
                Write-Log "`n=== Policy creation failed ==="
                $policyProgressBar.Dispatcher.Invoke([Action]{ $policyProgressBar.IsIndeterminate = $false; $policyProgressBar.Value = 0 })
                $txtPolicyProgress.Dispatcher.Invoke([Action]{ $txtPolicyProgress.Text = 'Error' })
            }
        } catch { Write-Log "ERROR: $($_.Exception.Message)"; $policyProgressBar.Dispatcher.Invoke([Action]{ $policyProgressBar.IsIndeterminate = $false }) }
        finally { $btnPolicyExport.Dispatcher.Invoke([Action]{ $btnPolicyExport.IsEnabled = $true; $btnPolicyCreate.IsEnabled = $true }) }
    }) | Out-Null
    $ps.BeginInvoke() | Out-Null
})

# Open Report button
$btnOpenReport.Add_Click({
    $reportPath = $btnOpenReport.Tag
    if ($reportPath -and (Test-Path $reportPath)) {
        Start-Process -FilePath $reportPath
    } else {
        [System.Windows.MessageBox]::Show('Report file not found. Run a policy creation first.', 'No Report', 'OK', 'Warning')
    }
})

# ══════════════════════════════════════════════════════════════════════════════
#  TAB 4 — INTUNE BASELINE REPORT EVENT HANDLERS
# ══════════════════════════════════════════════════════════════════════════════

$btnProfileReportGraphConnect.Add_Click({
    try {
        $requiredTenant = Get-ResolverTenantOverride
        $ctx = Get-MgContext -ErrorAction SilentlyContinue
        if ($ctx -and $ctx.TenantId -and (-not $requiredTenant -or ([string]$ctx.TenantId -ieq [string]$requiredTenant))) {
            Disconnect-MgGraph -ErrorAction SilentlyContinue
            Update-PolicyTenantInfo
            Update-ProfileReportTenantInfo | Out-Null
        } else {
            Connect-BaselineReadGraph -TenantId $requiredTenant
            Update-PolicyTenantInfo
            Update-ProfileReportTenantInfo | Out-Null
        }
    } catch {
        [System.Windows.MessageBox]::Show("Graph error: $($_.Exception.Message)", 'Error', 'OK', 'Error') | Out-Null
    }
})

$btnProfileReportLoad.Add_Click({
    if (-not (Ensure-ProfileReportGraphConnection)) { return }

    $btnProfileReportLoad.IsEnabled = $false
    $btnProfileReportGenerate.IsEnabled = $false
    $profileReportProgressBar.IsIndeterminate = $true
    $txtProfileReportProgress.Text = 'Loading profiles...'
    $txtProfileReportLog.Text = ''

    Write-ProfileReportLog 'Loading all configuration policies from Intune...'

    $runspace = [runspacefactory]::CreateRunspace()
    $runspace.ApartmentState = 'STA'
    $runspace.Open()

    $ps = [powershell]::Create()
    $ps.Runspace = $runspace
    $ps.AddScript({ param($modulePath) Import-Module $modulePath -DisableNameChecking; Get-BaselineIntuneProfiles }).AddArgument((Join-Path $PSScriptRoot 'Shared\BaselineToolkit.Common.psm1')) | Out-Null

    $startTime = Get-Date
    $handle = $ps.BeginInvoke()

    $timer = New-Object System.Windows.Threading.DispatcherTimer
    $timer.Interval = [timespan]::FromMilliseconds(400)
    $timer.Add_Tick({
        if ($handle.IsCompleted) {
            $timer.Stop()
            try {
                $result = $ps.EndInvoke($handle)
                $rawCount = 0
                $mapped = @()
                if ($result) {
                    $resultObj = @($result)[0]
                    if ($resultObj) {
                        try { $rawCount = [int]$resultObj.RawCount } catch {}
                        try { $mapped = @($resultObj.Mapped) } catch {}
                    }
                }

                # Rehydrate into local PSCustomObjects to guarantee stable DataGrid bindings.
                $normalized = @()
                foreach ($m in $mapped) {
                    $idv = Get-SafeObjectPropertyValue -Object $m -Name 'Id'
                    if (-not $idv) { $idv = Get-SafeObjectPropertyValue -Object $m -Name 'id' }
                    $namev = Get-SafeObjectPropertyValue -Object $m -Name 'Name'
                    if (-not $namev) { $namev = Get-SafeObjectPropertyValue -Object $m -Name 'name' }
                    $platv = Get-SafeObjectPropertyValue -Object $m -Name 'Platforms'
                    if (-not $platv) { $platv = Get-SafeObjectPropertyValue -Object $m -Name 'platforms' }
                    $lastv = Get-SafeObjectPropertyValue -Object $m -Name 'LastModifiedDateTime'
                    if (-not $lastv) { $lastv = Get-SafeObjectPropertyValue -Object $m -Name 'lastModifiedDateTime' }
                    $kindv = Get-SafeObjectPropertyValue -Object $m -Name 'PolicyKind'
                    $typev = Get-SafeObjectPropertyValue -Object $m -Name 'OdataType'
                    if (-not $idv -or -not $namev) { continue }
                    $normalized += [pscustomobject]@{
                        Selected = $false
                        Id = [string]$idv
                        Name = [string]$namev
                        Platforms = if ($platv) { [string]$platv } else { '-' }
                        LastModifiedDateTime = if ($lastv) { [string]$lastv } else { '-' }
                        PolicyKind = if ($kindv) { [string]$kindv } else { 'SettingsCatalog' }
                        OdataType  = if ($typev) { [string]$typev } else { '' }
                    }
                }

                if (@($normalized).Count -eq 0 -and @($mapped).Count -gt 0) {
                    Write-ProfileReportLog 'Primary remap produced 0 rows. Applying fallback map...'
                    foreach ($m in $mapped) {
                        $idv = Get-SafeObjectPropertyValue -Object $m -Name 'Id'
                        if (-not $idv) { $idv = Get-SafeObjectPropertyValue -Object $m -Name 'id' }
                        $namev = Get-SafeObjectPropertyValue -Object $m -Name 'Name'
                        if (-not $namev) { $namev = Get-SafeObjectPropertyValue -Object $m -Name 'name' }
                        if (-not $idv -or -not $namev) { continue }

                        $platv = Get-SafeObjectPropertyValue -Object $m -Name 'Platforms'
                        if (-not $platv) { $platv = Get-SafeObjectPropertyValue -Object $m -Name 'platforms' }
                        $lastv = Get-SafeObjectPropertyValue -Object $m -Name 'LastModifiedDateTime'
                        if (-not $lastv) { $lastv = Get-SafeObjectPropertyValue -Object $m -Name 'lastModifiedDateTime' }
                        $kindv = Get-SafeObjectPropertyValue -Object $m -Name 'PolicyKind'
                        $typev = Get-SafeObjectPropertyValue -Object $m -Name 'OdataType'

                        $normalized += [pscustomobject]@{
                            Selected = $false
                            Id = [string]$idv
                            Name = [string]$namev
                            Platforms = if ($platv) { [string]$platv } else { '-' }
                            LastModifiedDateTime = if ($lastv) { [string]$lastv } else { '-' }
                            PolicyKind = if ($kindv) { [string]$kindv } else { 'SettingsCatalog' }
                            OdataType  = if ($typev) { [string]$typev } else { '' }
                        }
                    }
                }

                $script:intuneReportProfilesAll = $normalized

                if (-not $script:intuneReportProfilesCollection) {
                    $script:intuneReportProfilesCollection = [System.Collections.ObjectModel.ObservableCollection[object]]::new()
                    $dgProfileReportProfiles.ItemsSource = $script:intuneReportProfilesCollection
                }

                $script:intuneReportProfilesCollection.Clear()
                foreach ($p in $script:intuneReportProfilesAll) {
                    $script:intuneReportProfilesCollection.Add($p)
                }
                if (-not $script:intuneReportProfilesView) {
                    $script:intuneReportProfilesView = [System.Windows.Data.CollectionViewSource]::GetDefaultView($script:intuneReportProfilesCollection)
                }

                Write-ProfileReportLog "Raw objects received: $rawCount"
                Write-ProfileReportLog "Mapped objects received: $($mapped.Count)"
                Write-ProfileReportLog "Profiles with id+name: $($script:intuneReportProfilesAll.Count)"

                Apply-ProfileReportFilter -FilterText $txtProfileReportFilter.Text
                if (@($script:intuneReportProfilesAll).Count -gt 0 -and @($script:intuneReportProfilesFiltered).Count -eq 0) {
                    Write-ProfileReportLog "Filter returned 0 matches after load. Clearing filter and showing all profiles."
                    $txtProfileReportFilter.Text = ''
                    Apply-ProfileReportFilter -FilterText ''
                }
                Write-ProfileReportLog "DataGrid rows after bind: $($script:intuneReportProfilesFiltered.Count)"
                Write-ProfileReportLog "Loaded $($script:intuneReportProfilesAll.Count) profiles."
                $profileReportProgressBar.IsIndeterminate = $false
                $profileReportProgressBar.Value = 100
                $txtProfileReportProgress.Text = 'Done'
            } catch {
                $profileReportProgressBar.IsIndeterminate = $false
                $profileReportProgressBar.Value = 0
                $txtProfileReportProgress.Text = 'Error'
                Write-ProfileReportLog "ERROR: $($_.Exception.Message)"
                [System.Windows.MessageBox]::Show("Failed to load profiles: $($_.Exception.Message)", 'Error', 'OK', 'Error') | Out-Null
            } finally {
                $ps.Dispose()
                $runspace.Close()
                $runspace.Dispose()
                $btnProfileReportLoad.IsEnabled = $true
                $btnProfileReportGenerate.IsEnabled = $true
            }
        } else {
            $elapsed = [int]((Get-Date) - $startTime).TotalSeconds
            $txtProfileReportProgress.Text = "Loading... ${elapsed}s"
        }
    }.GetNewClosure())
    $timer.Start()
})

$txtProfileReportFilter.Add_TextChanged({
    Apply-ProfileReportFilter -FilterText $txtProfileReportFilter.Text
})

$btnProfileReportSelectAll.Add_Click({
    $gridSource = $dgProfileReportProfiles.ItemsSource
    if ($null -eq $gridSource) { return }
    $view = [System.Windows.Data.CollectionViewSource]::GetDefaultView($gridSource)
    $source = if ($null -ne $view) { @($view | ForEach-Object { $_ }) } else { @($gridSource | ForEach-Object { $_ }) }
    foreach ($p in $source) { $p.Selected = $true }
    $dgProfileReportProfiles.Items.Refresh()
    Update-ProfileReportSelectionCount
})

$btnProfileReportSelectNone.Add_Click({
    $gridSource = $dgProfileReportProfiles.ItemsSource
    if ($null -eq $gridSource) { return }
    $view = [System.Windows.Data.CollectionViewSource]::GetDefaultView($gridSource)
    $source = if ($null -ne $view) { @($view | ForEach-Object { $_ }) } else { @($gridSource | ForEach-Object { $_ }) }
    foreach ($p in $source) { $p.Selected = $false }
    $dgProfileReportProfiles.Items.Refresh()
    Update-ProfileReportSelectionCount
})

$dgProfileReportProfiles.Add_CellEditEnding({
    $dgProfileReportProfiles.Dispatcher.BeginInvoke([Action]{
        Update-ProfileReportSelectionCount
    }) | Out-Null
})

$btnProfileReportBrowseOutput.Add_Click({
    $dlg = [Microsoft.Win32.SaveFileDialog]::new()
    $dlg.Title = 'Save Intune Policy Report As'
    $dlg.Filter = 'HTML files (*.html)|*.html|All files (*.*)|*.*'
    $dlg.FileName = 'Intune-Policy-Report.html'
    if ($dlg.ShowDialog($window)) {
        $txtProfileReportOutputPath.Text = $dlg.FileName
    }
})

$btnProfileReportBrowseBaseline.Add_Click({
    Add-Type -AssemblyName System.Windows.Forms
    $fbd = [System.Windows.Forms.FolderBrowserDialog]::new()
    $fbd.Description = 'Select folder containing JSON Intune policy backup files'
    $fbd.ShowNewFolderButton = $false
    if ($txtProfileReportBaselineFolder.Text.Trim() -and (Test-Path $txtProfileReportBaselineFolder.Text.Trim())) {
        $fbd.SelectedPath = $txtProfileReportBaselineFolder.Text.Trim()
    }
    if ($fbd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $txtProfileReportBaselineFolder.Text = $fbd.SelectedPath
    }
})

$btnProfileReportWarmCache.Add_Click({
    if (-not (Ensure-ProfileReportGraphConnection)) { return }

    $profiles = @($script:intuneReportProfilesAll)
    if ($profiles.Count -eq 0) {
        [System.Windows.MessageBox]::Show('No profiles loaded. Click Load Profiles first.', 'No Data', 'OK', 'Warning') | Out-Null
        return
    }

    $cacheMode = Get-CacheModeFromComboBox -ComboBox $cmbProfileReportCacheMode
    $modeLabel = Get-CacheModeLabel -Mode $cacheMode
    $resolverTenantCacheKey = Get-GraphTenantCacheKey

    try {
        $btnProfileReportWarmCache.IsEnabled = $false
        $btnProfileReportLoad.IsEnabled = $false
        $btnProfileReportGenerate.IsEnabled = $false
        $txtProfileReportLog.Text = ''
        $profileReportProgressBar.IsIndeterminate = $false
        $profileReportProgressBar.Value = 0
        $txtProfileReportProgress.Text = 'Warming cache...'

        Write-ProfileReportLog "Starting cache warm-up for tenant '$resolverTenantCacheKey'"
        Write-ProfileReportLog "Resolver mode: $modeLabel"

        $warmResult = Invoke-TenantPolicyCacheWarmup `
            -Profiles $profiles `
            -TenantKey $resolverTenantCacheKey `
            -CacheMode $cacheMode `
            -ProgressCallback {
                param($done, $total)
                if ($total -gt 0) {
                    $pct = [math]::Round(($done / [double]$total) * 100)
                    $profileReportProgressBar.Value = $pct
                    $txtProfileReportProgress.Text = "Warming cache $done/$total"
                }
            } `
            -LogCallback {
                param($message)
                Write-ProfileReportLog $message
            } `
            -AuthChallengeCallback {
                param($reason)
                Write-ProfileReportLog 'Reconnecting Microsoft Graph due to authentication challenge...'
                Connect-BaselineReadGraph -TenantId (Get-ResolverTenantOverride) -ForceReconnect
            }

        $profileReportProgressBar.Value = 100
        $txtProfileReportProgress.Text = 'Warm-up completed'
        Write-ProfileReportLog "Warm-up complete. Settings catalog profiles scanned: $($warmResult.SettingsCatalogProfileCount)"
        Write-ProfileReportLog "Referenced setting IDs: $($warmResult.ReferencedDefinitionCount)"
        Write-ProfileReportLog "Resolved via Graph this run: $($warmResult.ResolveResolved) in $($warmResult.GraphCalls) batch call(s)"
        Write-ProfileReportLog "Cache now contains $($warmResult.SavedDisplayNameCount) names and $($warmResult.SavedOptionCount) options."
    } catch {
        $profileReportProgressBar.Value = 0
        $txtProfileReportProgress.Text = 'Error'
        Write-ProfileReportLog "ERROR: $($_.Exception.Message)"
        [System.Windows.MessageBox]::Show("Cache warm-up failed: $($_.Exception.Message)", 'Error', 'OK', 'Error') | Out-Null
    } finally {
        $btnProfileReportWarmCache.IsEnabled = $true
        $btnProfileReportLoad.IsEnabled = $true
        $btnProfileReportGenerate.IsEnabled = $true
    }
})

$btnProfileReportGenerate.Add_Click({
    if (-not (Ensure-ProfileReportGraphConnection)) { return }

    # Commit any in-progress checkbox edit before reading selected rows.
    try {
        $null = $dgProfileReportProfiles.CommitEdit([System.Windows.Controls.DataGridEditingUnit]::Cell, $true)
        $null = $dgProfileReportProfiles.CommitEdit([System.Windows.Controls.DataGridEditingUnit]::Row, $true)
    } catch {}

    $gridSource = $dgProfileReportProfiles.ItemsSource
    $allGridItems = if ($null -ne $gridSource) { @($gridSource | ForEach-Object { $_ }) } else { @() }
    if (($allGridItems | Measure-Object).Count -eq 0) {
        [System.Windows.MessageBox]::Show('No profiles loaded. Click Load Profiles first.', 'No Data', 'OK', 'Warning') | Out-Null
        return
    }

    $targetProfiles = @($allGridItems | Where-Object { $_.Selected })

    if (($targetProfiles | Measure-Object).Count -eq 0) {
        [System.Windows.MessageBox]::Show('No profiles selected for export.', 'No Selection', 'OK', 'Warning') | Out-Null
        return
    }

    $targetCount = ($targetProfiles | Measure-Object).Count
    if ($targetCount -gt 0) {
        Write-ProfileReportLog "Selected profiles for export: $targetCount"
    } else {
        Write-ProfileReportLog 'No selected profiles detected at export start.'
    }

    $outputPath = $txtProfileReportOutputPath.Text.Trim()
    if (-not $outputPath) {
        $outputPath = Join-Path ([Environment]::GetFolderPath('Desktop')) 'Intune-Policy-Report.html'
        $txtProfileReportOutputPath.Text = $outputPath
    }

    try {
        $btnProfileReportGenerate.IsEnabled = $false
        $btnProfileReportWarmCache.IsEnabled = $false
        $btnProfileReportLoad.IsEnabled = $false
        $btnProfileReportOpenHtml.IsEnabled = $false
        $profileReportProgressBar.IsIndeterminate = $false
        $profileReportProgressBar.Value = 0
        $txtProfileReportProgress.Text = 'Collecting settings...'
        $txtProfileReportLog.Text = ''

        $cacheMode = Get-CacheModeFromComboBox -ComboBox $cmbProfileReportCacheMode
        $reportResult = Invoke-BaselineProfileReport -TargetProfiles $targetProfiles -OutputPath $outputPath -BaselineFolder $txtProfileReportBaselineFolder.Text.Trim() -CacheMode $cacheMode -LogCallback { param($message) Write-ProfileReportLog $message } -ProgressCallback { param($percent,$status) if ($percent -ge 0) { $profileReportProgressBar.Value = $percent }; $txtProfileReportProgress.Text = $status }
        $failedProfiles = @($reportResult.FailedProfiles)
        $script:intuneReportLastHtmlPath = $outputPath
        $btnProfileReportOpenHtml.IsEnabled = $true

        $profileReportProgressBar.Value = 100
        $txtProfileReportProgress.Text = 'Completed'
        Write-ProfileReportLog "Report created: $outputPath"
        if (@($failedProfiles).Count -gt 0) {
            Write-ProfileReportLog "Completed with $(@($failedProfiles).Count) failed profile(s)."
        }

        if ($chkProfileReportAutoOpen.IsChecked) {
            Start-Process -FilePath $outputPath
        }
    } catch {
        $profileReportProgressBar.Value = 0
        $txtProfileReportProgress.Text = 'Error'
        Write-ProfileReportLog "ERROR: $($_.Exception.Message)"
        [System.Windows.MessageBox]::Show("Failed to generate report: $($_.Exception.Message)", 'Error', 'OK', 'Error') | Out-Null
    } finally {
        $btnProfileReportGenerate.IsEnabled = $true
        $btnProfileReportWarmCache.IsEnabled = $true
        $btnProfileReportLoad.IsEnabled = $true
    }
})

$btnProfileReportOpenHtml.Add_Click({
    if ($script:intuneReportLastHtmlPath -and (Test-Path $script:intuneReportLastHtmlPath)) {
        Start-Process -FilePath $script:intuneReportLastHtmlPath
    } else {
        [System.Windows.MessageBox]::Show('Report file not found.', 'No Report', 'OK', 'Warning') | Out-Null
    }
})

# ══════════════════════════════════════════════════════════════════════════════
#  TAB 5 — BASELINE ASSESSMENT EVENT HANDLERS
# ══════════════════════════════════════════════════════════════════════════════

$btnAssessmentGraphConnect.Add_Click({
    try {
        $requiredTenant = Get-ResolverTenantOverride
        $ctx = Get-MgContext -ErrorAction SilentlyContinue
        if ($ctx -and $ctx.TenantId -and (-not $requiredTenant -or ([string]$ctx.TenantId -ieq [string]$requiredTenant))) {
            Disconnect-MgGraph -ErrorAction SilentlyContinue
            Update-PolicyTenantInfo
            Update-ProfileReportTenantInfo | Out-Null
            Update-AssessmentTenantInfo | Out-Null
        } else {
            Connect-BaselineReadGraph -TenantId $requiredTenant
            Update-PolicyTenantInfo
            Update-ProfileReportTenantInfo | Out-Null
            Update-AssessmentTenantInfo | Out-Null
        }
    } catch {
        [System.Windows.MessageBox]::Show("Graph error: $($_.Exception.Message)", 'Error', 'OK', 'Error') | Out-Null
    }
})

$btnAssessmentBrowseBaseline.Add_Click({
    Add-Type -AssemblyName System.Windows.Forms
    $fbd = [System.Windows.Forms.FolderBrowserDialog]::new()
    $fbd.Description = 'Select folder containing JSON Intune policy backup files'
    $fbd.ShowNewFolderButton = $false
    if ($txtAssessmentBaselineFolder.Text.Trim() -and (Test-Path $txtAssessmentBaselineFolder.Text.Trim())) {
        $fbd.SelectedPath = $txtAssessmentBaselineFolder.Text.Trim()
    }
    if ($fbd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $txtAssessmentBaselineFolder.Text = $fbd.SelectedPath
    }
})

$btnAssessmentLoad.Add_Click({
    if (-not (Ensure-ProfileReportGraphConnection)) { return }

    $script:assessmentUiState.IsBusy = $true
    $btnAssessmentLoad.IsEnabled = $false
    $btnAssessmentRun.IsEnabled = $false
    $assessmentProgressBar.Visibility = 'Visible'
    $assessmentProgressBar.IsIndeterminate = $true
    $txtAssessmentProgress.Text = 'Loading profiles...'
    $txtAssessmentProgress.Foreground = $window.FindResource('SubTextBrush')
    $txtAssessmentLog.Text = ''

    Write-AssessmentLog 'Loading all configuration policies from Intune...'

    # ── Shared state for progress reporting between runspace and UI timer ─────
    $loadSyncHash = [hashtable]::Synchronized(@{
        CommonModulePath = (Join-Path $PSScriptRoot 'Shared\BaselineToolkit.Common.psm1')
        StatusText = 'Fetching Settings Catalog profiles...'
        LogQueue   = [System.Collections.Concurrent.ConcurrentQueue[string]]::new()
        Done       = $false
        Error      = $null
        Result     = $null
    })

    $runspace = [runspacefactory]::CreateRunspace()
    $runspace.ApartmentState = 'STA'
    $runspace.Open()

    $ps = [powershell]::Create()
    $ps.Runspace = $runspace
    $ps.AddScript({ param([hashtable]$syncHash) try { Import-Module $syncHash.CommonModulePath -DisableNameChecking; $syncHash.Result = Get-BaselineIntuneProfiles } catch { $syncHash.Error = $_.Exception.Message }; $syncHash.Done = $true }) | Out-Null
    $ps.AddParameter('syncHash', $loadSyncHash) | Out-Null

    $handle = $ps.BeginInvoke()

    # Capture references for the timer closure — GetNewClosure creates a new module scope
    # where ambient function/control lookups are unreliable in pwsh 7.
    $localAssessmentCollection = $script:assessmentProfilesCollection
    $localDgAssessment = $dgAssessmentProfiles
    $localTxtFilter = $txtAssessmentFilter
    $localAssessmentProgressText = $txtAssessmentProgress
    $localAssessmentSuccessBrush = $window.FindResource('SuccessBrush')
    $localAssessmentDangerBrush = $window.FindResource('DangerBrush')
    $localAssessmentProgressBar = $assessmentProgressBar
    $localAssessmentLoadButton = $btnAssessmentLoad
    $localAssessmentRunButton = $btnAssessmentRun
    $localAssessmentUiState = $script:assessmentUiState
    $localUpdateAssessmentActionState = ${function:Update-AssessmentActionState}
    $localWriteAssessmentLog = ${function:Write-AssessmentLog}
    $localGetSafeObjectPropertyValue = ${function:Get-SafeObjectPropertyValue}
    $localApplyAssessmentFilter = ${function:Apply-AssessmentFilter}

    $timer = New-Object System.Windows.Threading.DispatcherTimer
    $timer.Interval = [timespan]::FromMilliseconds(250)
    $localLoadTimer = $timer
    $timer.Add_Tick({
        # Drain log messages from runspace
        $qmsg = [string]''
        while ($loadSyncHash.LogQueue.TryDequeue([ref]$qmsg)) { & $localWriteAssessmentLog $qmsg }
        $localAssessmentProgressText.Text = $loadSyncHash.StatusText

        if ($loadSyncHash.Done) {
            $localLoadTimer.Stop()

            # Final drain
            while ($loadSyncHash.LogQueue.TryDequeue([ref]$qmsg)) { & $localWriteAssessmentLog $qmsg }

            try {
                $null = $ps.EndInvoke($handle)

                if ($loadSyncHash.Error) {
                    throw $loadSyncHash.Error
                }

                $mapped = @()
                $resultObj = $loadSyncHash.Result
                if ($resultObj) { try { $mapped = @($resultObj.Mapped) } catch {} }

                $normalized = @()
                foreach ($m in $mapped) {
                    $idv = & $localGetSafeObjectPropertyValue -Object $m -Name 'Id'
                    $namev = & $localGetSafeObjectPropertyValue -Object $m -Name 'Name'
                    if (-not $idv -or -not $namev) { continue }
                    $platv = & $localGetSafeObjectPropertyValue -Object $m -Name 'Platforms'
                    $lastv = & $localGetSafeObjectPropertyValue -Object $m -Name 'LastModifiedDateTime'
                    $kindv = & $localGetSafeObjectPropertyValue -Object $m -Name 'PolicyKind'
                    $typev = & $localGetSafeObjectPropertyValue -Object $m -Name 'OdataType'
                    $normalized += [pscustomobject]@{
                        Selected = $false
                        Id = [string]$idv
                        Name = [string]$namev
                        Platforms = if ($platv) { [string]$platv } else { '-' }
                        LastModifiedDateTime = if ($lastv) { [string]$lastv } else { '-' }
                        PolicyKind    = if ($kindv) { [string]$kindv } else { 'SettingsCatalog' }
                        OdataType     = if ($typev) { [string]$typev } else { '' }
                        HasAssignment = $true
                        AssignedTo    = ''
                    }
                }

                $script:assessmentProfilesAll = $normalized

                if (-not $localAssessmentCollection) {
                    $localAssessmentCollection = [System.Collections.ObjectModel.ObservableCollection[object]]::new()
                    $localDgAssessment.ItemsSource = $localAssessmentCollection
                    $script:assessmentProfilesCollection = $localAssessmentCollection
                }

                $localAssessmentCollection.Clear()
                foreach ($p in $script:assessmentProfilesAll) {
                    $localAssessmentCollection.Add($p)
                }

                if ($localTxtFilter) { $localTxtFilter.Text = '' }
                & $localApplyAssessmentFilter -FilterText ''
                & $localWriteAssessmentLog "Loaded $($script:assessmentProfilesAll.Count) profiles."
                $localAssessmentProgressBar.IsIndeterminate = $false
                $localAssessmentProgressBar.Visibility = 'Collapsed'
                $localAssessmentProgressBar.Value = 100
                $localAssessmentProgressText.Text = "Loaded $($normalized.Count) profiles"
                $localAssessmentProgressText.Foreground = $localAssessmentSuccessBrush
            } catch {
                $localAssessmentProgressBar.IsIndeterminate = $false
                $localAssessmentProgressBar.Visibility = 'Collapsed'
                $localAssessmentProgressBar.Value = 0
                $localAssessmentProgressText.Text = 'Profile load failed'
                $localAssessmentProgressText.Foreground = $localAssessmentDangerBrush
                & $localWriteAssessmentLog "ERROR: $($_.Exception.Message)"
                if ($_.ScriptStackTrace) { & $localWriteAssessmentLog "Stack: $($_.ScriptStackTrace)" }
            } finally {
                $ps.Dispose()
                $runspace.Close()
                $runspace.Dispose()
                $localAssessmentLoadButton.IsEnabled = $true
                $localAssessmentRunButton.IsEnabled = $true
                $localAssessmentUiState.IsBusy = $false
                & $localUpdateAssessmentActionState
            }
        }
    }.GetNewClosure())
    $timer.Start()
})

$txtAssessmentFilter.Add_TextChanged({
    $txtAssessmentFilterPlaceholder.Visibility = if ($txtAssessmentFilter.Text.Length -gt 0) { 'Collapsed' } else { 'Visible' }
    Apply-AssessmentFilter -FilterText $txtAssessmentFilter.Text
})

$txtAssessmentBaselineFolder.Add_TextChanged({ Update-AssessmentActionState })

$btnAssessmentSelectAll.Add_Click({
    $gridSource = $dgAssessmentProfiles.ItemsSource
    if ($null -eq $gridSource) { return }
    $view = [System.Windows.Data.CollectionViewSource]::GetDefaultView($gridSource)
    $source = if ($null -ne $view) { @($view | ForEach-Object { $_ }) } else { @($gridSource | ForEach-Object { $_ }) }
    foreach ($p in $source) { $p.Selected = $true }
    $dgAssessmentProfiles.Items.Refresh()
    Update-AssessmentSelectionCount
})

$btnAssessmentSelectNone.Add_Click({
    $gridSource = $dgAssessmentProfiles.ItemsSource
    if ($null -eq $gridSource) { return }
    $view = [System.Windows.Data.CollectionViewSource]::GetDefaultView($gridSource)
    $source = if ($null -ne $view) { @($view | ForEach-Object { $_ }) } else { @($gridSource | ForEach-Object { $_ }) }
    foreach ($p in $source) { $p.Selected = $false }
    $dgAssessmentProfiles.Items.Refresh()
    Update-AssessmentSelectionCount
})

$dgAssessmentProfiles.Add_CellEditEnding({
    $dgAssessmentProfiles.Dispatcher.BeginInvoke([Action]{ Update-AssessmentSelectionCount }) | Out-Null
})

$btnAssessmentBrowseOutput.Add_Click({
    $dlg = [Microsoft.Win32.SaveFileDialog]::new()
    $dlg.Title = 'Save Assessment Report As'
    $dlg.Filter = 'HTML files (*.html)|*.html|All files (*.*)|*.*'
    $dlg.FileName = 'Baseline-Import-Assessment.html'
    if ($dlg.ShowDialog($window)) {
        $txtAssessmentOutputPath.Text = $dlg.FileName
    }
})

$btnAssessmentOpenHtml.Add_Click({
    $htmlPath = $script:assessmentResultStore.HtmlPath
    if ($htmlPath -and (Test-Path $htmlPath)) {
        Start-Process -FilePath $htmlPath
    } else {
        [System.Windows.MessageBox]::Show('Report file not found.', 'No Report', 'OK', 'Warning') | Out-Null
    }
})

$btnAssessmentCreateSplitPolicy.Add_Click({
    $requiredTenant = Get-ResolverTenantOverride
    try {
        $ctx = Get-MgContext -ErrorAction SilentlyContinue
        $hasWriteScope = $false
        if ($ctx -and $ctx.Scopes) {
            $hasWriteScope = @($ctx.Scopes) -contains 'DeviceManagementConfiguration.ReadWrite.All'
        }
        if (-not $ctx -or -not $hasWriteScope) {
            $connectParams = @{
                Scopes      = @('DeviceManagementConfiguration.ReadWrite.All','Organization.Read.All')
                NoWelcome   = $true
                ErrorAction = 'Stop'
            }
            if ($requiredTenant) { $connectParams['TenantId'] = $requiredTenant }
            Connect-MgGraph @connectParams | Out-Null
            Update-AssessmentTenantInfo
        }
    } catch {
        [System.Windows.MessageBox]::Show("Failed to connect to Microsoft Graph with write scope:`n$($_.Exception.Message)", 'Graph Error', 'OK', 'Error') | Out-Null
        return
    }

    if (-not (Ensure-ProfileReportGraphConnection)) { return }

    $assessmentData = $script:assessmentResultStore.Data
    if (-not $assessmentData -or -not $assessmentData.Baselines) {
        [System.Windows.MessageBox]::Show('Run the assessment first so split-required baselines are available.', 'No Assessment', 'OK', 'Warning') | Out-Null
        return
    }

    $splitBaselines = @($assessmentData.Baselines | Where-Object { $_.Classification -eq 'SplitRequired' })
    if ($splitBaselines.Count -eq 0) {
        [System.Windows.MessageBox]::Show('No split-required baselines were found in the last assessment.', 'No Split Required Items', 'OK', 'Warning') | Out-Null
        return
    }

    $selectedBaselines = @(Show-SplitRequiredBaselinePicker -Baselines $splitBaselines)
    if ($selectedBaselines.Count -eq 0) { return }

    $confirm = [System.Windows.MessageBox]::Show(
        "Create $($selectedBaselines.Count) new split policy/policies from the selected baseline(s)?",
        'Create Split Policy',
        'YesNo',
        'Question'
    )
    if ($confirm -ne [System.Windows.MessageBoxResult]::Yes) { return }

    try {
        $script:assessmentUiState.IsBusy = $true
        $btnAssessmentCreateSplitPolicy.IsEnabled = $false
        $btnAssessmentRun.IsEnabled = $false
        $btnAssessmentLoad.IsEnabled = $false
        $assessmentProgressBar.Visibility = 'Visible'
        $assessmentProgressBar.IsIndeterminate = $true
        $txtAssessmentProgress.Foreground = $window.FindResource('SubTextBrush')

        $createdPolicies = [System.Collections.Generic.List[string]]::new()
        $failedPolicies  = [System.Collections.Generic.List[string]]::new()

        foreach ($selectedBaseline in $selectedBaselines) {
            $baselinePolicyName = [string](Get-SafeObjectPropertyValue -Object $selectedBaseline -Name 'PolicyName')
            $newPolicyName = Get-BaselinePolicyNameWithoutSecurityBaseline -PolicyName $baselinePolicyName
            $txtAssessmentProgress.Text = "Creating split policy '$newPolicyName'..."
            Write-AssessmentLog "Creating split policy '$newPolicyName'..."
            try {
                $result = New-SplitRequiredPolicyFromAssessment -AssessmentData $assessmentData -Baseline $selectedBaseline
                $createdId = if ($result) { [string]$result.Id } else { '' }
                $sourceName = if ($result -and $result.SourcePolicyName) { [string]$result.SourcePolicyName } else { '' }
                $createdCount = if ($result) { [int]$result.ExtraSettingCount } else { 0 }
                $createdPolicies.Add("$($result.Name) ($createdCount setting(s) from '$sourceName')")
                Write-AssessmentLog "Created policy '$($result.Name)' from '$sourceName' with $createdCount extra setting(s)$(if ($createdId) { " (ID: $createdId)" } else { '' })."
            } catch {
                $failedPolicies.Add("$newPolicyName - $($_.Exception.Message)")
                Write-AssessmentLog "ERROR creating split policy '$newPolicyName': $($_.Exception.Message)"
            }
        }

        $summary = "Created $($createdPolicies.Count) policy/policies."
        if ($createdPolicies.Count -gt 0) { $summary += "`n`n" + ($createdPolicies -join "`n") }
        if ($failedPolicies.Count -gt 0) { $summary += "`n`nFailed ($($failedPolicies.Count)):`n" + ($failedPolicies -join "`n") }
        $txtAssessmentProgress.Text = "Split policy creation finished: $($createdPolicies.Count) created, $($failedPolicies.Count) failed"
        $txtAssessmentProgress.Foreground = if ($failedPolicies.Count -gt 0) { $window.FindResource('DangerBrush') } else { $window.FindResource('SuccessBrush') }
        [System.Windows.MessageBox]::Show($summary, 'Split Policy Creation', 'OK', 'Information') | Out-Null
    } catch {
        $txtAssessmentProgress.Foreground = $window.FindResource('DangerBrush')
        Write-AssessmentLog "ERROR: $($_.Exception.Message)"
        [System.Windows.MessageBox]::Show("Failed to create split policy:`n$($_.Exception.Message)", 'Error', 'OK', 'Error') | Out-Null
    } finally {
        $assessmentProgressBar.IsIndeterminate = $false
        $assessmentProgressBar.Visibility = 'Collapsed'
        $btnAssessmentRun.IsEnabled = $true
        $btnAssessmentLoad.IsEnabled = $true
        $btnAssessmentCreateSplitPolicy.IsEnabled = $true
        $script:assessmentUiState.IsBusy = $false
        Update-AssessmentActionState
    }
})

$btnAssessmentRun.Add_Click({
    # ── UI-thread validation (fast, no Graph calls) ──────────────────────────
    $baselineFolder = $txtAssessmentBaselineFolder.Text.Trim()
    if (-not $baselineFolder -or -not (Test-Path $baselineFolder -PathType Container)) {
        [System.Windows.MessageBox]::Show('Please select a baseline folder containing JSON files.', 'Missing', 'OK', 'Warning') | Out-Null
        return
    }

    if (-not (Ensure-ProfileReportGraphConnection)) { return }

    try {
        $null = $dgAssessmentProfiles.CommitEdit([System.Windows.Controls.DataGridEditingUnit]::Cell, $true)
        $null = $dgAssessmentProfiles.CommitEdit([System.Windows.Controls.DataGridEditingUnit]::Row, $true)
    } catch {}

    $gridSource   = $dgAssessmentProfiles.ItemsSource
    $allGridItems = if ($null -ne $gridSource) { @($gridSource | ForEach-Object { $_ }) } else { @() }
    $targetProfiles = @($allGridItems | Where-Object { $_.Selected })

    if ($targetProfiles.Count -eq 0) {
        [System.Windows.MessageBox]::Show('No profiles selected. Select customer profiles to assess against.', 'No Selection', 'OK', 'Warning') | Out-Null
        return
    }

    # Build / default output path
    $outputPath = $txtAssessmentOutputPath.Text.Trim()
    if (-not $outputPath) {
        $tenantPrefix = ''
        try {
            $org2 = Invoke-MgGraphRequest -Uri 'https://graph.microsoft.com/v1.0/organization' -Method Get -ErrorAction SilentlyContinue
            $orgName2 = Get-SafeObjectPropertyValue -Object ($org2.value | Select-Object -First 1) -Name 'displayName'
            if ($orgName2) { $tenantPrefix = ($orgName2 -replace '[^A-Za-z0-9\-_ ]', '' -replace '\s+', '-') + '_' }
        } catch {}
        $outputPath = Join-Path ([Environment]::GetFolderPath('Desktop')) "${tenantPrefix}Baseline-Import-Assessment_$(Get-Date -Format 'yyyy-MM-dd').html"
        $txtAssessmentOutputPath.Text = $outputPath
    }

    # Resolve tenant display name once on UI thread (fast)
    $tenantDisplayName = ''
    try {
        $orgT = Invoke-MgGraphRequest -Uri 'https://graph.microsoft.com/v1.0/organization' -Method Get -ErrorAction SilentlyContinue
        $tn = Get-SafeObjectPropertyValue -Object ($orgT.value | Select-Object -First 1) -Name 'displayName'
        if ($tn) { $tenantDisplayName = [string]$tn }
    } catch {}

    # ── Lock UI ───────────────────────────────────────────────────────────────
    $script:assessmentUiState.IsBusy = $true
    $btnAssessmentRun.IsEnabled     = $false
    $btnAssessmentLoad.IsEnabled    = $false
    $btnAssessmentOpenHtml.IsEnabled = $false
    $btnAssessmentCreateSplitPolicy.IsEnabled = $false
    $assessmentProgressBar.Visibility = 'Visible'
    $assessmentProgressBar.IsIndeterminate = $false
    $assessmentProgressBar.Value    = 0
    $txtAssessmentProgress.Text     = 'Starting...'
    $txtAssessmentProgress.Foreground = $window.FindResource('SubTextBrush')
    $txtAssessmentLog.Text          = ''

    $requiredTenant = Get-ResolverTenantOverride
    if ($requiredTenant) {
        Write-AssessmentLog "Resolver tenant override active: $requiredTenant"
    }
    $resolverTenantCacheKey = Get-GraphTenantCacheKey
    $cacheMode = Get-CacheModeFromComboBox -ComboBox $cmbAssessmentCacheMode
    Write-AssessmentLog "Resolver mode: $(Get-CacheModeLabel -Mode $cacheMode)"

    # ── Shared state between UI thread and runspace ───────────────────────────
    $syncHash = [hashtable]::Synchronized(@{
        # Inputs
        TargetProfiles     = @($targetProfiles | ForEach-Object {
            [pscustomobject]@{
                Id         = [string]$_.Id
                Name       = [string]$_.Name
                Platforms  = [string]$_.Platforms
                PolicyKind = [string]$_.PolicyKind
                OdataType  = [string]$_.OdataType
            }
        })
        BaselineFolder     = $baselineFolder
        OutputPath         = $outputPath
        FeatureScriptPath = (Join-Path $PSScriptRoot 'Features\Invoke-BaselineAssessment.ps1')
        CustomerPolicyRawSettings = @{}
        CustomerPolicyMeta = @{}
        SkipUnassigned        = [bool]$chkAssessmentSkipUnassigned.IsChecked
        SkipAssignmentLookups = [bool]$chkAssessmentSkipAssignments.IsChecked
        ProfileAssignments    = @{}
        TenantName            = $tenantDisplayName
        ResolverTenantOverride = $requiredTenant
        ResolverTenantCacheKey = $resolverTenantCacheKey
        CacheMode             = $cacheMode
        AutoOpen              = [bool]$chkAssessmentAutoOpen.IsChecked

        # Function bodies serialised as strings so they can be rebuilt in the new runspace
        

        # Progress (written by runspace, read by timer on UI thread)
        Progress        = 0
        IsIndeterminate = $false
        StatusText      = 'Starting...'
        LogQueue        = [System.Collections.Concurrent.ConcurrentQueue[string]]::new()

        # Result
        Done  = $false
        Error = $null
    })

    # ── Background runspace ───────────────────────────────────────────────────
    $runspace = [runspacefactory]::CreateRunspace()
    $runspace.ApartmentState = 'STA'
    $runspace.Open()
    $ps = [powershell]::Create()
    $ps.Runspace = $runspace
    $ps.AddScript({ param([hashtable]$syncHash) . $syncHash.FeatureScriptPath; Invoke-BaselineAssessmentWorker -syncHash $syncHash }) | Out-Null
    $ps.AddParameter('syncHash', $syncHash) | Out-Null

    $handle = $ps.BeginInvoke()

    # ── DispatcherTimer: drains log queue and updates progress bar ────────────
    $localAssessmentProgressBar = $assessmentProgressBar
    $localAssessmentProgressText = $txtAssessmentProgress
    $localAssessmentSuccessBrush = $window.FindResource('SuccessBrush')
    $localAssessmentAttentionBrush = $window.FindResource('AttentionBrush')
    $localAssessmentDangerBrush = $window.FindResource('DangerBrush')
    $localAssessmentLoadButton = $btnAssessmentLoad
    $localAssessmentRunButton = $btnAssessmentRun
    $localAssessmentOpenHtmlButton = $btnAssessmentOpenHtml
    $localAssessmentSplitPolicyButton = $btnAssessmentCreateSplitPolicy
    $localAssessmentResultStore = $assessmentResultStore
    $localAssessmentUiState = $script:assessmentUiState
    $localUpdateAssessmentActionState = ${function:Update-AssessmentActionState}
    $localWriteAssessmentLog = ${function:Write-AssessmentLog}

    $assessmentRunTimer = New-Object System.Windows.Threading.DispatcherTimer
    $assessmentRunTimer.Interval = [timespan]::FromMilliseconds(250)
    $localAssessmentRunTimer = $assessmentRunTimer
    $assessmentRunTimer.Add_Tick({
        # Drain enqueued log lines onto the UI-thread log control
        $qmsg = [string]''
        while ($syncHash.LogQueue.TryDequeue([ref]$qmsg)) { & $localWriteAssessmentLog $qmsg }

        $localAssessmentProgressBar.IsIndeterminate = $syncHash.IsIndeterminate
        if (-not $syncHash.IsIndeterminate) { $localAssessmentProgressBar.Value = $syncHash.Progress }
        $localAssessmentProgressText.Text = $syncHash.StatusText

        if ($syncHash.Done) {
            $localAssessmentRunTimer.Stop()

            # Final drain
            while ($syncHash.LogQueue.TryDequeue([ref]$qmsg)) { & $localWriteAssessmentLog $qmsg }

            try { $null = $ps.EndInvoke($handle) } catch {}
            $ps.Dispose()
            $runspace.Close()
            $runspace.Dispose()

            if ($syncHash.Error) {
                $localAssessmentProgressBar.IsIndeterminate = $false
                $localAssessmentProgressBar.Visibility = 'Collapsed'
                $localAssessmentProgressBar.Value = 0
                $localAssessmentProgressText.Text  = if ($syncHash.Error -eq 'NO_PROFILES') { if ($syncHash.StatusText) { [string]$syncHash.StatusText } else { 'No assigned profiles' } } else { 'Assessment failed' }
                $localAssessmentProgressText.Foreground = if ($syncHash.Error -eq 'NO_PROFILES') { $localAssessmentAttentionBrush } else { $localAssessmentDangerBrush }
                if ($syncHash.Error -ne 'NO_PROFILES') {
                    [System.Windows.MessageBox]::Show("Assessment failed:`n$($syncHash.Error)", 'Error', 'OK', 'Error') | Out-Null
                }
            } else {
                $localAssessmentResultStore.HtmlPath  = $syncHash.OutputPath
                $localAssessmentResultStore.Data      = $syncHash.AssessmentData
                $script:assessmentLastHtmlPath   = $syncHash.OutputPath
                $script:assessmentLastData       = $syncHash.AssessmentData
                $localAssessmentOpenHtmlButton.IsEnabled  = $true
                $localAssessmentSplitPolicyButton.IsEnabled = @($syncHash.AssessmentData.Baselines | Where-Object { $_.Classification -eq 'SplitRequired' }).Count -gt 0
                $localAssessmentProgressBar.IsIndeterminate = $false
                $localAssessmentProgressBar.Visibility = 'Collapsed'
                $localAssessmentProgressBar.Value     = 100
                $localAssessmentProgressText.Text      = '✓ Assessment completed successfully'
                $localAssessmentProgressText.Foreground = $localAssessmentSuccessBrush
                if ($syncHash.AutoOpen) { Start-Process -FilePath $syncHash.OutputPath }
            }

            $localAssessmentRunButton.IsEnabled  = $true
            $localAssessmentLoadButton.IsEnabled = $true
            $localAssessmentUiState.IsBusy = $false
            & $localUpdateAssessmentActionState
        }
    }.GetNewClosure())
    $assessmentRunTimer.Start()
})

# ══════════════════════════════════════════════════════════════════════════════
#  STARTUP
# ══════════════════════════════════════════════════════════════════════════════

# Check Python prerequisites for Tab 1
$prereq = Test-Prerequisites
Update-PdfDependencyUi -Prerequisites $prereq

# Check Graph connection for Tab 3
try { Update-PolicyTenantInfo } catch { }
try { Update-ProfileReportTenantInfo | Out-Null } catch { }

# Tab 4 defaults
$txtProfileReportOutputPath.Text = Join-Path ([Environment]::GetFolderPath('Desktop')) 'Intune-Policy-Report.html'
$txtProfileReportProgress.Text = ''
Update-ProfileReportSelectionCount

$btnAssessmentWorkbookBrowseCsv.Add_Click({
    $dialog = [Microsoft.Win32.OpenFileDialog]::new()
    $dialog.Title = 'Select Baseline Assessment Details CSV'
    $dialog.Filter = 'CSV files (*.csv)|*.csv|All files (*.*)|*.*'
    if ($dialog.ShowDialog($window)) {
        $txtAssessmentWorkbookCsv.Text = $dialog.FileName
        $btnAssessmentWorkbookOpen.IsEnabled = $false
        $script:lastAssessmentWorkbookPath = $null
        if (-not $txtAssessmentWorkbookOutput.Text) {
            $csvBaseName = [System.IO.Path]::GetFileNameWithoutExtension($dialog.FileName)
            $txtAssessmentWorkbookOutput.Text = Join-Path (Split-Path -Path $dialog.FileName -Parent) "$csvBaseName-Overview.xlsx"
        }
        Write-AssessmentWorkbookLog "Details CSV selected: $($dialog.FileName)"
    }
})

$btnAssessmentWorkbookBrowseMarkdown.Add_Click({
    $dialog = [System.Windows.Forms.FolderBrowserDialog]::new()
    $dialog.Description = 'Select the folder containing source CIS Markdown files'
    $dialog.ShowNewFolderButton = $false
    if ($txtAssessmentWorkbookMarkdownFolder.Text -and (Test-Path -Path $txtAssessmentWorkbookMarkdownFolder.Text -PathType Container)) {
        $dialog.SelectedPath = $txtAssessmentWorkbookMarkdownFolder.Text
    }
    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $txtAssessmentWorkbookMarkdownFolder.Text = $dialog.SelectedPath
        Write-AssessmentWorkbookLog "Markdown folder selected: $($dialog.SelectedPath)"
    }
})

$btnAssessmentWorkbookBrowseOutput.Add_Click({
    $dialog = [Microsoft.Win32.SaveFileDialog]::new()
    $dialog.Title = 'Save Baseline Settings Workbook As'
    $dialog.Filter = 'Excel workbook (*.xlsx)|*.xlsx'
    $dialog.DefaultExt = '.xlsx'
    $dialog.AddExtension = $true
    $dialog.OverwritePrompt = $false
    $dialog.FileName = if ($txtAssessmentWorkbookOutput.Text) {
        [System.IO.Path]::GetFileName($txtAssessmentWorkbookOutput.Text)
    } elseif ($txtAssessmentWorkbookCsv.Text) {
        ([System.IO.Path]::GetFileNameWithoutExtension($txtAssessmentWorkbookCsv.Text) + '-Overview.xlsx')
    } else {
        'Customer-SecurityBaselineSettings-Overview.xlsx'
    }
    if ($dialog.ShowDialog($window)) {
        $txtAssessmentWorkbookOutput.Text = $dialog.FileName
        $btnAssessmentWorkbookOpen.IsEnabled = $false
        $script:lastAssessmentWorkbookPath = $null
    }
})

$btnAssessmentWorkbookExport.Add_Click({
    $csvPath = $txtAssessmentWorkbookCsv.Text.Trim()
    $markdownFolder = $txtAssessmentWorkbookMarkdownFolder.Text.Trim()
    $outputPath = $txtAssessmentWorkbookOutput.Text.Trim()
    if (-not $csvPath) {
        [System.Windows.MessageBox]::Show('Select a Baseline Assessment Details CSV.', 'Missing CSV', 'OK', 'Warning') | Out-Null
        return
    }
    if (-not $markdownFolder) {
        [System.Windows.MessageBox]::Show('Select the folder containing source Markdown files.', 'Missing Markdown Folder', 'OK', 'Warning') | Out-Null
        return
    }
    if (-not $outputPath) {
        [System.Windows.MessageBox]::Show('Choose an output workbook path.', 'Missing Output Path', 'OK', 'Warning') | Out-Null
        return
    }
    $similarityThresholdRaw = $txtAssessmentWorkbookSimilarityThreshold.Text.Trim()
    $similarityThreshold = 0.9
    if (-not [double]::TryParse($similarityThresholdRaw, [System.Globalization.NumberStyles]::Float, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$similarityThreshold)) {
        [System.Windows.MessageBox]::Show('Text match threshold must be a number (for example 0.90 or 90).', 'Invalid Threshold', 'OK', 'Warning') | Out-Null
        return
    }
    if ($similarityThreshold -gt 1.0 -and $similarityThreshold -le 100.0) { $similarityThreshold = $similarityThreshold / 100.0 }
    if ($similarityThreshold -lt 0.0 -or $similarityThreshold -gt 1.0) {
        [System.Windows.MessageBox]::Show('Text match threshold must be between 0 and 1, or 0 and 100.', 'Invalid Threshold', 'OK', 'Warning') | Out-Null
        return
    }
    if ([System.IO.Path]::GetExtension($outputPath) -ine '.xlsx') {
        [System.Windows.MessageBox]::Show('The output workbook path must end in .xlsx.', 'Invalid Output Path', 'OK', 'Warning') | Out-Null
        return
    }
    if ([System.IO.Path]::GetFullPath($outputPath) -ieq [System.IO.Path]::GetFullPath((Join-Path (Split-Path -Path $script:thisScriptPath -Parent) 'Resources\Customer-SecurityBaselineSettings-Overview.xlsx'))) {
        [System.Windows.MessageBox]::Show('Choose a different output path. The bundled workbook template will not be overwritten.', 'Invalid Output Path', 'OK', 'Warning') | Out-Null
        return
    }
    if (Test-Path -Path $outputPath -PathType Leaf) {
        $overwriteChoice = [System.Windows.MessageBox]::Show("Replace the existing workbook?`n$outputPath", 'Confirm Overwrite', 'YesNo', 'Question')
        if ($overwriteChoice -ne [System.Windows.MessageBoxResult]::Yes) { return }
    }

    $btnAssessmentWorkbookExport.IsEnabled = $false
    $btnAssessmentWorkbookBrowseCsv.IsEnabled = $false
    $btnAssessmentWorkbookBrowseMarkdown.IsEnabled = $false
    $btnAssessmentWorkbookBrowseOutput.IsEnabled = $false
    $btnAssessmentWorkbookOpen.IsEnabled = $false
    $assessmentWorkbookProgressBar.IsIndeterminate = $true
    $assessmentWorkbookProgressBar.Value = 0
    $txtAssessmentWorkbookStatus.Text = 'Importing CSV and scanning Markdown...'
    $txtAssessmentWorkbookLog.Text = ''
    $assessmentWorkbookActivityExpander.IsExpanded = $true
    Write-AssessmentWorkbookLog 'Starting workbook export.'
    Write-AssessmentWorkbookLog "CSV: $csvPath"
    Write-AssessmentWorkbookLog "Markdown folder: $markdownFolder"
    Write-AssessmentWorkbookLog "Text match threshold: $similarityThreshold"
    Write-AssessmentWorkbookLog "Output: $outputPath"

    try {
        $result = Export-BaselineAssessmentWorkbook -CsvPath $csvPath -MarkdownFolder $markdownFolder -OutputPath $outputPath -SimilarityThreshold $similarityThreshold
        $script:lastAssessmentWorkbookPath = $result.OutputPath
        $btnAssessmentWorkbookOpen.IsEnabled = $true
        $assessmentWorkbookProgressBar.IsIndeterminate = $false
        $assessmentWorkbookProgressBar.Value = 100
        $txtAssessmentWorkbookStatus.Text = "$($result.RowCount) settings exported | $($result.MatchedCount) references matched | $($result.TwodayCount) Twoday | $($result.NoMatchCount) not found | $($result.AmbiguousCount) ambiguous"
        Write-AssessmentWorkbookLog "Workbook created: $($result.OutputPath)"
        Write-AssessmentWorkbookLog "Rows: $($result.RowCount) | Matched: $($result.MatchedCount) | Twoday: $($result.TwodayCount) | No match: $($result.NoMatchCount) | Ambiguous: $($result.AmbiguousCount)"
        Write-AssessmentWorkbookLog "Duplicates removed: $($result.DuplicatesRemoved) | Conflicts: $($result.ConflictCount)"
    } catch {
        $assessmentWorkbookProgressBar.IsIndeterminate = $false
        $assessmentWorkbookProgressBar.Value = 0
        $txtAssessmentWorkbookStatus.Text = 'Export failed. See activity for details.'
        Write-AssessmentWorkbookLog "ERROR: $($_.Exception.Message)"
        [System.Windows.MessageBox]::Show("Workbook export failed:`n$($_.Exception.Message)", 'Export Error', 'OK', 'Error') | Out-Null
    } finally {
        $btnAssessmentWorkbookExport.IsEnabled = $true
        $btnAssessmentWorkbookBrowseCsv.IsEnabled = $true
        $btnAssessmentWorkbookBrowseMarkdown.IsEnabled = $true
        $btnAssessmentWorkbookBrowseOutput.IsEnabled = $true
    }
})

$btnAssessmentWorkbookOpen.Add_Click({
    if ($script:lastAssessmentWorkbookPath -and (Test-Path -Path $script:lastAssessmentWorkbookPath -PathType Leaf)) {
        Start-Process -FilePath $script:lastAssessmentWorkbookPath | Out-Null
    } else {
        [System.Windows.MessageBox]::Show('Workbook not found. Run the export first.', 'No Workbook', 'OK', 'Warning') | Out-Null
    }
})

# Apply initial dark theme
$res = $window.Resources
$res['BgBrush']      = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.ColorConverter]::ConvertFromString('#0e151d'))
$res['TextBrush']    = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.ColorConverter]::ConvertFromString('#e6eef8'))
$res['SubTextBrush'] = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.ColorConverter]::ConvertFromString('#9cb2c8'))
$res['SurfaceBrush'] = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.ColorConverter]::ConvertFromString('#1a2430'))
$res['InputBgBrush'] = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.ColorConverter]::ConvertFromString('#111820'))
$res['BorderBrush2'] = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.ColorConverter]::ConvertFromString('#2f4358'))

# Initial log messages
Write-MergeLog 'CIS Baseline Merger ready.'
Write-ProfileReportLog 'Intune Policy Report ready.'

# Tab 5 defaults
$txtAssessmentOutputPath.Text = ''
$txtAssessmentProgress.Text = ''
Update-AssessmentSelectionCount
try { Update-AssessmentTenantInfo | Out-Null } catch { }
Write-AssessmentLog 'Baseline Assessment ready.'

$txtAssessmentWorkbookCsv.Text = ''
$txtAssessmentWorkbookMarkdownFolder.Text = ''
$txtAssessmentWorkbookOutput.Text = ''
$assessmentWorkbookProgressBar.IsIndeterminate = $false
$assessmentWorkbookProgressBar.Value = 0
$txtAssessmentWorkbookStatus.Text = ''
Write-AssessmentWorkbookLog 'Baseline workbook export ready.'

[void]$window.ShowDialog()
