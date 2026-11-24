@echo off
setlocal EnableExtensions DisableDelayedExpansion
title Assistente de desempenho
set "ASSISTENTE_ARQUIVO=%~f0"
"%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe" -NoLogo -NoProfile -STA -Command "try{$ErrorActionPreference='Stop';$s=[IO.File]::ReadAllText($env:ASSISTENTE_ARQUIVO,[Text.Encoding]::UTF8);$m=[char]10+'#INTERFACE'+[char]13+[char]10;$i=$s.IndexOf($m);if($i -lt 0){throw 'Arquivo incompleto'};& ([scriptblock]::Create($s.Substring($i+$m.Length)))}catch{Write-Host 'Nao foi possivel abrir o assistente.' $_.Exception.Message;exit 1}"
if errorlevel 1 pause
exit /b
#INTERFACE
Add-Type -AssemblyName PresentationFramework
[xml]$layout=@'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" Title="Assistente de desempenho" Width="600" Height="620" MinWidth="430" MinHeight="380" WindowStartupLocation="CenterScreen" Background="#F2F5F6" FontFamily="Segoe UI" FontSize="14" Foreground="#192F3D">
<Window.Resources>
<Style TargetType="Button"><Setter Property="Padding" Value="14,7"/><Setter Property="HorizontalAlignment" Value="Left"/><Setter Property="Background" Value="#E5F0EF"/><Setter Property="BorderBrush" Value="#819B9A"/><Setter Property="Cursor" Value="Hand"/></Style>
<Style TargetType="TextBlock"><Setter Property="TextWrapping" Value="Wrap"/></Style>
</Window.Resources>
<DockPanel>
<Border DockPanel.Dock="Top" Background="#153A45" Padding="24,16"><StackPanel>
<TextBlock Foreground="#B9DEDA" FontSize="13" Text="ESCRITÓRIO   /   WINDOWS 10 E 11"/>
<TextBlock Foreground="White" FontSize="25" FontWeight="SemiBold" Text="Assistente de desempenho"/>
<TextBlock Foreground="#D4E8E6" FontSize="14" Margin="0,4,0,0" Text="Enzo Roque"/>
</StackPanel></Border>
<ScrollViewer VerticalScrollBarVisibility="Auto"><StackPanel Margin="20,16">
<Border Background="White" CornerRadius="8" Padding="18"><StackPanel>
<TextBlock FontSize="19" FontWeight="SemiBold" Text="Liberar espaço"/>
<TextBlock Margin="0,5,0,12" Text="Procure arquivos temporários antigos. Você verá a quantidade e o tamanho antes de decidir se deseja excluí-los."/>
<Button Name="Scan" Content="Analisar temporários" Background="#17665F" Foreground="White"/>
<Border Background="#F1F6F5" Padding="10" Margin="0,12,0,0"><TextBlock Name="Status" Text="Pronto para analisar."/></Border>
<Button Name="Clean" Margin="0,10,0,0" Visibility="Collapsed" Content="Revisar e excluir…"/>
</StackPanel></Border>
<Border Background="White" CornerRadius="8" Padding="18" Margin="0,12,0,0"><StackPanel>
<TextBlock FontSize="19" FontWeight="SemiBold" Text="Agilizar a inicialização"/>
<TextBlock Margin="0,5,0,10" Text="Escolha quais aplicativos abrem ao ligar o computador. Desative apenas os que você reconhece e prefere abrir manualmente. Você pode reativá-los na mesma tela."/>
<Button Name="Startup" Content="Revisar aplicativos iniciais"/>
</StackPanel></Border>
<Border Background="White" CornerRadius="8" Padding="18" Margin="0,12,0,0"><StackPanel>
<TextBlock FontSize="19" FontWeight="SemiBold" Text="Reduzir efeitos visuais"/>
<TextBlock Margin="0,5,0,10" Text="Na aba Efeitos Visuais, escolha Personalizar e desmarque as animações que deseja reduzir. Clique em Aplicar. Para voltar, marque as opções novamente."/>
<Button Name="Effects" Content="Abrir efeitos visuais"/>
</StackPanel></Border>
</StackPanel></ScrollViewer></DockPanel></Window>
'@
$window=[Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader $layout))
foreach($name in 'Scan','Clean','Status','Startup','Effects'){Set-Variable -Name $name -Value ($window.FindName($name))}
$script:items=@()
$folder=Join-Path ([Environment]::GetFolderPath('LocalApplicationData')) 'Temp'
function Assert-Temp {
    if([IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd('\') -ne $folder -or $folder.StartsWith('\\')){
        throw 'A pasta temporária é personalizada. Consulte o suporte de TI.'
    }
    $dir=Get-Item -LiteralPath $folder -Force
    while($dir){
        if($dir.Attributes -band [IO.FileAttributes]::ReparsePoint){throw 'A pasta temporária é redirecionada. Consulte o suporte de TI.'}
        $dir=$dir.Parent
    }
}
function Test-Eligible($file,$cut){
    $excluded=[IO.FileAttributes]'Directory,Hidden,ReadOnly,System,ReparsePoint'
    return ($file -is [IO.FileInfo] -and $file.DirectoryName -eq $folder -and
        $file.Extension -eq '.tmp' -and !($file.Attributes -band $excluded) -and
        $file.CreationTimeUtc -lt $cut -and $file.LastWriteTimeUtc -lt $cut)
}
function Format-Size([long]$bytes){
    if($bytes -lt 1KB){return "$bytes bytes"}
    if($bytes -lt 1MB){return ('{0:N1} KB' -f ($bytes/1KB))}
    return ('{0:N1} MB' -f ($bytes/1MB))
}
$Scan.Add_Click({
    $Clean.Visibility='Collapsed'
    $script:items=@()
    try{
        Assert-Temp
        $cut=[datetime]::UtcNow.AddDays(-7)
        $script:items=@(Get-ChildItem -LiteralPath $folder -File -Filter '*.tmp' | Where-Object {Test-Eligible $_ $cut} | Select-Object FullName,Length,CreationTimeUtc,LastWriteTimeUtc)
        if(!$script:items.Count){$Status.Text='Tudo certo. Nenhum temporário com mais de sete dias para limpar.';return}
        $size=Format-Size (($script:items | Measure-Object Length -Sum).Sum)
        $Status.Text="Encontrados: $($script:items.Count) arquivo(s) temporário(s), total de $size. Feche os aplicativos de trabalho antes de excluir."
        $Clean.Visibility='Visible'
    }catch{$Status.Text="Análise interrompida. $($_.Exception.Message)"}
})
$Clean.Add_Click({
    $question="Excluir definitivamente os $($script:items.Count) arquivo(s) temporário(s) identificados? Eles não serão enviados à Lixeira."
    if([Windows.MessageBox]::Show($window,$question,'Confirmar limpeza','YesNo','Question','No') -ne 'Yes'){return}
    $removed=0
    [long]$bytes=0
    $total=$script:items.Count
    try{
        $cut=[datetime]::UtcNow.AddDays(-7)
        foreach($snapshot in $script:items){
            Assert-Temp
            try{
                $f=Get-Item -LiteralPath $snapshot.FullName -Force -ErrorAction Stop
                if((Test-Eligible $f $cut) -and $f.CreationTimeUtc -eq $snapshot.CreationTimeUtc -and $f.LastWriteTimeUtc -eq $snapshot.LastWriteTimeUtc -and $f.Length -eq $snapshot.Length){
                    $f.Delete()
                    $bytes+=$snapshot.Length
                    $removed++
                }
            }catch [System.IO.IOException]{}catch [System.UnauthorizedAccessException]{}catch [System.Management.Automation.ItemNotFoundException]{}
        }
        $Status.Text="Concluído. Excluídos: $removed ($(Format-Size $bytes)). Ignorados: $($total-$removed)."
    }catch{$Status.Text="Limpeza interrompida. Arquivos já excluídos: $removed. $($_.Exception.Message)"}
    finally{$script:items=@();$Clean.Visibility='Collapsed'}
})
$Startup.Add_Click({
    try{Start-Process 'ms-settings:startupapps'}
    catch{[void][Windows.MessageBox]::Show($window,'Abra o Gerenciador de Tarefas com Ctrl + Shift + Esc e procure Inicializar ou Aplicativos de inicialização.','Inicialização')}
})
$Effects.Add_Click({
    try{Start-Process (Join-Path ([Environment]::SystemDirectory) 'SystemPropertiesPerformance.exe')}
    catch{[void][Windows.MessageBox]::Show($window,'Pesquise por Ajustar a aparência e o desempenho do Windows no menu Iniciar.','Efeitos visuais')}
})
[void]$window.ShowDialog()
