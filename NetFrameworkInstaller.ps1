# .NET Framework 参考程序集安装工具（版本 4.0 - 4.8）
#
# 功能：
# 1. 在选择版本界面标识当前电脑已安装的版本
# 2. 如果选择已安装的版本，提示覆盖或跳过
# 3. 简洁美观的控制台界面
# 4. 详细的安装状态反馈

#requires -RunAsAdministrator
Add-Type -AssemblyName System.Security
Add-Type -AssemblyName System.IO.Compression.FileSystem

function Test-Admin {
    $currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
    $adminRole = [Security.Principal.WindowsBuiltInRole]::Administrator
    $principal = New-Object Security.Principal.WindowsPrincipal($currentUser)
    return $principal.IsInRole($adminRole)
}

function Write-ColorText {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Text,
        [Parameter(Mandatory = $false)]
        [string]$Color = "White",
        [switch]$NoNewLine
    )

    $origColor = $host.UI.RawUI.ForegroundColor
    try {
        $host.UI.RawUI.ForegroundColor = $Color
        if ($NoNewLine) {
            Write-Host $Text -NoNewline
        } else {
            Write-Host $Text
        }
    } finally {
        $host.UI.RawUI.ForegroundColor = $origColor
    }
}

function Show-Divider {
    try {
        $consoleWidth = $host.UI.RawUI.WindowSize.Width
    } catch {
        $consoleWidth = 80
    }
    if ($consoleWidth -lt 1) { $consoleWidth = 80 }
    $divider = "=" * $consoleWidth
    if (![string]::IsNullOrWhiteSpace($divider)) {
        Write-ColorText $divider -Color "DarkCyan"
    }
}

function Get-InstalledVersions {
    $refBase = Join-Path ${env:ProgramFiles(x86)} "Reference Assemblies\Microsoft\Framework\.NETFramework"
    $installed = @()
    if (Test-Path $refBase) {
        Get-ChildItem -Path $refBase -Directory | ForEach-Object {
            if (Test-Path (Join-Path $_.FullName "mscorlib.dll")) {
                $installed += $_.Name
            }
        }
    }
    return $installed
}

function Test-VersionInstalled {
    param(
        [string]$Version,
        [hashtable]$VersionMap
    )
    $refDir = Join-Path ${env:ProgramFiles(x86)} "Reference Assemblies\Microsoft\Framework\.NETFramework"
    $targetDir = Join-Path $refDir $VersionMap[$Version].TargetDir
    $targetFile = Join-Path $targetDir "mscorlib.dll"
    return (Test-Path $targetFile)
}

function Show-VersionMenu {
    param(
        [hashtable]$VersionMap
    )
    $installed = Get-InstalledVersions
    Show-Divider
    Write-ColorText "  .NET Framework 参考程序集安装工具" -Color "Green"
    Write-ColorText "  版本范围: 4.0 - 4.8" -Color "Yellow"
    Show-Divider
    Write-Host ""
    $index = 1
    foreach ($key in ($VersionMap.Keys | Sort-Object)) {
        $status = if ($installed -contains $VersionMap[$key].TargetDir) { "[已安装]" } else { "" }
        $color = if ($status) { "Cyan" } else { "White" }
        Write-ColorText "  $index. .NET Framework $key" -Color $color -NoNewLine
        Write-ColorText " $status" -Color "Green"
        $index++
    }
    Write-ColorText "  $index. 退出"
    Write-Host ""
    Show-Divider
    return $installed
}

function Install-ReferenceAssemblies {
    param(
        [hashtable]$VersionMap,
        [string]$SelectedVersion,
        [switch]$Force
    )

    $host.UI.RawUI.WindowTitle = "正在安装 .NET Framework $SelectedVersion 参考程序集..."
    $refBase = Join-Path ${env:ProgramFiles(x86)} "Reference Assemblies\Microsoft\Framework\.NETFramework"
    $refDir = Join-Path $refBase $VersionMap[$SelectedVersion].TargetDir
    $targetFile = Join-Path $refDir "mscorlib.dll"

    if (-not $Force -and (Test-Path $targetFile)) {
        Write-ColorText "`n  [!] $SelectedVersion 版本已安装!" -Color "Yellow"
        return $true
    }

    try {
        $tempDir = New-Item -ItemType Directory -Path (Join-Path $env:TEMP ".net-ref-$SelectedVersion") -Force
        $pkg = $VersionMap[$SelectedVersion].PackageName
        $targetDir = $VersionMap[$SelectedVersion].TargetDir
        $url = "https://www.nuget.org/api/v2/package/$pkg"
        $alt = "https://globalcdn.nuget.org/packages/$pkg.latest.nupkg"
        $pkgPath = Join-Path $tempDir "$pkg.nupkg"

        Write-Host "`n  [步骤1/4] 下载包 ($pkg)..." -ForegroundColor Cyan
        $ProgressPreference = 'SilentlyContinue'
        try {
            Invoke-WebRequest -Uri $url -OutFile $pkgPath -ErrorAction Stop
        } catch {
            Invoke-WebRequest -Uri $alt -OutFile $pkgPath -ErrorAction Stop
        }

        if (-not (Test-Path $pkgPath)) { throw "无法下载包: $pkg" }
        $size = [math]::Round((Get-Item $pkgPath).Length / 1MB, 2)
        Write-ColorText "  ✓ 下载完成 ($size MB)" -Color "Green"

        Write-Host "`n  [步骤2/4] 解压缩文件..." -ForegroundColor Cyan
        $extractDir = Join-Path $tempDir "extracted"
        [System.IO.Compression.ZipFile]::ExtractToDirectory($pkgPath, $extractDir)

        if ($Force -and (Test-Path $refDir)) {
            Write-Host "  - 删除目录: $refDir" -ForegroundColor DarkGray
            Remove-Item -Path $refDir -Recurse -Force -ErrorAction Stop
        }

        Write-Host "`n  [步骤3/4] 复制文件..." -ForegroundColor Cyan
        $srcDir = Join-Path $extractDir "build\.NETFramework\$($targetDir)"
        New-Item -Path $refDir -ItemType Directory -Force | Out-Null
        Copy-Item -Path "$srcDir\*" -Destination $refDir -Recurse -Force

        Write-Host "`n  [步骤4/4] 验证安装..." -ForegroundColor Cyan
        if (-not (Test-Path $targetFile)) {
            throw "验证失败: $targetFile 未找到"
        }

        $msg = if ($Force) {
            "✓ .NET Framework $SelectedVersion 已更新!"
        } else {
            "✓ .NET Framework $SelectedVersion 安装成功!"
        }
        Write-ColorText "`n  $msg" -Color "Green"
        return $true

    } catch {
        Write-ColorText "`n  安装失败: $($_.Exception.Message)" -Color "Red"
        return $false
    } finally {
        if ($null -ne $tempDir -and (Test-Path $tempDir)) {
           Remove-Item -Path $tempDir -Recurse -Force -ErrorAction SilentlyContinue
        }
        $host.UI.RawUI.WindowTitle = ".NET Framework 参考程序集安装工具"
    }
}

function Show-Warning {
    param(
        [string]$Message,
        [string]$Color = "Red"
    )
    for ($i = 0; $i -lt 3; $i++) {
        Write-ColorText "! $Message !" -Color $Color
        Start-Sleep -Milliseconds 300
    }
    Write-Host ""
}

function Uninstall-ReferenceAssemblies {
    param(
        [hashtable]$VersionMap,
        [string]$SelectedVersion
    )

    $refBase = Join-Path ${env:ProgramFiles(x86)} "Reference Assemblies\Microsoft\Framework\.NETFramework"
    $refDir = Join-Path $refBase $VersionMap[$SelectedVersion].TargetDir
    
    if (-not (Test-Path $refDir)) {
        Write-ColorText "  [!] .NET Framework $SelectedVersion 参考程序集未找到，无需卸载" -Color "Yellow"
        return $true
    }

    try {
        Write-Host "`n  [步骤1/2] 卸载 $SelectedVersion 版本..." -ForegroundColor Cyan
        Remove-Item -Path $refDir -Recurse -Force -ErrorAction Stop
        Write-ColorText "  ✓ 程序集文件已移除" -Color "Green"

        Write-Host "`n  [步骤2/2] 验证卸载结果..." -ForegroundColor Cyan
        if (Test-Path $refDir) {
            throw "卸载不彻底: $refDir 仍然存在"
        }
        
        Write-ColorText "`n  ✓ .NET Framework $SelectedVersion 卸载成功!" -Color "Green"
        return $true
    } catch {
        Write-ColorText "`n  卸载失败: $($_.Exception.Message)" -Color "Red"
        return $false
    }
}

function Main {
    Clear-Host

    $versionMap = [ordered]@{
        "4.0" = @{ PackageName = "Microsoft.NETFramework.ReferenceAssemblies.net40"; TargetDir = "v4.0" }
        "4.5" = @{ PackageName = "Microsoft.NETFramework.ReferenceAssemblies.net45"; TargetDir = "v4.5" }
        "4.5.1" = @{ PackageName = "Microsoft.NETFramework.ReferenceAssemblies.net451"; TargetDir = "v4.5.1" }
        "4.5.2" = @{ PackageName = "Microsoft.NETFramework.ReferenceAssemblies.net452"; TargetDir = "v4.5.2" }
        "4.6" = @{ PackageName = "Microsoft.NETFramework.ReferenceAssemblies.net46"; TargetDir = "v4.6" }
        "4.6.1" = @{ PackageName = "Microsoft.NETFramework.ReferenceAssemblies.net461"; TargetDir = "v4.6.1" }
        "4.6.2" = @{ PackageName = "Microsoft.NETFramework.ReferenceAssemblies.net462"; TargetDir = "v4.6.2" }
        "4.7" = @{ PackageName = "Microsoft.NETFramework.ReferenceAssemblies.net47"; TargetDir = "v4.7" }
        "4.7.1" = @{ PackageName = "Microsoft.NETFramework.ReferenceAssemblies.net471"; TargetDir = "v4.7.1" }
        "4.7.2" = @{ PackageName = "Microsoft.NETFramework.ReferenceAssemblies.net472"; TargetDir = "v4.7.2" }
        "4.8" = @{ PackageName = "Microsoft.NETFramework.ReferenceAssemblies.net48"; TargetDir = "v4.8" }
        "4.8.1" = @{ PackageName = "Microsoft.NETFramework.ReferenceAssemblies.net481"; TargetDir = "v4.8.1" }
    }

    $host.UI.RawUI.WindowTitle = ".NET Framework 参考程序集安装工具"
    while ($true) {
        Clear-Host
        $installed = Show-VersionMenu -VersionMap $versionMap
        $exitChoice = $versionMap.Count + 1
        Write-Host ""
        $choice = Read-Host "  请选择要安装的版本 (1-$exitChoice)"

        if ($choice -eq $exitChoice.ToString()) {
            Write-Host "`n  感谢使用! 退出程序..." -ForegroundColor Green
            break
        } elseif ($choice -notmatch "^\d+$" -or [int]$choice -lt 1 -or [int]$choice -gt $versionMap.Count) {
            Write-ColorText "`n  [!] 无效选择，请输入 1 到 $($versionMap.Count) 之间的数字`n" -Color "Red"
            Start-Sleep -Seconds 1
            continue
        }

        $index = 1
        foreach ($key in $versionMap.Keys) {
            if ($index -eq [int]$choice) {
                $selected = $key
                break
            }
            $index++
        }

        if (Test-VersionInstalled -Version $selected -VersionMap $versionMap) {
            $result = $false
            Write-ColorText "`n  [!] .NET Framework $selected 已安装!" -Color "Yellow"
            Write-ColorText "`n  请选择操作:" -Color "Cyan"
            Write-ColorText "  1. 覆盖安装" -Color "White"
            Write-ColorText "  2. 跳过安装" -Color "White"
            Write-ColorText "  3. 卸载安装" -Color "White"
            Write-ColorText "  4. 取消操作" -Color "White"
            $action = Read-Host "  请选择操作 (1-4)"

            switch ($action) {
                "1" { $result = Install-ReferenceAssemblies -VersionMap $versionMap -SelectedVersion $selected -Force }
                "2" { Write-ColorText "  ✓ 保留现有版本" -Color "Green" 
                      $result = $true
                    }
                "3" { $result = Uninstall-ReferenceAssemblies -VersionMap $versionMap -SelectedVersion $selected }
                default { $result = $null
                          continue 
                        }
            }
        } else {
            $result = Install-ReferenceAssemblies -VersionMap $versionMap -SelectedVersion $selected
        }

        # 处理安装结果
        if ($result -eq $false) {
            Write-ColorText "`n  [!] 安装失败，按回车键返回主菜单..." -Color "Red"
            Read-Host  # 失败时暂停等待确认
        } elseif ($result -eq $true) {
            Start-Sleep -Milliseconds 800  # 成功时短暂暂停
        }
    }
}

Main
