
$config = Get-Content 'C:\Projects\gemini-excel-tool\config.json' | ConvertFrom-Json
$Excel = New-Object -ComObject Excel.Application
$Excel.Visible = $false
$Excel.DisplayAlerts = $false

foreach ($id in $config.data.PSObject.Properties.Name) {
    $items = $config.data.$id
    $ghValue = $items[0].combinedGH
    $fullID = $id + "_" + $ghValue
    $fileName = "連携機能設計書（" + $fullID + "）.xlsx"
    $outputPath = Join-Path $config.outputDir $fileName

    # Tìm file thiết kế gốc
    $designFiles = Get-ChildItem -Path $config.designDocsDir -Filter "*$id*"
    if ($designFiles.Count -eq 0) { 
        Write-Host "SKIP: Khong tim thay thiet ke cho $id"
        continue 
    }
    $sourceDesignPath = $designFiles[0].FullName

    Write-Host "Dang xu ly: $fileName"
    
    # Copy Template
    Copy-Item $config.templatePath $outputPath

    # Mo Workbook ket qua va Workbook thiet ke
    $outWB = $Excel.Workbooks.Open($outputPath)
    $srcWB = $Excel.Workbooks.Open($sourceDesignPath)

    try {
        # --- 1. Cap nhat sheet '表紙' ---
        $cover = $outWB.Sheets | Where-Object { $_.Name -like "*表紙*" }
        if ($cover) {
            $cover.Range("A15").Value2 = $fullID
            $cover.Range("A26").Value2 = $config.today
        }

        # --- 2. Cap nhat sheet '個別レイアウト情報' ---
        $layout = $outWB.Sheets | Where-Object { $_.Name -like "*個別レイアウト情報*" }
        if ($layout) {
            $layout.Range("O10").Value2 = $designFiles[0].Name
            $layout.Range("O11").Value2 = $ghValue
            for ($i = 0; $i -lt $items.Count; $i++) {
                $layout.Range("B" + (15 + $i)).Value2 = $items[$i].cotc
            }
        }

        # --- 3. Copy sheet '機能概要' va 'マッピング定義' ---
        $mapping = @{ "機能概要" = "IFA_機能概要"; "マッピング定義" = "IFA_マッピング定義" }
        foreach ($key in $mapping.Keys) {
            $srcSheet = $srcWB.Sheets | Where-Object { $_.Name -like "*$key*" }
            if ($srcSheet) {
                # Xoa sheet cu neu co
                $oldDest = $outWB.Sheets | Where-Object { $_.Name -eq $mapping[$key] }
                if ($oldDest) { $oldDest.Delete() }

                # Copy sheet sang
                $srcSheet.Copy($outWB.Sheets.Item(1))
                $newSheet = $outWB.Sheets.Item(1)
                $newSheet.Name = $mapping[$key]
            }
        }

        $outWB.Save()
        Write-Host "[OK] $fileName"
    } catch {
        Write-Host "[ERR] $id : $($_.Exception.Message)"
    } finally {
        $outWB.Close()
        $srcWB.Close()
    }
}

$Excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel)
