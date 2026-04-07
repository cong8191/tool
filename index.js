const { execSync } = require('child_process');
const ExcelJS = require('exceljs');
const fs = require('fs');
const path = require('path');

async function processInterfaces(inputIDs = null) {
    const sourcePath = path.resolve('01.要件定義_インターフェース一覧（STEP3）.xlsx');
    const templatePath = path.resolve('template_FtoF.xlsx');
    const designDocsDir = 'D:\\Project\\151_ISA_AsteriaWrap\\trunk\\99_FromJP\\10_プロジェクト資材\\02.IFAgreement\\03.確定';
    const outputDir = __dirname;

    try {
        // 1. Dùng ExcelJS để đọc dữ liệu từ file Master (rất nhanh)
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(sourcePath);
        const sheet = workbook.getWorksheet('【STEP3】インターフェース一覧');

        const results = {}; 
        const searchList = Array.isArray(inputIDs) ? inputIDs : (inputIDs ? [inputIDs] : null);

        const getVal = (cell) => {
            if (cell.value && typeof cell.value === 'object') {
                return cell.value.result !== undefined ? cell.value.result : cell.value;
            }
            return cell.value;
        };

        sheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
            if (rowNumber < 21) return;
            const id = getVal(row.getCell(67))?.toString().trim();
            if (!id || id === '-') return;
            if (searchList && !searchList.includes(id)) return;

            if (!results[id]) results[id] = [];
            const rawG = getVal(row.getCell(7));
            const rawH = getVal(row.getCell(8));
            const paddedG = (rawG !== null && rawG !== undefined ? rawG : '').toString().padStart(4, '0');
            const combinedGH = `${paddedG}-${rawH !== null && rawH !== undefined ? rawH : ''}`;

            results[id].push({ 
                cotc: getVal(row.getCell(3)), 
                cotI: getVal(row.getCell(9)),
                combinedGH 
            });
        });

        // 2. Tạo file cấu hình JSON tạm thời để PowerShell đọc
        const configPath = path.join(__dirname, 'config.json');
        fs.writeFileSync(configPath, JSON.stringify({
            templatePath,
            designDocsDir,
            outputDir,
            data: results,
            today: new Date().toLocaleDateString('ja-JP')
        }, null, 2));

        // 3. Tạo Script PowerShell để thực hiện các thao tác Excel "nặng"
        const psScript = `
$config = Get-Content '${configPath.replace(/'/g, "''")}' | ConvertFrom-Json
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
            $layout.Range("J16").Value2 = $items[0].cotI
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
`;

        const psPath = path.join(__dirname, 'process.ps1');
        fs.writeFileSync(psPath, psScript, { encoding: 'utf8' });

        // 4. Chạy PowerShell
        console.log('--- Bắt đầu chạy Excel COM (PowerShell) ---');
        execSync(`powershell -ExecutionPolicy Bypass -File "${psPath}"`, { stdio: 'inherit' });

    } catch (error) {
        console.error('Lỗi hệ thống:', error.message);
    }
}

const args = process.argv.slice(2);
processInterfaces(args.length > 0 ? args : null);
