from pathlib import Path

text = r"""$ErrorActionPreference = 'Stop'
$src = '\\bcc\eng\PrDoc\377_НИПИГАЗ\АГХК\КСБ\RFP_MTO_VO\_РЕЗУЛЬТАТА_ПРОВЕРКИ\_результат_проверки_2026.08.18.16.55'
$dst = 'c:\Users\ydruzev\PycharmProjects\Documentation_PDF_out_NK\tmp\inspect_20260818'
New-Item -ItemType Directory -Force -Path $dst | Out-Null
Get-ChildItem -LiteralPath $src | ForEach-Object {
    Write-Output ('SRC ' + $_.Name + ' ' + $_.Length)
    Copy-Item -LiteralPath $_.FullName -Destination (Join-Path $dst $_.Name) -Force
}
$reports = '\\bcc\eng\PrDoc\377_НИПИГАЗ\АГХК\КСБ\RFP_MTO_VO\_RFP\RFP сводный файл'
Write-Output '--- reports dirs ---'
Get-ChildItem -LiteralPath $reports -Directory | Sort-Object Name -Descending | Select-Object -First 10 | ForEach-Object {
    $net = Join-Path $_.FullName 'rfp_parts_net.xlsx'
    $netExists = Test-Path -LiteralPath $net
    $netLen = if ($netExists) { (Get-Item -LiteralPath $net).Length } else { 0 }
    Write-Output ($_.Name + ' net=' + $netExists + ' size=' + $netLen)
}
Get-ChildItem -LiteralPath $reports -Directory | Where-Object { $_.Name -like '2026.08.18*' } | ForEach-Object {
    $net = Join-Path $_.FullName 'rfp_parts_net.xlsx'
    if (Test-Path -LiteralPath $net) {
        $name = 'rfp_parts_net__' + $_.Name + '.xlsx'
        Copy-Item -LiteralPath $net -Destination (Join-Path $dst $name) -Force
        Write-Output ('copied dated net ' + $name)
    }
}
Write-Output '--- dst ---'
Get-ChildItem -LiteralPath $dst | ForEach-Object { Write-Output ($_.Name + ' ' + $_.Length) }
"""
Path(__file__).with_name("copy_inspect_20260818.ps1").write_text(text, encoding="utf-8-sig")
print("ok")
