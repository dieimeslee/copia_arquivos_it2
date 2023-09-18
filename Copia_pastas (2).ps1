# Importa os módulos necessários
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName Microsoft.Office.Interop.Excel

# Solicita ao usuário para inserir a data
$date = Read-Host -Prompt 'Por favor, insira a data no formato DD/MM/AAAA'

# Converte a string de data para um objeto datetime
$date = [datetime]::ParseExact($date, 'dd/MM/yyyy', $null)

# Converte o objeto datetime de volta para uma string no formato desejado
$date = $date.ToString('yyyyMMdd')

Write-Output "Data inserida pelo usuário: $date"

# Cria um objeto Excel.Application
$excel = New-Object -ComObject Excel.Application

# Abrindo o arquivo Excel
$workbook = $excel.Workbooks.Open('C:\Users\99837808\OneDrive - Anheuser-Busch InBev\Desktop\1PROJECT\copia pasta it2\bat\caminho_para_arquivo_IT2.xlsx')

# Lendo as sheets
$sheetOrigem = $workbook.Worksheets.Item('origem')
$sheetDestino = $workbook.Worksheets.Item('destino')

# Obtendo os caminhos das pastas origem e destino
$pastasOrigem = @($sheetOrigem.UsedRange.Value2 | Where-Object { $_ -ne $null })
$pastaDestino = $sheetDestino.Cells.Item(1, 1).Text

Write-Output "Pastas de origem: $pastasOrigem"
Write-Output "Pasta de destino: $pastaDestino"

# Fecha o arquivo Excel e libera o objeto COM
$workbook.Close($false)
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null

foreach ($pasta in $pastasOrigem) {
    $encontrouArquivo = $false
    Get-ChildItem -Path $pasta -Recurse | ForEach-Object {
        # Extrai a data do nome do arquivo
        $dataArquivo = $_.BaseName.Substring($_.BaseName.Length - 15, 8)
        Write-Output "Data do arquivo: $dataArquivo"
        # Se a data do arquivo corresponder à data inserida pelo usuário
        if ($dataArquivo -eq $date) {
            $encontrouArquivo = $true
            Copy-Item -Path $_.FullName -Destination $pastaDestino
            Write-Output "Arquivo copiado: $($_.FullName)"
        }
    }
    if (-not $encontrouArquivo) {
        Write-Output "Nenhum arquivo encontrado com a data $date em $pasta"
    }
}

# Aguarda o usuário pressionar Enter para sair
Read-Host -Prompt 'Pressione ENTER para sair'
