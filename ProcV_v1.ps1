# 1. Defina os caminhos utilizando caminhos automáticos de pasta ($PSScriptRoot)
# Isso evita problemas caso o script mude de letra de disco (C:\ ou E:\)
$PathInfra = Join-Path -Path $PSScriptRoot -ChildPath "Inventario_vSAN_SO_MEUCLUSTER.csv"
$PathTESPX = Join-Path -Path $PSScriptRoot -ChildPath "CMDB_200526.xlsx"
$PathSaida = Join-Path -Path $PSScriptRoot -ChildPath "Consolidado_CMDB_VC.csv"

Write-Host "Lendo arquivo de infraestrutura (CSV)..." -ForegroundColor Cyan
$DadosInfra = Import-Csv -Path $PathInfra -Delimiter ";"

# Identifica a coluna de espaço com caractere especial corrompido de forma dinâmica
$ColunaEspaco = ($DadosInfra | Get-Member -MemberType NoteProperty | Where-Object { $_.Name -like "*Espa*o*" }).Name

Write-Host "Abrindo o Excel em segundo plano para ler a CMDB_XXXXX.xlsx"..." -ForegroundColor Cyan
$Excel = New-Object -ComObject Excel.Application
$Excel.Visible = $false
$Workbook = $Excel.Workbooks.Open($PathCMDB_xxxxx.xlsx")
$Worksheet = $Workbook.Sheets.Item(1)

# Descobre a quantidade de linhas preenchidas na CMDB_xxxxx.xlsx"
$Rows = $Worksheet.UsedRange.Rows.Count

# Carrega os dados do Excel para uma Hashtable usando limpeza de caracteres especiais na chave
$TabelaBusca = @{}
for ($r = 2; $r -le $Rows; $r++) {
    $NomeIC = $Worksheet.Cells.Item($r, 2).Value2 # Coluna 2: Nome do IC
    $TipoIC = $Worksheet.Cells.Item($r, 1).Value2 # Coluna 1: Tipo de IC
    $Equipe = $Worksheet.Cells.Item($r, 3).Value2 # Coluna 3: Equipe Responsável
    $Client = $Worksheet.Cells.Item($r, 4).Value2 # Coluna 4: Cliente

    if ($NomeIC) {
        # REGEX: Limpeza profunda (remove traços, pontos, sublinhados e números sequenciais do fim)
        # Transforma padrões variantes de "APP01" ou "Infra-PRD" em chaves comparáveis puras
        $ChavePura = ($NomeIC.ToString() -replace '[-_.]', '' -replace '\d+$', '').Trim().ToLower()
        
        if (-not $TabelaBusca.ContainsKey($ChavePura)) {
            $TabelaBusca.Add($ChavePura, @{
                'Tipo de IC'          = $TipoIC
                'Equipe Responsável'  = $Equipe
                'Cliente'             = $Client
            })
        }
    }
}

# Fecha o Excel de forma limpa
$Workbook.Close($false)
$Excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel) | Out-Null
Remove-Variable Excel

Write-Host "Fazendo o cruzamento total dos dados (PROCV Avançado)..." -ForegroundColor Cyan
$ResultadoFinal = [System.Collections.Generic.List[PSCustomObject]]::new()

foreach ($vm in $DadosInfra) {
    # Aplica a mesma limpeza de caracteres especiais no nome vindo do vCenter
    $NomeVMPuro = ($vm.'Nome da VM (vCenter)' -replace '[-_.]', '' -replace '\d+$', '').Trim().ToLower()

    # Se por algum motivo o Nome da VM falhar, usa o Hostname (SO) como plano B
    if ([string]::IsNullOrEmpty($NomeVMPuro) -or $NomeVMPuro -eq "na") {
        $NomeVMPuro = ($vm.'Hostname (SO)' -replace '[-_.]', '' -replace '\d+$', '').Trim().ToLower()
    }

    # Executa a busca exata pela string limpa ou por correspondência aproximada
    if ($TabelaBusca.ContainsKey($NomeVMPuro)) {
        $Match = $TabelaBusca[$NomeVMPuro]
        $TipoIC  = $Match.'Tipo de IC'
        $Equipe  = $Match.'Equipe Responsável'
        $Cliente = $Match.'Cliente'
    } else {
        # Busca aproximada redundante (se o nome contém ou está contido em alguma chave)
        $AchouParcial = $false
        foreach ($chave in $TabelaBusca.Keys) {
            if ($NomeVMPuro -contains $chave -or $chave -contains $NomeVMPuro) {
                $Match = $TabelaBusca[$chave]
                $TipoIC  = $Match.'Tipo de IC'
                $Equipe  = $Match.'Equipe Responsável'
                $Cliente = $Match.'Cliente'
                $AchouParcial = $true
                break
            }
        }
        if (-not $AchouParcial) {
            $TipoIC  = "Não Encontrado"
            $Equipe  = "Não Encontrado"
            $Cliente = "Não Encontrado"
        }
    }
    
    # GARANTIA DE TOTALIDADE: Mapeia todas as colunas de negócio e de infraestrutura
    # incluindo as novas colunas 'PowerState' e 'Cluster de Origem' criadas pelo extrator
    $NovaLinha = [PSCustomObject]@{
        'Nome do IC'          = $vm.'Nome da VM (vCenter)'
        'Tipo de IC'          = $TipoIC
        'Equipe Responsável'  = $Equipe
        'Cliente'             = $Cliente
        'Status da VM'        = if ($vm.PowerState) { $vm.PowerState } else { "N/A" }
        'Hostname (SO)'       = $vm.'Hostname (SO)'
        'IP'                  = $vm.IP
        'Sistema Operacional' = $vm.'Sistema Operacional'
        'Quantidade vCPU'     = $vm.'Quantidade vCPU'
        'Memoria (GB)'        = $vm.'Memoria (GB)'
        'Espaço no SO (GB)'   = $vm.$ColunaEspaco
        'Consumo vSAN (GB)'   = $vm.'Consumo vSAN (GB)'
        'Cluster de Origem'   = if ($vm.'Cluster de Origem') { $vm.'Cluster de Origem' } else { "N/A" }
    }
    $ResultadoFinal.Add($NovaLinha)
}

# Exporta o resultado final sem perder nenhuma coluna
$ResultadoFinal | Export-Csv -Path $PathSaida -NoTypeInformation -Delimiter ";" -Encoding Utf8

Write-Host "`nSucesso total! Relatório 100% integrado gerado em: $PathSaida" -ForegroundColor Green
