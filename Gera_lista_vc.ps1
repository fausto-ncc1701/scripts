# Ignora erros de certificado SSL
Set-PowerCLIConfiguration -InvalidCertificateAction Ignore -Confirm:$false | Out-Null

# Solicita o nome do Datacenter (pode digitar parcial, ex: datacenter01, DATACENTER0)
$DCInput = Read-Host "Digite o nome (ou parte do nome) do DATACENTER no vCenter"

# Verifica se há uma conexão ativa com o vCenter
if (-not $global:DefaultVIServer) {
    $vCenter = Read-Host "Digite o FQDN ou IP do vCenter"
    Connect-VIServer -Server $vCenter
}

# BUSCA FLEXÍVEL DE DATACENTER: Procura usando caracteres curinga (*)
$DCEncontrado = Get-Datacenter -Name "*$DCInput*" -ErrorAction SilentlyContinue

if (-not $DCEncontrado) {
    Write-Host "[ERRO] Nenhum Datacenter correspondente a '*$DCInput*' foi encontrado!" -ForegroundColor Red
    break
} 
elseif ($DCEncontrado -is [array]) {
    Write-Host "`n[AVISO] Mais de um Datacenter encontrado. Usando o primeiro: $($DCEncontrado[0].Name)" -ForegroundColor Yellow
    $DCEncontrado = $DCEncontrado[0]
}

Write-Host "`n[INFO] Minerando dados no Datacenter: $($DCEncontrado.Name)... Aguarde.`n" -ForegroundColor Cyan

# Coleta todas as VMs pertencentes a esse Datacenter
$VMs = $DCEncontrado | Get-VM
$Relatorio = [System.Collections.Generic.List[PSCustomObject]]::new()

foreach ($VM in $VMs) {
    
    # Descobre o cluster da VM. Se ela estiver fora de um cluster (direto no Host), marca como 'Standalone'
    $ClusterNome = if ($VM.VMHost.Parent.GetType().Name -like "*Cluster*") { $VM.VMHost.Parent.Name } else { "Fora de Cluster (Standalone)" }

    # 1. STATUS DE ENERGIA (Ligada/Desligada)
    $StatusServidor = if ($VM.PowerState -eq "PoweredOn") { "Ligada" } else { "Desligada" }

    # 2. VALOR ENTREGUE AO SO (Soma das partições via VMware Tools)
    $DiscosNoSO = $VM.Guest.Disks
    $CapacidadeNoSOGB = 0
    
    if ($VM.PowerState -eq "PoweredOn" -and $DiscosNoSO) {
        foreach ($Disco in $DiscosNoSO) {
            # O vVMware Tools reporta em Bytes, convertemos para GB
            $CapacidadeNoSOGB += $Disco.Capacity / 1GB
        }
        $CapacidadeNoSOGB = [math]::Round($CapacidadeNoSOGB, 2)
    } else {
        $CapacidadeNoSOGB = "VM Desligada ou sem dados do Tools"
    }

    # 3. ESPAÇO FÍSICO ALOCADO (Funciona para vSAN e storages tradicionais)
    $ConsumoFisicoGB = [math]::Round(($VM.UsedSpaceGB), 2)

    # Tratamento para IPs
    $IPPrincipal = "N/A"
    if ($VM.Guest.IPAddress) { $IPPrincipal = $VM.Guest.IPAddress[0] }

    # Monta a estrutura da planilha
    $InfoVM = [PSCustomObject]@{
        "Nome da VM (vCenter)"  = $VM.Name
        "PowerState"            = $StatusServidor       
        "Hostname (SO)"         = if ($VM.Guest.HostName) { $VM.Guest.HostName } else { "N/A" }
        "IP"                    = $IPPrincipal
        "Sistema Operacional"   = if ($VM.Guest.OSDescription) { $VM.Guest.OSDescription } else { "N/A" }
        "Quantidade vCPU"       = $VM.NumCpu
        "Memoria (GB)"          = $VM.MemoryGB
        "Espaço no SO (GB)"     = $CapacidadeNoSOGB      
        "Consumo vSAN (GB)"     = $ConsumoFisicoGB
        "Cluster de Origem"     = $ClusterNome         # <-- Adicionado ao fim da planilha, conforme solicitado!
    }
    
    $Relatorio.Add($InfoVM)
}

# Exibe na tela em formato de tabela interativa unificada do Windows
$Relatorio | Out-GridView -Title "Inventário vCenter - Datacenter $($DCEncontrado.Name)"

# CORREÇÃO DO CAMINHO: Salva AUTOMATICAMENTE na mesma pasta onde este script estiver rodando
$PathSaida = Join-Path -Path $PSScriptRoot -ChildPath "Inventario_vSAN_SO_MEUCLUSTER.csv"

# Exportação em CSV
$Relatorio | Export-Csv -Path $PathSaida -NoTypeInformation -Delimiter ";" -Encoding Utf8

# Mensagem de sucesso limpa no console
Write-Host ""
Write-Host "[SUCESSO] Extração concluída!" -ForegroundColor Green
Write-Host "O arquivo foi gerado em: $PathSaida" -ForegroundColor White