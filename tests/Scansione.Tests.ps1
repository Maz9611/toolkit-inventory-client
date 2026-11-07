BeforeAll {
    . (Join-Path $PSScriptRoot 'Inizializza.ps1')
    function Read-Scansione([string]$Testo) {
        InModuleScope HelpDeskToolkit -Parameters @{ T = $Testo } { param($T) ConvertFrom-HDScansione -Testo $T }
    }
}
AfterAll { . (Join-Path $PSScriptRoot 'Termina.ps1') }

Describe 'Export CSV di Advanced IP Scanner' {
    It 'inglese, separatore virgola: tiene gli attivi e scarta gli spenti' {
        $s = Read-Scansione (@(
                'Status,Name,IP,Manufacturer,MAC address,Comments'
                '"On","pc-01.studio.local","192.168.1.20","Intel Corporate","00:11:22:33:44:55",""'
                '"On","NPI3D4E5F","192.168.1.50","Hewlett Packard","AA:BB:CC:DD:EE:01","stampante, piano terra"'
                '"Off","","192.168.1.77","","",""'
            ) -join "`r`n")
        $s.Formato | Should -Match 'CSV'
        $s.Dispositivi | Should -HaveCount 2
        $s.Ignorati | Should -Be 1
        $s.Dispositivi[0].Hostname | Should -Be 'PC-01'
        $s.Dispositivi[0].IP | Should -Be '192.168.1.20'
        $s.Dispositivi[0].MAC | Should -Be '00:11:22:33:44:55'
        $s.Dispositivi[0].Produttore | Should -Be 'Intel Corporate'
        $s.Dispositivi[1].Commenti | Should -Be 'stampante, piano terra'
    }

    It 'italiano, separatore punto e virgola, MAC con i trattini' {
        $s = Read-Scansione (@(
                'Stato;Nome;IP;Produttore;Indirizzo MAC;Commenti'
                'Acceso;PC-RECEPTION;192.168.1.30;Dell Inc.;10-20-30-40-50-6a;'
                'Spento;;192.168.1.31;;;'
            ) -join "`n")
        $s.Dispositivi | Should -HaveCount 1
        $s.Dispositivi[0].MAC | Should -Be '10:20:30:40:50:6A'
        $s.Formato | Should -Match "';'"
    }

    It 'tabulazione, righe di titolo prima dell''intestazione e campi su più righe' {
        $s = Read-Scansione ("Risultati scansione`r`n`r`nName`tIP`tMAC address`tComments`r`nPC-02`t10.0.0.2`t001122334466`t""riga uno`r`nriga due""`r`n")
        $s.Dispositivi | Should -HaveCount 1
        $s.Dispositivi[0].MAC | Should -Be '00:11:22:33:44:66'
        $s.Dispositivi[0].Commenti | Should -Be "riga uno`r`nriga due"
    }

    It 'senza colonna stato tiene solo i dispositivi con nome o MAC' {
        $s = Read-Scansione ("IP,Name,MAC`n192.168.1.1,router,`n192.168.1.9,,`n")
        $s.Dispositivi | Should -HaveCount 1
        $s.Dispositivi[0].Hostname | Should -Be 'ROUTER'
    }

    It 'elimina i doppioni (stesso MAC)' {
        $s = Read-Scansione ("Name,IP,MAC address`nPC-01,192.168.1.20,00:11:22:33:44:55`nPC-01,192.168.1.21,00-11-22-33-44-55`n")
        $s.Dispositivi | Should -HaveCount 1
    }

    It 'legge il file anche se salvato in UTF-16 ("Unicode")' {
        $p = Join-Path $TestDrive 'scansione-utf16.csv'
        [System.IO.File]::WriteAllText($p, "Status,Name,IP,Manufacturer,MAC address`r`nOn,PC-È,192.168.1.40,Acer,00:11:22:33:44:77`r`n", [System.Text.Encoding]::Unicode)
        $testo = InModuleScope HelpDeskToolkit -Parameters @{ P = $p } { param($P) Read-HDFileTesto $P }
        (Read-Scansione $testo).Dispositivi[0].Hostname | Should -Be 'PC-È'
    }

    It 'un file che non è una scansione dà un errore chiaro' {
        { Read-Scansione "Nome,Cognome`nMario,Rossi`n" } | Should -Throw '*Non riconosco le colonne*'
    }
}

Describe 'Export XML' {
    It 'righe con attributi' {
        $s = Read-Scansione '<?xml version="1.0"?><results><row status="on" name="PC-10" ip="192.168.1.10" mac="00:11:22:33:44:10" manufacturer="Intel"/><row status="off" name="" ip="192.168.1.11" mac="" manufacturer=""/></results>'
        $s.Formato | Should -Be 'XML'
        $s.Dispositivi | Should -HaveCount 1
        $s.Dispositivi[0].Hostname | Should -Be 'PC-10'
    }

    It 'righe con elementi figli' {
        $s = Read-Scansione '<scan><host><Name>pc-11.local</Name><IPAddress>192.168.1.12</IPAddress><MACAddress>00-11-22-33-44-12</MACAddress><Vendor>Dell</Vendor></host></scan>'
        $s.Dispositivi | Should -HaveCount 1
        $s.Dispositivi[0].Hostname | Should -Be 'PC-11'
        $s.Dispositivi[0].Produttore | Should -Be 'Dell'
        $s.Dispositivi[0].MAC | Should -Be '00:11:22:33:44:12'
    }
}

Describe 'Normalizzazioni' {
    It 'MAC <Testo> -> <Atteso>' -TestCases @(
        @{ Testo = '00-11-22-aa-bb-cc'; Atteso = '00:11:22:AA:BB:CC' }
        @{ Testo = '0011.22aa.bbcc'; Atteso = '00:11:22:AA:BB:CC' }
        @{ Testo = '001122AABBCC'; Atteso = '00:11:22:AA:BB:CC' }
        @{ Testo = '00:00:00:00:00:00'; Atteso = '' }
        @{ Testo = 'boh'; Atteso = '' }
    ) {
        InModuleScope HelpDeskToolkit -Parameters @{ T = $Testo } { param($T) ConvertTo-HDMac $T } | Should -Be $Atteso
    }

    It 'hostname: <Testo> -> <Atteso>' -TestCases @(
        @{ Testo = 'pc-01.studio.local'; Atteso = 'PC-01' }
        @{ Testo = ' PC-02 '; Atteso = 'PC-02' }
        @{ Testo = '192.168.1.5'; Atteso = '' }
    ) {
        InModuleScope HelpDeskToolkit -Parameters @{ T = $Testo } { param($T) ConvertTo-HDNomeHost $T } | Should -Be $Atteso
    }
}

Describe 'Calcoli sulla rete' {
    It 'rete <Testo> -> <Rete>, <Indirizzi> indirizzi' -TestCases @(
        @{ Testo = '192.168.1.37/24'; Rete = '192.168.1.0/24'; Indirizzi = 254; Primo = '192.168.1.1'; Ultimo = '192.168.1.254' }
        @{ Testo = '10.0.5.200/22'; Rete = '10.0.4.0/22'; Indirizzi = 1022; Primo = '10.0.4.1'; Ultimo = '10.0.7.254' }
        @{ Testo = ' 172.16.0.9 / 30 '; Rete = '172.16.0.8/30'; Indirizzi = 2; Primo = '172.16.0.9'; Ultimo = '172.16.0.10' }
    ) {
        $i = InModuleScope HelpDeskToolkit -Parameters @{ T = $Testo } { param($T) Get-HDInfoRete $T }
        $i.Rete | Should -Be $Rete
        $i.Indirizzi | Should -Be $Indirizzi
        InModuleScope HelpDeskToolkit -Parameters @{ N = $i.Primo } { param($N) ConvertFrom-HDNumeroIp $N } | Should -Be $Primo
        InModuleScope HelpDeskToolkit -Parameters @{ N = $i.Ultimo } { param($N) ConvertFrom-HDNumeroIp $N } | Should -Be $Ultimo
    }

    It 'rifiuta reti troppo grandi o scritte male' {
        InModuleScope HelpDeskToolkit {
            { Get-HDInfoRete '10.0.0.0/16' } | Should -Throw '*troppi*'
            { Get-HDInfoRete '192.168.1.0' } | Should -Throw '*non valida*'
        }
    }

    It 'maschera, elenco degli indirizzi e appartenenza alla rete' {
        InModuleScope HelpDeskToolkit {
            Get-HDPrefissoDaMaschera '255.255.255.0' | Should -Be 24
            Get-HDPrefissoDaMaschera '255.255.252.0' | Should -Be 22
            $info = Get-HDInfoRete '192.168.1.0/29'
            @(Get-HDIndirizziRete $info) | Should -Be @('192.168.1.1', '192.168.1.2', '192.168.1.3', '192.168.1.4', '192.168.1.5', '192.168.1.6')
            Test-HDIpNellaRete -Ip '192.168.1.6' -Info $info | Should -BeTrue
            Test-HDIpNellaRete -Ip '192.168.1.7' -Info $info | Should -BeFalse
            Test-HDIpNellaRete -Ip 'boh' -Info $info | Should -BeFalse
        }
    }
}

Describe 'Cosa scansionare (casella "Rete")' {
    It '<Testo> -> <Rete>, <Indirizzi> indirizzi' -TestCases @(
        @{ Testo = '192.168.1.37/24'; Rete = '192.168.1.0/24'; Indirizzi = 254 }
        @{ Testo = '192.168.1.0/24, 192.168.2.0/24'; Rete = '192.168.1.0/24, 192.168.2.0/24'; Indirizzi = 508 }
        @{ Testo = '10.0.0.1 - 10.0.0.50; 10.0.1.10-20'; Rete = '10.0.0.1-10.0.0.50, 10.0.1.10-20'; Indirizzi = 61 }
        @{ Testo = '10.0.0.15'; Rete = '10.0.0.15'; Indirizzi = 1 }
        @{ Testo = '10.10.0.0/20'; Rete = '10.10.0.0/20'; Indirizzi = 4094 }
    ) {
        $i = InModuleScope HelpDeskToolkit -Parameters @{ T = $Testo } { param($T) Get-HDInfoScansione $T }
        $i.Rete | Should -Be $Rete
        $i.Indirizzi | Should -Be $Indirizzi
    }

    It 'elenca gli indirizzi di più reti senza doppioni e riconosce chi ne fa parte' {
        InModuleScope HelpDeskToolkit {
            $info = Get-HDInfoScansione '192.168.1.1-3, 192.168.1.2-4, 10.0.0.9'
            @(Get-HDIndirizziRete $info) | Should -Be @('192.168.1.1', '192.168.1.2', '192.168.1.3', '192.168.1.4', '10.0.0.9')
            Test-HDIpNellaRete -Ip '10.0.0.9' -Info $info | Should -BeTrue
            Test-HDIpNellaRete -Ip '192.168.1.4' -Info $info | Should -BeTrue
            Test-HDIpNellaRete -Ip '192.168.1.5' -Info $info | Should -BeFalse
        }
    }

    It 'rifiuta testi sbagliati e scansioni troppo grandi' {
        InModuleScope HelpDeskToolkit {
            { Get-HDInfoScansione '' } | Should -Throw
            { Get-HDInfoScansione 'rete di casa' } | Should -Throw '*non valida*'
            { Get-HDInfoScansione '192.168.1.300' } | Should -Throw '*non valido*'
            { Get-HDInfoScansione '10.0.0.50-10.0.0.1' } | Should -Throw '*più alto*'
            { Get-HDInfoScansione '10.0.0.0/16' } | Should -Throw '*troppi*'
            { Get-HDInfoScansione '10.0.0.0/20, 10.0.16.0/24' } | Should -Throw '*troppi*'
        }
    }
}

Describe 'ARP e NetBIOS' {
    It 'legge "arp -a" (anche in italiano) e scarta broadcast e multicast' {
        $voci = @(InModuleScope HelpDeskToolkit {
                ConvertFrom-HDTabellaArp -Righe @(
                    'Interfaccia: 192.168.1.37 --- 0x7'
                    '  Indirizzo Internet    Indirizzo fisico      Tipo'
                    '  192.168.1.1           a0-b1-c2-d3-e4-f5     dinamico'
                    '  192.168.1.20          00-11-22-33-44-55     dinamico'
                    '  192.168.1.255         ff-ff-ff-ff-ff-ff     statico'
                    '  224.0.0.22            01-00-5e-00-00-16     statico'
                )
            })
        $voci | Should -HaveCount 2
        $voci[0].MAC | Should -Be 'A0:B1:C2:D3:E4:F5'
    }

    It 'la richiesta NetBIOS è quella standard (nome jolly "*")' {
        $q = InModuleScope HelpDeskToolkit { New-HDRichiestaNbstat }
        $q.Length | Should -Be 50
        $q[12] | Should -Be 0x20
        [System.Text.Encoding]::ASCII.GetString($q, 13, 32) | Should -Be ('CK' + ('A' * 30))
        $q[47] | Should -Be 0x21
        $q[49] | Should -Be 0x01
    }

    It 'legge nome, gruppo e MAC dalla risposta (<Caso>)' -TestCases @(
        @{ Caso = 'nome intero'; Compresso = $false }
        @{ Caso = 'nome compresso'; Compresso = $true }
    ) {
        $r = New-Object System.Collections.Generic.List[byte]
        $r.AddRange([byte[]]@(0x48, 0x44, 0x84, 0x00, 0, 0, 0, 1, 0, 0, 0, 0))
        if ($Compresso) { $r.AddRange([byte[]]@(0xC0, 0x0C)) }
        else {
            $q = InModuleScope HelpDeskToolkit { New-HDRichiestaNbstat }
            $r.AddRange([byte[]]$q[12..45])
        }
        $r.AddRange([byte[]]@(0, 0x21, 0, 1, 0, 0, 0, 0, 0, 0))
        $r.Add(3)
        foreach ($voce in @(@('PC-RECEPTION', 0x00, 0x04), @('STUDIO', 0x00, 0x84), @('PC-RECEPTION', 0x20, 0x04))) {
            $r.AddRange([byte[]][System.Text.Encoding]::ASCII.GetBytes(([string]$voce[0]).PadRight(15)))
            $r.AddRange([byte[]]@($voce[1], $voce[2], 0))
        }
        $r.AddRange([byte[]]@(0x00, 0x11, 0x22, 0x33, 0x44, 0x55))
        $r.AddRange([byte[]](@(0) * 40))
        $p = InModuleScope HelpDeskToolkit -Parameters @{ B = $r.ToArray() } { param($B) ConvertFrom-HDRispostaNbstat -Byte $B }
        $p.Nome | Should -Be 'PC-RECEPTION'
        $p.Gruppo | Should -Be 'STUDIO'
        $p.MAC | Should -Be '00:11:22:33:44:55'
    }

    It 'una risposta troncata non dà errori' {
        InModuleScope HelpDeskToolkit { ConvertFrom-HDRispostaNbstat -Byte (New-Object byte[] 20) } | Should -BeNullOrEmpty
    }
}

Describe 'Produttore dal MAC e nomi mDNS' {
    It 'controllo delle porte: rifiuto valido solo sulla rete del PC, scartato se risponde "tutto"' {
        InModuleScope HelpDeskToolkit {
            $locale = Get-HDInfoRete '192.168.1.0/24'
            $sonde = @(
                [pscustomobject]@{ IP = '192.168.1.50'; Aperta = $false }
                [pscustomobject]@{ IP = '10.0.5.7'; Aperta = $false }
                [pscustomobject]@{ IP = '10.0.5.8'; Aperta = $true }
            )
            @(Select-HDSondeAffidabili -Sonde $sonde -InfoLocale $locale -Provati 250) | Should -Be @('192.168.1.50', '10.0.5.8')
            $tutti = @(1..200 | ForEach-Object { [pscustomobject]@{ IP = "10.0.5.$_"; Aperta = $false } })
            @(Select-HDSondeAffidabili -Sonde $tutti -InfoLocale $locale -Provati 250 -WarningAction SilentlyContinue) | Should -BeNullOrEmpty
            $tuttiLocali = @(1..200 | ForEach-Object { [pscustomobject]@{ IP = "192.168.1.$_"; Aperta = ($_ -eq 7) } })
            @(Select-HDSondeAffidabili -Sonde $tuttiLocali -InfoLocale $locale -Provati 250 -WarningAction SilentlyContinue) | Should -Be @('192.168.1.7')
        }
    }

    It 'scarta i nomi segnaposto come "(none)"' {
        InModuleScope HelpDeskToolkit {
            Test-HDNomeValido '(none)-2' | Should -BeFalse
            Test-HDNomeValido 'localhost' | Should -BeFalse
            Test-HDNomeValido '' | Should -BeFalse
            Test-HDNomeValido 'GT-AX11000_Pro' | Should -BeTrue
        }
    }

    It 'riconosce il produttore dall''elenco IEEE e i MAC casuali dei telefoni' {
        InModuleScope HelpDeskToolkit {
            Get-HDProduttoreMac 'EC:71:DB:AE:F6:2F' | Should -Be 'Reolink Innovation Limited'
            Get-HDProduttoreMac '94-08-53-D6-B9-9A' | Should -Be 'Liteon Technology Corporation'
            Get-HDProduttoreMac '50284a9e5377' | Should -Be 'Intel Corporate'
            Get-HDProduttoreMac '52:A2:0A:B3:51:B0' | Should -BeLike 'MAC casuale*'
            # blocchi piccoli dell'IEEE Registration Authority (MA-M e MA-S); blocco non assegnato: produttore ignoto
            Get-HDProduttoreMac 'E8:FF:1E:C0:20:B3' | Should -Be 'Fracarro Radioindustrie Srl'
            Get-HDProduttoreMac '00:1B:C5:00:10:22' | Should -Be 'OpenRB.com, Direct SIA'
            Get-HDProduttoreMac 'E8:FF:1E:F0:00:01' | Should -Be ''
            Get-HDProduttoreMac '' | Should -Be ''
        }
    }

    It 'prepara la domanda mDNS (PTR di in-addr.arpa) e legge il nome dalla risposta, anche compressa' {
        InModuleScope HelpDeskToolkit {
            $q = New-HDPacchettoMdnsInverso -Ip '192.168.3.1' -Id 0x1234
            $q[0..5] | Should -Be @(0x12, 0x34, 0, 0, 0, 1)
            [System.Text.Encoding]::ASCII.GetString($q, 12, $q.Length - 12) | Should -BeLike '*1*3*168*192*in-addr*arpa*'
            $q[-4..-1] | Should -Be @(0, 12, 0, 1)
            # risposta: intestazione, la stessa domanda, poi un PTR con il nome e "local" compresso
            $nomeDomanda = @($q[12..($q.Length - 5)])
            $risposta = [System.Collections.Generic.List[byte]]::new()
            foreach ($b in @(0x12, 0x34, 0x84, 0, 0, 1, 0, 1, 0, 0, 0, 0)) { $risposta.Add([byte]$b) }
            foreach ($b in $nomeDomanda) { $risposta.Add([byte]$b) }
            foreach ($b in @(0, 12, 0, 1)) { $risposta.Add([byte]$b) }
            foreach ($b in @(0xC0, 12, 0, 12, 0, 1, 0, 0, 0, 120)) { $risposta.Add([byte]$b) }
            $dati = [System.Collections.Generic.List[byte]]::new()
            $dati.Add(14); foreach ($c in 'GT-AX11000_Pro'.ToCharArray()) { $dati.Add([byte][int]$c) }
            $dati.Add(5); foreach ($c in 'local'.ToCharArray()) { $dati.Add([byte][int]$c) }
            $dati.Add(0)
            $risposta.Add(0); $risposta.Add([byte]$dati.Count)
            foreach ($b in $dati) { $risposta.Add($b) }
            (Get-HDNomiDaMulticast -Byte $risposta.ToArray() -Da '192.168.3.1')['192.168.3.1'] | Should -Be 'GT-AX11000_Pro'
            (Get-HDNomiDaMulticast -Byte ([byte[]]@(1, 2, 3)) -Da '192.168.3.1').Count | Should -Be 0
        }
    }

    It 'legge i nomi dai servizi annunciati (record A e SRV) e dalle domande inverse al gruppo multicast' {
        InModuleScope HelpDeskToolkit {
            $scriviNome = {
                param($Lista, [string]$Testo)
                foreach ($e in $Testo.Split('.')) { $t = [System.Text.Encoding]::UTF8.GetBytes($e); $Lista.Add([byte]$t.Length); $Lista.AddRange($t) }
                $Lista.Add([byte]0)
            }
            $scriviRecord = {
                param($Lista, [string]$Testo, [int]$Tipo, [object[]]$Prima, [string]$Bersaglio, [byte[]]$Indirizzo)
                & $scriviNome $Lista $Testo
                $dati = [System.Collections.Generic.List[byte]]::new()
                foreach ($b in $Prima) { $dati.Add([byte]$b) }
                if ($Bersaglio) { & $scriviNome $dati $Bersaglio }
                if ($Indirizzo) { $dati.AddRange($Indirizzo) }
                foreach ($b in @(0, $Tipo, 0x80, 1, 0, 0, 0, 120, 0, $dati.Count)) { $Lista.Add([byte]$b) }
                $Lista.AddRange($dati)
            }
            $intestazione = { param($Lista, [int]$Risposte, [int]$Aggiuntivi) foreach ($b in @(0x48, 0x45, 0x84, 0, 0, 0, 0, $Risposte, 0, 0, 0, $Aggiuntivi)) { $Lista.Add([byte]$b) } }

            # un Mac risponde alla domanda sul servizio "companion-link": PTR, SRV con il nome host e il suo IP
            $mac = [System.Collections.Generic.List[byte]]::new()
            & $intestazione $mac 1 2
            & $scriviRecord $mac '_companion-link._tcp.local' 12 $null 'Air di Bianca._companion-link._tcp.local' $null
            & $scriviRecord $mac 'Air di Bianca._companion-link._tcp.local' 33 @(0, 0, 0, 0, 0xC0, 0x3D) 'Air-di-Bianca.local' $null
            & $scriviRecord $mac 'Air-di-Bianca.local' 1 $null '' ([byte[]]@(192, 168, 3, 184))
            $nomi = Get-HDNomiDaMulticast -Byte $mac.ToArray() -Da '192.168.3.184'
            $nomi.Count | Should -Be 1
            $nomi['192.168.3.184'] | Should -Be 'Air-di-Bianca'
            # la stessa risposta data per conto del Mac da un altro dispositivo: il nome resta al Mac
            @((Get-HDNomiDaMulticast -Byte $mac.ToArray() -Da '192.168.3.50').Keys) | Should -Be @('192.168.3.184')

            # solo SRV, senza indirizzo: il nome vale per chi ha risposto
            $stampante = [System.Collections.Generic.List[byte]]::new()
            & $intestazione $stampante 1 0
            & $scriviRecord $stampante 'Ufficio._ipp._tcp.local' 33 @(0, 0, 0, 0, 2, 0x77) 'stampante-ufficio.local' $null
            (Get-HDNomiDaMulticast -Byte $stampante.ToArray() -Da '192.168.3.60')['192.168.3.60'] | Should -Be 'stampante-ufficio'

            # risposta inversa (domanda al gruppo per più indirizzi): l'IP viene dal nome in-addr.arpa
            $q = New-HDPacchettoMdnsInverso -Ip '192.168.3.184', '192.168.3.36' -Unicast
            $q[4..5] | Should -Be @(0, 2)
            $q[-2..-1] | Should -Be @(0x80, 1)
            $inversa = [System.Collections.Generic.List[byte]]::new()
            & $intestazione $inversa 1 0
            & $scriviRecord $inversa '184.3.168.192.in-addr.arpa' 12 $null 'Air-di-Bianca.local' $null
            $nomi = Get-HDNomiDaMulticast -Byte $inversa.ToArray() -Da '192.168.3.184'
            $nomi.Count | Should -Be 1
            $nomi['192.168.3.184'] | Should -Be 'Air-di-Bianca'

            # domande al gruppo: una per i servizi, poi le inverse a gruppi di 30
            $pacchetti = @(Get-HDDomandeMulticast -Indirizzi @(1..70 | ForEach-Object { "192.168.3.$_" }))
            $pacchetti.Count | Should -Be 4
            $pacchetti[0][5] | Should -Be $script:HDServiziMdns.Count
            $pacchetti[1][5] | Should -Be 30
            $pacchetti[3][5] | Should -Be 10
            $pacchetti[3].Length | Should -BeLessThan 1400
        }
    }
}

Describe 'Invoke-ScansioneRete (rete simulata)' {
    BeforeEach {
        Mock -ModuleName HelpDeskToolkit Test-HDWindows { $true }
        Mock -ModuleName HelpDeskToolkit Get-HDReteLocale { [pscustomobject]@{ IP = '192.168.1.37'; MAC = 'AA:AA:AA:AA:AA:37'; Rete = '192.168.1.0/24'; Nota = ''; Hostname = 'PC-TECNICO' } }
        # .69 è un telefono lento: risponde solo al secondo giro, con un timeout più lungo
        Mock -ModuleName HelpDeskToolkit Invoke-HDPingParallelo {
            $vivi = @('192.168.1.1', '192.168.1.20', '192.168.1.37')
            if ($TimeoutMs -ge 1500) { $vivi += '192.168.1.69' }
            @($vivi | Where-Object { $Indirizzi -contains $_ })
        }
        Mock -ModuleName HelpDeskToolkit Get-HDViciniArp {
            [pscustomobject]@{ IP = '192.168.1.1'; MAC = 'CC:28:AA:D2:EE:F0' }
            [pscustomobject]@{ IP = '192.168.1.20'; MAC = '00:11:22:33:44:55' }
            [pscustomobject]@{ IP = '192.168.1.30'; MAC = '00:11:22:33:44:30' }
        }
        Mock -ModuleName HelpDeskToolkit Resolve-HDNetbiosRemoto {
            @{
                '192.168.1.20' = [pscustomobject]@{ Nome = 'PC-RECEPTION'; Gruppo = 'STUDIO'; MAC = '00:11:22:33:44:55' }
                '192.168.1.30' = [pscustomobject]@{ Nome = 'PC-MAGAZZINO'; Gruppo = 'STUDIO'; MAC = '' }
            }
        }
        Mock -ModuleName HelpDeskToolkit Resolve-HDNomeDnsInverso { @{ '192.168.1.1' = 'router.fritz.box' } }
        # .184 ignora ping e ARP ma ha una porta aperta; .200 risponde solo a mDNS (un Mac)
        Mock -ModuleName HelpDeskToolkit Invoke-HDSondaTcp { [pscustomobject]@{ IP = '192.168.1.184'; Aperta = $false } }
        Mock -ModuleName HelpDeskToolkit Resolve-HDNomeMulticast { @{ '192.168.1.69' = 'Redmi-Note-12'; '192.168.1.200' = 'Air-di-Bianca'; '10.9.9.9' = 'fuori-rete' } }
    }

    It 'unisce ping, ARP e nomi, in ordine di IP, e salva una copia CSV' {
        $s = Invoke-ScansioneRete -Confirm:$false
        $s.Rete | Should -Be '192.168.1.0/24'
        @($s.Dispositivi | ForEach-Object { $_.IP }) | Should -Be @('192.168.1.1', '192.168.1.20', '192.168.1.30', '192.168.1.37', '192.168.1.69', '192.168.1.184', '192.168.1.200')
        $router = $s.Dispositivi[0]
        $router.Hostname | Should -Be 'ROUTER'
        $router.MAC | Should -Be 'CC:28:AA:D2:EE:F0'
        $router.Produttore | Should -Be 'ASUSTek COMPUTER INC.'
        $router.Fonte | Should -Be 'ping, ARP, DNS'
        $s.Dispositivi[2].Fonte | Should -Be 'ARP, NetBIOS'
        $s.Dispositivi[3].Hostname | Should -Be 'PC-TECNICO'
        $s.Dispositivi[3].MAC | Should -Be 'AA:AA:AA:AA:AA:37'
        Test-Path -LiteralPath $s.FileCsv | Should -BeTrue
        $s.Dispositivi[4].Hostname | Should -Be 'REDMI-NOTE-12'
        $s.Dispositivi[4].Fonte | Should -Be 'ping, mDNS'
        $s.Dispositivi[5].Fonte | Should -Be 'TCP'
        $s.Dispositivi[6].Hostname | Should -Be 'AIR-DI-BIANCA'
        $s.Dispositivi[6].Fonte | Should -Be 'mDNS'
        Should -Invoke -ModuleName HelpDeskToolkit Invoke-HDPingParallelo -Times 2
        Should -Invoke -ModuleName HelpDeskToolkit Invoke-HDSondaTcp -Times 1 -ParameterFilter { @($Indirizzi).Count -eq 249 }
        Should -Invoke -ModuleName HelpDeskToolkit Resolve-HDNetbiosRemoto -Times 1 -ParameterFilter { @($Indirizzi).Count -eq 254 }
        Should -Invoke -ModuleName HelpDeskToolkit Resolve-HDNomeMulticast -Times 1 -ParameterFilter { @($Indirizzi).Count -eq 254 -and $IpLocale -eq '192.168.1.37' -and @($Locali).Count -eq 254 }
        Should -Invoke -ModuleName HelpDeskToolkit Resolve-HDNomeDnsInverso -Times 1 -ParameterFilter { @($Indirizzi).Count -eq 5 }
    }

    It 'con -NelCensimento mette i dispositivi nell''Excel (senza scrivere il dominio)' {
        $s = Invoke-ScansioneRete -NelCensimento -Cliente 'Rete Simulata' -NonChiedere -Confirm:$false
        $s.Import.Aggiunti | Should -Be 7
        $file = Join-Path $script:CartellaCensimentoTest 'Censimento_Rete-Simulata.xlsx'
        $t = InModuleScope HelpDeskToolkit -Parameters @{ P = $file } { param($P) Get-HDTabellaCensimento $P }
        $t.Righe | Should -HaveCount 7
        $pc = $t.Righe | Where-Object { $_.Valori['Hostname'] -eq 'PC-RECEPTION' }
        $pc.Valori['MAC'] | Should -Be '00:11:22:33:44:55'
        $pc.Valori.ContainsKey('Dominio') | Should -BeFalse
        $s2 = Invoke-ScansioneRete -NelCensimento -Censimento $file -NonChiedere -Confirm:$false
        $s2.Import.Aggiunti | Should -Be 0
    }

    It 'fuori da Windows si ferma con un messaggio chiaro' {
        Mock -ModuleName HelpDeskToolkit Test-HDWindows { $false }
        { Invoke-ScansioneRete -Confirm:$false } | Should -Throw '*solo su Windows*'
    }
}
