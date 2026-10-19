# LAN Scanner [by Mauro Tassera]

Utility Windows per scansionare la LAN e ottenere **IP, MAC, nome host, sistema
operativo (stima) e utente loggato** dei dispositivi, con export CSV/Excel.
.NET 8 · WinForms · tema scuro grigio grafite.

## File del progetto
- `Program.cs` — tutta l'app (interfaccia + logica)
- `LanScanner.csproj` — progetto .NET 8 (dipendenze: ClosedXML, System.Management)
- `antenna.ico` — icona dell'applicazione
- `LanScanner.iss` — script Inno Setup per generare l'installer

## Requisiti
- **.NET 8 SDK** (per compilare)
- **Inno Setup** (solo per creare l'installer): https://jrsoftware.org/isdl.php

## 1) Provare l'app
Da Visual Studio: apri `LanScanner.csproj` e premi F5.
Da terminale, nella cartella del progetto:
```
dotnet run
```

## 2) Creare l'eseguibile autonomo (un solo .exe, senza .NET sul PC di destinazione)
```
dotnet publish -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true -p:IncludeNativeLibrariesForSelfExtract=true -p:EnableCompressionInSingleFile=true
```
Risultato: `bin\Release\net8.0-windows\win-x64\publish\LanScanner.exe`

## 3) Creare l'installer
Apri `LanScanner.iss` con Inno Setup e premi **Compila** (oppure `ISCC.exe LanScanner.iss`).
Risultato: `installer\Setup-LANScanner-1.0.0.exe`

## Note d'uso
- La casella **"Dettagli host: OS e utente"** attiva nome host, stima OS e lettura
  dell'utente loggato. È più lenta: lasciala spenta per una scansione rapida solo IP+MAC.
- La stima **OS / tipo** combina: il **gateway** della rete (→ "Router / Gateway"),
  il **produttore dal MAC** (OUI: Apple, AVM, TP-Link… → colonna "Produttore"),
  il riconoscimento dei **MAC casuali/privacy** (tipici degli smartphone) e infine
  TTL del ping + porte (445/3389 → Windows, 22 → Linux/Unix). È una stima.
- Il DB produttori integrato (nel codice) copre i vendor più comuni. Per coverage
  COMPLETA c'è il file **`oui.txt`** (~40.000 produttori) accanto a `LanScanner.exe`:
  il progetto lo copia lì in automatico a ogni build/publish e l'app lo carica all'avvio.
- La barra di stato all'avvio indica se `oui.txt` è stato caricato e quante voci ha il DB.
- `oui.txt` è generato dal registro pubblico **IEEE MA-L** usando i nomi aziendali
  ufficiali, quindi è ridistribuibile citando la fonte. Per rigenerarlo: scarica `oui.csv`
  da https://standards-oui.ieee.org/ e convertilo in righe "PREFISSO Nome" (6 hex + nome).
  L'app accetta anche il formato manuf di Wireshark.

Fonte dati produttori: IEEE Registration Authority — https://standards-oui.ieee.org/
- L'**utente loggato** si legge via WMI (`Win32_ComputerSystem.UserName`) solo sui
  dispositivi con un nome host, e solo dove hai **diritti amministrativi** sul target.
  Serve che sul target sia consentito **WMI-In** nel Windows Firewall (via GPO);
  altrimenti la colonna mostra "—".
- MAC e utente sono leggibili solo per i dispositivi della **stessa sottorete**.

© 2026 Mauro Tassera
