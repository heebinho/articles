Title: Unzip in Powershell
Published: 1/24/2021
Tags: [A, B]
Lead: A tale of power in a shell
---

# Intro
Obwohl die PowerShell ein mächtiges Instrument ist, verwende ich sie selten.
Durchdachte Workflows und automatisierte Pipelines erübrigen oft mühselige Scripting Tasks und wenn es doch mal etwas zu automatisieren gibt, verwende ich oftmals LinqPad.
LinqPad ist ein extrem hilfreiches Werkzeug, eine Art REPL mit einer mächtigen .Dump() Methode die Berge versetzen kann - aber das ist eine andere Story.

Anfangs Jahr wechselte ich in ein DevOps-Team. Neben den eigens entwickelten Lösungen wird auch eine von einem externen Dienstleister entwickelte Software betrieben.
Die vom Dienstleister gelieferten Pakete werden durch PowerShell Scripts ergänzt und installiert.
Ein Schritt in diesem Prozess ist das Extrahieren eines Zip-Archivs.

Im Jahr 2021 angekommen, dürfte man meinen, dass Zeichensatz-Probleme der Vergangenheit angehören...
Aber Microsoft belehrt uns eines besseren:

Zutaten:
1. Man erstelle ein File und füge ein diakritisches Zeichen in den Namen e.g. hänsel.txt
2. Verwende den Rechtsklick und die Option "Send to Compressed (zipped) folder"
3. Man öffne die PowerShell und navigiere in den gewählten Ordner und gebe den folgenden Befehl zum entpacken des Archivs an:


```powershell
Expand-Archive -Path .\hänsel.zip -DestinationPath .\hänsel
```

Angerichtet. Unglaublich aber Microsoft schafft es tatsächlich, dass die eigenen Tools unterschiedliche Zeichensätze verwenden und inkompatibel sind.


## Quick Fix ... but

Mit dem Wissen dass die PowerShell in C# entwickelt wurde und auf dem .Net Framework beruht, war ich der prädestinierte Kandidat das Problem kurz anzuschauen.
C# und das .Net Framework begleiten mich doch schon seit einigen Jahren, mal mehr und mal weniger intensiv.
Im Hinterkopf hatte ich bereits zwei Ideen um das Problem schnellstmöglich aus der Welt zu schaffen:
1. Die Expand-Archive Funktion müsste doch sicher einen Parameter haben um das Encoding anzugeben
2. Alternativ gibt es im .Net Framework den Namespace System.IO.Compression und darin eine ZipFile Klasse

Fehlanzeige beim Parameter aber die zweite Option müsste doch hinhauen.
Noch kurz den vermuteten Zeichensatz recherchieren und ready test:

```powershell
$enc = [System.Text.Encoding]::GetEncoding(437) #IBM Codepage 437
[System.IO.Compression.ZipFile]::ExtractToDirectory($zip, $target, $enc)
```

Funktioniert, Problem gelöst. Nur noch schnell die vorhandenen Scripts anpassen...
Doch beim zweiten Aufruf des Expand-Archive Commandlets ist schnell klar, dass da noch mehr zu tun ist.

```powershell
Get-ChildItem -Path ".\*.zip"  | Expand-Archive -DestinationPath ".\target" -F
```

## Pipeline Input

Pipelines, ein Konzept geschmidet in den dunklen Kellern der Bell Labs lang vor meiner Geburt. Eine geniale Idee die bis heute seine Kreise zieht und durch das vermehrte aufkommen von fp Paradigmen gefühlt ein Revival feiert.

Die Funktion Get-ChidItem leitet oder eben "piped" eine Sequenz gefundener Zip-Archive weiter an die Expand-Archive Funktion. Oder einfach gesagt, das Ergebnis der ersten Funktion dient als Eingabe für die zweite Funktion.

Eine Funktion zu schreiben, die einen Pipeline Input Parameter hat, kann nicht so schwierig sein:


```powershell

function Unzip {
    
    param(  [Parameter(ValueFromPipeline=$true)][string]$zip,
            [Parameter()][string]$target,
            [Parameter()][switch]$force,
            [Parameter()][System.Text.Encoding]$encoding=[System.Text.Encoding]::GetEncoding(437),
            [Parameter()][switch]$v)

    begin  { } #setup
    process{ } #process item
    end    { } #teardown
}
```
Das Gerüst steht. Der schwierigste Part dabei war, herauszufinden wie die Funktion korrekt aufgerufen wird:

```powershell
Get-ChildItem -Path ".\*.zip"  | Unzip -target "C:\Users\renato\code\gist\unzipped" -f -encoding $encoding -v

```


## .Net Jungle

Bleibt noch die ExtractToDirectory-Methode mit den vier Parameter aufzurufen:

```csharp
public static void ExtractToDirectory (string sourceArchiveFileName, string destinationDirectoryName, System.Text.Encoding? entryNameEncoding, bool overwriteFiles);
```

Doch die nächste Überraschung folgt sogleich, die Überladung existiert erst seit .Net Core 2.0 und da die PowerShell in der bei uns eingesetzten Variante noch auf dem klassischen .Net Framework basiert, ist nun auch der zweite Lösungsansatz dahin.

Nicht so schlimm, mit der ZipFile-Klasse kann man das Archiv immer noch lesen und da gibt es ja noch die Extension-Method ExtractToFile:


```csharp
public static void ExtractToFile (this System.IO.Compression.ZipArchiveEntry source, string destinationFileName, bool overwrite);
```

Leider gibts aber auch diese Methode erst in späteren Versionen.
Sich über Microsoft zu ärgern nützt nichts, es muss eine eigene Implementierung her.


## Implementierung in C#

Da die Implementierung in C# für mich persönlich einfacher ist, starte ich dort.
Der System.IO Namespace bietet ein Ensemble an Klassen und Funktionen, die das "Heavy Lifting" übernehmen.
Da Microsoft das .Net Framework unter einer Open Source Lizenz veröffentlicht hat und die ganzen Sourcen auf GitHub liegen, muss das Rad nicht neu erfunden werden.

```csharp
$source = @"
namespace System.IO.Compression
{
    public static class ZipArchiveEntryExtensions
    {
        public static void ExtractToFile(ZipArchiveEntry source, string destinationFileName, bool overwrite)
        {
            FileMode fMode = overwrite ? FileMode.Create : FileMode.CreateNew;

            using (Stream fs = new FileStream(destinationFileName, fMode, FileAccess.Write, FileShare.None, bufferSize: 0x1000, useAsync: false))
            {
                using (Stream es = source.Open())
                    es.CopyTo(fs);
            }

            File.SetLastWriteTime(destinationFileName, source.LastWriteTime.DateTime);
        }
    }
}
"@

Add-Type -TypeDefinition $source

[System.IO.Compression.ZipArchiveEntryExtensions]::ExtractToFile($entry, $fullname, $force)

```
Der spannende Aspekt dieses Code Blocks ist, dass zur Laufzeit eine Assembly (DLL) generiert und geladen wird.
Danach kann wie gewohnt auf die entsprechenden Funktionen zugegriffen werden.


## Back to PowerShell

Um die Scipts der Kollegen nicht mit C# zu "belasten", habe ich den Code zum extrahieren noch in eine PowerShell-Funktion übersetzt.
Es folgt die ganze Implementierung mit Beispiel-Aufrufen:


```powershell
Set-StrictMode -Version Latest
Add-Type -AssemblyName System.IO.Compression.FileSystem
function Unzip {
    
    param(  [Parameter(ValueFromPipeline=$true)][string]$zip,
            [Parameter()][string]$target,
            [Parameter()][switch]$force,
            #ZIP uses a default codepage of IBM437.
            [Parameter()][System.Text.Encoding]$encoding=[System.Text.Encoding]::GetEncoding(437),
            [Parameter()][switch]$v)

    begin{
        [System.Diagnostics.Stopwatch] $stopwatch = [System.Diagnostics.Stopwatch]::new()
        $stopwatch.Start()
        if($v){
            Write-Output "verbose -> $v"
            Write-Output "target -> $target"
            Write-Output "force -> $force"
        }

        if($force -and [System.IO.Directory]::Exists($target)){
            Remove-Item -LiteralPath $target -Force -Recurse
            if($v){ Write-Host "removed directory: $target" }
        }
        if (![System.IO.Directory]::Exists($target)){
            $dir = [System.IO.Directory]::CreateDirectory($target)
            if($v){ Write-Host "created target directory: $dir `n" }
        }
    }

    process{
        if($v) {Write-Output "zip -> $zip" }
        $zip = Resolve-Path $zip
        if($v) {Write-Output "zip fullname -> $zip" }

        $entries = [System.IO.Compression.ZipFile]::Open($zip, [System.IO.Compression.ZipArchiveMode]::Read, $encoding).Entries
        foreach ($entry in $entries){

            $fullname = [System.IO.Path]::GetFullPath( [System.IO.Path]::Combine($target, $entry.FullName) )
            $dir = [System.IO.Path]::GetDirectoryName($fullname)
            if (![System.IO.Directory]::Exists($dir)){
                $dir = [System.IO.Directory]::CreateDirectory($dir)
                if($v) {Write-Host "created directory -> $dir" }
                continue
            }
            if($v){ write-host "extract file -> $fullname" }
            extractToFile $entry $fullname $force
        }
    }
    
    end{
        $stopwatch.Stop()
        $ms = $stopwatch.ElapsedMilliseconds
        Write-Output "unzipping $zip to $target finished in $ms ms"
    }
}

function extractToFile {
    param (
        [Parameter()][System.IO.Compression.ZipArchiveEntry]$source,
        [Parameter()][string]$destinationFileName,
        [Parameter()][bool]$overwrite=$false
    )
    $mode = [System.IO.FileMode]::CreateNew
    if($overwrite){ $mode = [System.IO.FileMode]::Create}

    try {
        [System.IO.Stream] $fs = [System.IO.FileStream]::new($destinationFileName, $mode, [System.IO.FileAccess]::Write, [System.IO.FileShare]::None)
        [System.IO.Stream] $es = $source.Open()
        $es.CopyTo($fs)
    }
    catch {
        Write-Host "error while extracting $source to $destinationFileName overwrite:$overwrite"
        Write-Host $Error[0]
    }
    finally {
        if($es){$es.Dispose()}
        if($fs){$fs.Dispose()}
    }

    [System.IO.File]::SetLastWriteTime($destinationFileName, $source.LastWriteTime.DateTime)
}


# Unzip with a custom name entry encoding
$encoding = [System.Text.Encoding]::GetEncoding(437)

#using piped input
Get-ChildItem -Path ".\*.zip"  | Unzip -target "C:\Users\renato\code\gist\unzipped" -f -encoding $encoding -v
#using absolute input
Unzip -zip "C:\Users\renato\code\gist\a.zip" -target "C:\Users\renato\code\gist\a" -f -encoding $encoding 
#using relative input
Unzip -zip ".\b.zip" -target "C:\Users\renato\code\gist\b" -f -encoding $encoding -v
```


## Performance

Das Zip von unserem Lieferanten ist relativ gross (ca. 110 mb) und das Extrahieren mit der Expand-Archive Funktion dauerte auf meinem Notebook 86733 ms. Das Extrahieren des gleichen Zip mit der neuen Funktion dauerte 12403 ms.
Die neue Funktion ist 7-Mal schneller! Das ist doch eine grosse Überraschung und ein schöner Nebeneffekt.

Meine Vermutung ist, dass das Zeichnen eines Fortschrittbalkens in der Expand-Archive Funktion Zeit kostet.

## Review

Falls jemand den ganzen Post gelesen haben sollte und andere Lösungsansätze oder Verbesserungsvorschläge hat, sind diese natürlich ausserordentlich Willkommen.


## Links

Expand-Archive [MSDN](https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.archive/expand-archive?view=powershell-7.1)

ZipFile.ExtractToDirectory [MSDN](https://docs.microsoft.com/en-us/dotnet/api/system.io.compression.zipfile.extracttodirectory?view=net-5.0)

ZipFileExtensions.ExtractToFile [MSDN](https://docs.microsoft.com/en-us/dotnet/api/system.io.compression.zipfileextensions.extracttofile?view=net-5.0&viewFallbackFrom=netframework-4.0#System_IO_Compression_ZipFileExtensions_ExtractToFile_System_IO_Compression_ZipArchiveEntry_System_String_System_Boolean_)

System.IO.Compression.ZipFileExtenstions [MSDN](https://github.com/dotnet/runtime/blob/061e0aaaf646e5f6a82eef97c6c4210782e5a7c9/src/libraries/System.IO.Compression.ZipFile/src/System/IO/Compression/ZipFileExtensions.ZipArchiveEntry.Extract.cs#L67)

Pipelines [Wikipedia](https://en.wikipedia.org/wiki/Pipeline_(Unix))


Programmer's Playground [LinqPad](https://www.linqpad.net/)

Code Page 437 [Wikipedia](https://en.wikipedia.org/wiki/Code_page_437)

Code page Identifiers [MSDN](https://docs.microsoft.com/en-us/windows/win32/intl/code-page-identifiers)

Diakritische Zeichen [Wikipedia](https://de.wikipedia.org/wiki/Diakritisches_Zeichen)

