# get_version.ps1
# Liest aktuelle Version aus AssemblyInfo.cs, zeigt InputBox-Dialog,
# validiert Format und gibt das Ergebnis auf stdout aus.
# Bei Abbruch oder ungueltigem Format: leere Ausgabe + Exit-Code 1.

$ErrorActionPreference = 'Stop'

$assemblyInfo = 'C:\Users\CC-Student\source\repos\WindowsFormsApp1\WindowsFormsApp1\Properties\AssemblyInfo.cs'

# Aktuelle Version aus AssemblyInfo.cs auslesen
$current = '1.0.0.0'
try {
    $match = Select-String -Path $assemblyInfo -Pattern 'AssemblyVersion\("([^"]+)"\)' | Select-Object -First 1
    if ($match) {
        $current = $match.Matches.Groups[1].Value
    }
} catch {
    # Bei Fehler: Default behalten
}

Add-Type -AssemblyName Microsoft.VisualBasic

$pattern = '^[0-9]+\.[0-9]+\.[0-9]+\.[0-9]+$'

while ($true) {
    $result = [Microsoft.VisualBasic.Interaction]::InputBox(
        "Version eingeben (Format: x.y.z.w)`n`nAktuelle Version: $current",
        'Inventarisierung Build',
        $current
    )

    # Abbrechen oder leer -> leere Ausgabe + Fehler-Exit
    if ([string]::IsNullOrWhiteSpace($result)) {
        exit 1
    }

    $result = $result.Trim()

    if ($result -match $pattern) {
        # Gueltig: Version ausgeben und beenden
        Write-Output $result
        exit 0
    }

    # Ungueltig: Hinweis und erneut fragen
    [Microsoft.VisualBasic.Interaction]::MsgBox(
        "Ungueltiges Format: '$result'`n`nErwartet: Major.Minor.Build.Revision (z.B. 1.2.3.4)",
        'OKOnly,Critical',
        'Inventarisierung Build'
    ) | Out-Null
}
