# Funktion, um die Daten in die CSV-Datei zu schreiben
function Write-DataToCSV {
    param(
        [string]$csvFile,
        [array]$data
    )
    $data | Export-Csv -Path $csvFile -NoTypeInformation
}

# Funktion, um die Zeit in das Format HH:MM umzurechnen
function Format-Time {
    param(
        [int]$totalMinutes
    )

    try {
        $hours = [math]::Floor($totalMinutes / 60)
        $minutes = $totalMinutes % 60
        "{0}:{1:D2}" -f $hours, $minutes
    }
    catch {
        "99:99"
    }
}

# Funktion, um zu überprüfen, ob das Datum ein Feiertag oder ein Wochenende ist
function IsHolidayOrWeekend {
    param(
        [DateTime]$date
    )

    $weekDay = $date.DayOfWeek
    $holidays = @(
        "01.01.2021", # Neujahr
        "06.01.2021", # Heilige Drei Könige
        "02.04.2021", # Karfreitag
        "05.04.2021", # Ostermontag
        "01.05.2021", # Tag der Arbeit
        "13.05.2021", # Christi Himmelfahrt
        "24.05.2021", # Pfingstmontag
        "03.06.2021", # Fronleichnam
        "15.08.2021", # Mariä Himmelfahrt
        "03.10.2021", # Tag der Deutschen Einheit
        "01.11.2021", # Allerheiligen
        "25.12.2021", # 1. Weihnachtstag
        "26.12.2021", # 2. Weihnachtstag
        "01.01.2022",
        "06.01.2022",
        "15.04.2022",
        "18.04.2022",
        "01.05.2022",
        "26.05.2022",
        "06.06.2022",
        "15.08.2022",
        "03.10.2022",
        "01.11.2022",
        "25.12.2022",
        "26.12.2022",
        "01.01.2023",
        "06.01.2023",
        "07.04.2023",
        "10.04.2023",
        "01.05.2023",
        "25.05.2023",
        "06.06.2023",
        "15.08.2023",
        "03.10.2023",
        "01.11.2023",
        "25.12.2023",
        "26.12.2023"
    )
    # Überprüfen, ob das Datum ein Wochenende ist (Samstag oder Sonntag)
    if ($weekDay -eq "Saturday" -or $weekDay -eq "Sunday") {
        return $true
    }
    
    # Hier sollte die Überprüfung für Feiertage in Bayern (katholische Region) erfolgen
    # Fügen Sie den entsprechenden Code für die Feiertagsüberprüfung hinzu, z. B. anhand einer externen Datenquelle

    $isHoliday = $holidays -contains $date.ToShortDateString()

    return $isHoliday
}

# Ereignislog "System" durchsuchen und Zeiten erfassen
$events = Get-WinEvent -FilterHashtable @{
    LogName = 'System'
    ID      = 12, 13
} | Sort-Object TimeCreated

# Variablen zur Aufzeichnung von Arbeitszeiten pro Tag initialisieren
$startTimeOfDay = $null
$endTimeOfDay = $null
$workTimeData = @()

$currentCalendarWeek = -1  # Variable, um die aktuelle Kalenderwoche zu speichern
$weekSum = 0  # Variable, um die wöchentliche Arbeitszeitsumme zu speichern
$currentMonth = -1  # Variable, um den aktuellen Monat zu speichern
$monthSum = 0  # Variable, um die monatliche Arbeitszeitsumme zu speichern

foreach ($event in $events) {
    $eventTime = $event.TimeCreated
    $calendar = [System.Globalization.DateTimeFormatInfo]::CurrentInfo.Calendar
    $calendarWeek = $calendar.GetWeekOfYear($eventTime, [System.Globalization.CalendarWeekRule]::FirstFourDayWeek, [System.DayOfWeek]::Monday)
    $month = $eventTime.Month

    # Überprüfen, ob das Datum ein Feiertag oder ein Wochenende ist
    $isHolidayOrWeekend = IsHolidayOrWeekend $eventTime

    # Wenn sich der Wert der Kalenderwoche ändert, füge eine Zeile mit "Wochensumme" hinzu
    if ($calendarWeek -ne $currentCalendarWeek -and $currentCalendarWeek -ne -1) {
        $workTimeData += [PSCustomObject]@{
            "Datum"                      = "Wochensumme"
            "Wochentag"                  = $null
            "Kalenderwoche"              = $currentCalendarWeek
            "Einschaltzeitpunkt"         = $null
            "Ausschaltzeitpunkt"         = $null
            "Zeitdifferenz (Minuten)"    = $null
            "Pause (Minuten)"            = $null
            "Nettoarbeitszeit (Minuten)" = $weekSum
            "Netto Arbeitszeit konv"     = Format-Time $weekSum
            "Feiertag"                   = $null
        }
        $weekSum = 0
    }

    # Wenn sich der Monat im Datum ändert, füge eine Zeile mit "Monatssumme" hinzu
    if ($month -ne $currentMonth -and $currentMonth -ne -1) {
        $workTimeData += [PSCustomObject]@{
            "Datum"                      = "Monatssumme"
            "Wochentag"                  = $null
            "Kalenderwoche"              = $null
            "Einschaltzeitpunkt"         = $null
            "Ausschaltzeitpunkt"         = $null
            "Zeitdifferenz (Minuten)"    = $null
            "Pause (Minuten)"            = $null
            "Nettoarbeitszeit (Minuten)" = $monthSum
            "Netto Arbeitszeit konv"     = Format-Time $monthSum
            "Feiertag"                   = $null
        }
        $monthSum = 0
    }

    # Wenn es ein Startereignis (ID 12) ist
    if ($event.Id -eq 12) {
        $startTimeOfDay = $eventTime
    }
    # Wenn es ein Stoppereignis (ID 13) ist
    elseif ($event.Id -eq 13 -and $startTimeOfDay -ne $null) {
        $endTimeOfDay = $eventTime

        # Berechne die Zeitdifferenz nur, wenn sowohl Start- als auch Stopp-Ereignisse vorhanden sind
        $timeDifference = $endTimeOfDay - $startTimeOfDay

        if ($timeDifference -ne $null) {
            $timeDifferenceMinutes = [math]::Floor($timeDifference.TotalMinutes)
            $pauseTimeInMinutes = 45
            $nettoArbeitszeitMinutes = $timeDifferenceMinutes - $pauseTimeInMinutes

            $weekSum += $nettoArbeitszeitMinutes
            $monthSum += $nettoArbeitszeitMinutes

            $workTimeData += [PSCustomObject]@{
                "Datum"                      = $eventTime.ToShortDateString()
                "Wochentag"                  = $eventTime.DayOfWeek
                "Kalenderwoche"              = $calendarWeek
                "Einschaltzeitpunkt"         = $startTimeOfDay.ToShortTimeString()
                "Ausschaltzeitpunkt"         = $endTimeOfDay.ToShortTimeString()
                "Zeitdifferenz (Minuten)"    = $timeDifferenceMinutes
                "Pause (Minuten)"            = $pauseTimeInMinutes
                "Nettoarbeitszeit (Minuten)" = $nettoArbeitszeitMinutes
                "Netto Arbeitszeit konv"     = Format-Time $nettoArbeitszeitMinutes
                "Feiertag"                   = if ($isHolidayOrWeekend) { "Feiertag" } else { "Wochentag" }
            }
        }

        $startTimeOfDay = $null
        $endTimeOfDay = $null
    }

    $currentCalendarWeek = $calendarWeek
    $currentMonth = $month
}

# Füge die letzte Wochensumme hinzu, falls vorhanden
if ($weekSum -ne 0) {
    $workTimeData += [PSCustomObject]@{
        "Datum"                      = "Wochensumme"
        "Wochentag"                  = $null
        "Kalenderwoche"              = $currentCalendarWeek
        "Einschaltzeitpunkt"         = $null
        "Ausschaltzeitpunkt"         = $null
        "Zeitdifferenz (Minuten)"    = $null
        "Pause (Minuten)"            = $null
        "Nettoarbeitszeit (Minuten)" = $weekSum
        "Netto Arbeitszeit konv"     = Format-Time $weekSum
        "Feiertag"                   = $null
    }
}

# Füge die letzte Monatssumme hinzu, falls vorhanden
if ($monthSum -ne 0) {
    $workTimeData += [PSCustomObject]@{
        "Datum"                      = "Monatssumme"
        "Wochentag"                  = $null
        "Kalenderwoche"              = $null
        "Einschaltzeitpunkt"         = $null
        "Ausschaltzeitpunkt"         = $null
        "Zeitdifferenz (Minuten)"    = $null
        "Pause (Minuten)"            = $null
        "Nettoarbeitszeit (Minuten)" = $monthSum
        "Netto Arbeitszeit konv"     = Format-Time $monthSum
        "Feiertag"                   = $null
    }
}

# Daten in die CSV-Datei schreiben
Write-DataToCSV "C:\temp\Development\Worktime\worktime_daily.csv" $workTimeData

Write-Host "Das Skript wurde erfolgreich ausgeführt. Die täglichen Einschalt- und Ausschaltvorgänge wurden in die CSV-Datei geschrieben."
