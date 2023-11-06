# PowerPoint Markennamen Korrektor

Dieses Tool analysiert PowerPoint-Dateien (.pptx), identifiziert falsch geschriebene Markennamen und bietet die korrigierten Versionen an.

## Beschreibung

Der PowerPoint Markennamen Korrektor wurde entwickelt, um die Effizienz bei der Überprüfung von Präsentationen zu steigern. Es scannt Texte auf Folien nach einer Liste bekannter falsch geschriebener Markennamen und korrigiert diese automatisch.

## Funktionsweise

Das Tool liest PowerPoint-Dateien ein und sucht in den Textfeldern der Folien nach Markennamen, die in der `BRAND_NAME_MAPPINGS`-Konstante definiert sind. Findet es eine Übereinstimmung, wird der falsch geschriebene Name durch die korrekte Schreibweise ersetzt. Die Ergebnisse werden in einer Textdatei zusammengefasst und können heruntergeladen werden.

## Installation

Klonen Sie das Repository und installieren Sie die erforderlichen Abhängigkeiten:

