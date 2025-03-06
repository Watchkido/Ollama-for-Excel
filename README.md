In diesem Beitrag zeige ich dir, wie du Ollama lokal auf deinem Computer betreibst und die KI direkt in Excel mit VBA ansprichst. Diese Methode erlaubt es dir, AI-gestützte Funktionen zu nutzen, ohne auf eine Internetverbindung angewiesen zu sein.

Voraussetzungen

Ollama lokal installiert – Falls noch nicht geschehen, lade Ollama herunter und installiere es. Es sollte als lokaler Server laufen.

Excel mit aktiviertem VBA – Du benötigst eine Excel-Version, die Makros unterstützt (z. B. Excel für Windows).

JSON-Parser für VBA – Falls du JSON-Daten verarbeiten möchtest, lade die JSONConverter.bas von VBA JSON herunter und importiere sie in dein VBA-Projekt.

VBA-Code zur Anbindung von Ollama

Füge obigen VBA-Code in dein Excel-VBA-Modul ein.
  ![ollama artikel vKI](https://github.com/user-attachments/assets/fae99bb3-e3b7-4521-96a1-640c2dcc0023)
       


Anleitung zur Nutzung

VBA-Editor öffnen – Drücke ALT + F11 in Excel. Wenn nötig schalte die Entwicklertools frei!

Neues Modul erstellen – Gehe zu „Einfügen“ > „Modul“.

Code einfügen – Kopiere den obigen VBA-Code in das Modul.

JSON-Parser hinzufügen – Falls noch nicht geschehen, importiere die JsonConverter.bas. https://github.com/VBA-tools/VBA-JSON

Funktion in Excel verwenden – In einer Zelle kannst du nun =OLLAMA("Was ist die Hauptstadt von "& A1 &"?") eingeben, um eine Antwort von Ollama zu erhalten.

Fazit

Mit diesem Setup kannst du Ollama offline in Excel verwenden, um KI-gestützte Antworten direkt in deine Tabellen einzufügen. Das ist besonders nützlich für Automatisierungen, Analysen oder einfach zum Experimentieren mit KI in Office-Umgebungen.

Hast du Fragen oder Verbesserungsvorschläge? Lass es mich wissen!
