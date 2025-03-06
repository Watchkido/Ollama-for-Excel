In diesem Beitrag zeige ich dir, wie du Ollama lokal auf deinem Computer betreibst und die KI direkt in Excel mit VBA ansprichst. Diese Methode erlaubt es dir, AI-gestützte Funktionen zu nutzen, ohne auf eine Internetverbindung angewiesen zu sein.

Voraussetzungen

Ollama lokal installiert – Falls noch nicht geschehen, lade Ollama herunter und installiere es. Es sollte als lokaler Server laufen.

Excel mit aktiviertem VBA – Du benötigst eine Excel-Version, die Makros unterstützt (z. B. Excel für Windows).

JSON-Parser für VBA – Falls du JSON-Daten verarbeiten möchtest, lade die JSONConverter.bas von VBA JSON herunter und importiere sie in dein VBA-Projekt.

VBA-Code zur Anbindung von Ollama

Füge folgenden VBA-Code in dein Excel-VBA-Modul ein:
--------------------------------------------
Option Explicit

Function OLLAMA(query As String) As String
    Dim http As Object
    Dim json As Object
    Dim url As String
    Dim payload As String
    Dim response As String
    Dim result As String
    
    ' Setze die API-URL
    url = "http://127.0.0.1:11434/v1/chat/completions"
    
    ' Erstelle das JSON-Payload
    payload = "{""model"": ""llama3.2"", ""messages"": [{""role"": ""system"", ""content"": ""Du bist ein hilfreicher KI-Assistent. Antworte Kurz und Prägnant.""}, {""role"": ""user"", ""content"": """ & query & """}], ""temperature"": 0.3}"
    
    ' Erstelle die HTTP-Anfrage
    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "POST", url, False
    http.setRequestHeader "Content-Type", "application/json"
    http.setRequestHeader "Accept-Charset", "UTF-8"
    
    ' Sende die Anfrage
    http.send payload
    
    ' Hole die Antwort
    response = http.responseText
    
    ' Debugging: Ausgabe der Antwort
    Debug.Print "Response: " & response
    
    ' Parse die JSON-Antwort
    On Error GoTo ErrorHandler
    Set json = JsonConverter.ParseJson(response)
    
    ' Extrahiere das Ergebnis
    result = json("choices")(1)("message")("content")
    
    ' Gib das Ergebnis zurück
    OLLAMA = result
    Exit Function
    
ErrorHandler:
    OLLAMA = "Fehler beim Parsen der JSON-Antwort: " & Err.Description
End Function

-----------------------------------------------------------
Anleitung zur Nutzung

VBA-Editor öffnen – Drücke ALT + F11 in Excel. Wenn nötig schalte die Entwicklertools frei!

Neues Modul erstellen – Gehe zu „Einfügen“ > „Modul“.

Code einfügen – Kopiere den obigen VBA-Code in das Modul.

JSON-Parser hinzufügen – Falls noch nicht geschehen, importiere die JsonConverter.bas.

Funktion in Excel verwenden – In einer Zelle kannst du nun =OLLAMA("Was ist die Hauptstadt von "& A1 &"?") eingeben, um eine Antwort von Ollama zu erhalten.

Fazit

Mit diesem Setup kannst du Ollama offline in Excel verwenden, um KI-gestützte Antworten direkt in deine Tabellen einzufügen. Das ist besonders nützlich für Automatisierungen, Analysen oder einfach zum Experimentieren mit KI in Office-Umgebungen.

Hast du Fragen oder Verbesserungsvorschläge? Lass es mich wissen!
