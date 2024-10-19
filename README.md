Das Projekt nimmt als Basis ein Excel File das den Spielplan des CCU enthält. Es extrahiert daraus alle Spieldaten der einzelnen Teams. Diese werden exportiert und in das Excel als neue Tabellenblätter geschrieben.

Es wurde vor langer Zeit gestartet und basiert noch auf .Net 4.8
Um die Excel Files zu handeln wird EPPlus verwendet. package id="EPPlus" version="4.5.3.1" targetFramework="net452"

Ein kleiner Ausschnitt aus dem Spielplan: ![grafik](https://github.com/user-attachments/assets/2639bd9e-3f2b-4237-aa74-5ed887c95b29)

Da ich der einzige Entwickler bin wurde bezüglich Branches und auch Code Quality sehr unsauber gearbeitet.
Neben gelegentlichen Erweiterungen für andere Ausgabeformate, wird keine Wartung betrieben.
Der Code eigent sich somit nicht für die Verwendung als Beispiel von anderen Projekten. Jede Verantwortung wird abgelehnt!
