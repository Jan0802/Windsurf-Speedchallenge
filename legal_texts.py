"""Rechtstexte an EINER Stelle: Impressum und Datenschutzerklärung.

Warum diese Datei existiert: die Texte standen nur in der Streamlit-App und waren
darum von der Website aus nur über `app.mywatersessions.com/?seite=…` erreichbar.
Damit hingen zwei Pflichtseiten daran, dass ein Streamlit-Server hochfährt, und
sie waren für Suchmaschinen praktisch unsichtbar. Jetzt lesen BEIDE Seiten
denselben Text:

  * die App          -> render_impressum() / render_datenschutz()
  * die Website      -> gen_legal_pages.py schreibt statische HTML-Seiten

Der Inhalt ist absichtlich HTML, nicht Markdown: so muss ihn niemand konvertieren
und beide Ausgaben können nicht auseinanderlaufen. Streamlit gibt ihn mit
`unsafe_allow_html=True` aus, der Generator legt ihn in eine Seitenvorlage.

WICHTIG bei Änderungen: Text hier pflegen, dann `python gen_legal_pages.py`
laufen lassen und die Landing neu deployen. LEGAL_STAND mit hochziehen.
"""

LEGAL_OPERATOR = {
    "name": "Jan Brinkman",
    "street": "Thorner Strasse 12",
    "city": "51469 Bergisch Gladbach",
    "country": "Deutschland",
    "email": "Windsurfspeedchallenge@outlook.de",
}

LEGAL_STAND = "August 2026"


IMPRESSUM_HTML = """
<p class="lead">Angaben gemäß § 5 Digitale-Dienste-Gesetz (DDG)</p>

<h2>Diensteanbieter</h2>
<p>{name}<br>
{street}<br>
{city}<br>
{country}</p>

<h2>Kontakt</h2>
<p>E-Mail: <a href="mailto:{email}">{email}</a></p>

<h2>Verantwortlich für den Inhalt</h2>
<p>nach § 18 Abs. 2 MStV: {name}, Anschrift wie oben.</p>

<h2>Haftung für Inhalte</h2>
<p>Als Diensteanbieter sind wir für eigene Inhalte auf diesen Seiten nach den
allgemeinen Gesetzen verantwortlich. Wir sind jedoch nicht verpflichtet,
übermittelte oder gespeicherte fremde Informationen zu überwachen.</p>

<h2>Haftung für Links</h2>
<p>Unser Angebot enthält Links zu externen Websites Dritter, auf deren Inhalte
wir keinen Einfluss haben. Für diese fremden Inhalte ist stets der jeweilige
Anbieter der Seiten verantwortlich.</p>

<h2>Nutzergenerierte Inhalte</h2>
<p>Ranglisten, Spot-Beschreibungen, Bewertungen und Spot-Fotos stammen zum Teil
von Nutzerinnen und Nutzern. Für solche Inhalte ist die jeweils einstellende
Person verantwortlich. Unzulässige Inhalte entfernen wir nach Kenntnis –
ein Hinweis an die oben genannte E-Mail-Adresse genügt.</p>

<h2>Keine Rettungs- oder Sicherheitsdienstleistung</h2>
<p>MyWaterSessions ist eine Freizeit-Anwendung. Der optionale Sicherheits-Check-in
ist eine Erinnerungsfunktion und <strong>kein Rettungssystem</strong>: er ersetzt
weder eine Absprache am Wasser noch einen Notruf. Angaben zu Spots, Gefahren und
Bedingungen sind Orientierungshilfen ohne Gewähr; die Einschätzung der eigenen
Sicherheit bleibt immer bei der fahrenden Person.</p>
"""


DATENSCHUTZ_HTML = """
<h2>1. Verantwortlicher</h2>
<p>Verantwortlich für die Datenverarbeitung ist:<br>
{name}, {street}, {city}, {country}<br>
E-Mail: <a href="mailto:{email}">{email}</a></p>

<h2>2. Welche Daten wir verarbeiten</h2>

<p><strong>Konto:</strong> Benutzername, E-Mail-Adresse und Passwort. Das Passwort
wird ausschließlich als gesalzener Hash (PBKDF2-HMAC-SHA256) gespeichert – im
Klartext ist es uns nicht bekannt. Optional Gewicht und Körpergröße, die nur in
Empfehlungen (z. B. Segelgröße) und in Spaß-Wertungen einfließen.</p>

<p><strong>Session- und Leistungsdaten:</strong> Pro Session speichern wir Datum,
Sportart, Spot, Material (Board, Segel/Kite/Wing, Finne) sowie die berechneten
Kennzahlen – unter anderem Höchstgeschwindigkeit über 2 s und 30 s, beste 500 m
und Seemeile, avg 5×10, Alpha 500, längster Run, Gesamtstrecke, Dauer und
Aktivzeit, Sprünge und Airtime, Paddelschläge – dazu die zur Session passenden
Wetterdaten (Wind, Böen, Windrichtung, Temperatur) und einen internen
Plausibilitätswert, mit dem offensichtlich unmögliche Werte aus den Ranglisten
ausgeschlossen werden.</p>

<p><strong>GPS-Daten:</strong> Deine hochgeladene Datei bzw. die von deiner Uhr
gesendete Session enthält GPS-Punkte (eine ausgedünnte Route). Diese Route wird
gespeichert und dient der Berechnung deiner Kennzahlen und der Kartenanzeige in
deinem persönlichen Bereich.</p>
<p>Öffentlich sichtbar wird davon <strong>nur ein kurzer Abschnitt</strong>: auf der
Live-Seite und im Spot-TV zeigen wir zu einer Session die Karte des
<strong>schnellsten Streckenabschnitts</strong> (in der Größenordnung von 500 m).
Die <strong>vollständige Route wird nie öffentlich angezeigt</strong> – so ist
weder erkennbar, wo du gestartet bist, noch wo du dein Fahrzeug abgestellt hast.
Du kannst deine Sessions inklusive Track jederzeit unter „Konto &amp; Daten
löschen" entfernen.</p>

<p><strong>Uhr-Verknüpfung:</strong> Zum Zuordnen von Sessions deiner Uhr
speichern wir ein zufälliges Gerätemerkmal (Token bzw. kurzer Verbindungscode).</p>

<p><strong>Sicherheits-Check-in (optional):</strong> Aktivierst du den Check-in,
speichern wir die von dir angegebene <strong>Notfall-E-Mail-Adresse</strong>, die
geplante Rückmeldezeit und den Spot, und versenden bei Überfälligkeit eine
E-Mail an diese Adresse. Bitte gib nur Adressen an, deren Inhaber damit
einverstanden sind – es handelt sich um Daten einer dritten Person. Die Angabe
lässt sich jederzeit löschen.</p>

<p><strong>Spot-Fotos und Bewertungen:</strong> Lädst du ein Spot-Foto hoch oder
bewertest einen Spot, speichern wir dies zusammen mit deinem Benutzernamen; Fotos
zeigen wir öffentlich in der Spot-Galerie. Lade bitte nur eigene, geeignete Fotos
hoch, an denen du die Rechte besitzt. Wir können unpassende Bilder jederzeit
entfernen; auf Anfrage löschen wir deine hochgeladenen Fotos.</p>

<p><strong>Kontaktformular:</strong> Deine Nachricht und, falls angegeben, dein
Name bzw. deine E-Mail-Adresse, um antworten zu können.</p>

<p><strong>Gast-Eintrag:</strong> Über den QR-Code am Spot-TV kann ohne Konto
teilgenommen werden. Dabei speichern wir den eingegebenen Namen und die Werte der
hochgeladenen Datei. Bitte gib keinen Namen an, unter dem du nicht öffentlich
erscheinen möchtest.</p>

<h2>3. Öffentliche Sichtbarkeit</h2>
<p>MyWaterSessions ist eine Community-Bestenliste, und ein Teil der Inhalte ist
<strong>ohne Anmeldung und für Suchmaschinen</strong> sichtbar. Öffentlich sind:
<strong>Benutzername, Datum, Spot, Sportart, Material, die Speed- und
Distanzwerte, die Wetterangaben</strong> sowie der oben beschriebene kurze
Kartenabschnitt. Das betrifft die Ranglisten, die Spot-Seiten, die Live-Seite, die
Rekord-Übersicht und das Spot-TV.</p>
<p>In <strong>privaten Gruppen</strong> sind Ergebnisse nur für bestätigte
Mitglieder sichtbar. Möchtest du nicht öffentlich erscheinen, verwende einen
Benutzernamen ohne Bezug zu deiner Person oder lösche deine Sessions bzw. dein
Konto.</p>

<h2>4. Zwecke und Rechtsgrundlagen</h2>
<ul>
  <li><strong>Konto und Anmeldung</strong> zur Bereitstellung des Dienstes –
      Art. 6 Abs. 1 lit. b DSGVO (Nutzungsverhältnis).</li>
  <li><strong>Veröffentlichung deiner Ergebnisse</strong> in den Ranglisten und auf
      den öffentlichen Seiten – Art. 6 Abs. 1 lit. a DSGVO (Einwilligung, die du
      bei der Registrierung erteilst und jederzeit mit Wirkung für die Zukunft
      widerrufen kannst).</li>
  <li><strong>Sicherheits-Check-in</strong> – Art. 6 Abs. 1 lit. a DSGVO
      (Einwilligung; die Funktion ist freiwillig und einzeln aktivierbar).</li>
  <li><strong>Betrieb, Sicherheit und Missbrauchsabwehr</strong>, einschließlich der
      Plausibilitätsprüfung von Sessions – Art. 6 Abs. 1 lit. f DSGVO
      (berechtigtes Interesse an einer manipulationsfreien Rangliste).</li>
</ul>

<h2>5. Cookies, Messung und Werbung</h2>
<p>Setzt du beim Login „Angemeldet bleiben", speichern wir ein funktional
notwendiges Cookie (<code>surf_auth</code>) mit einem zufälligen Anmelde-Token.
Ohne diese Option werden keine Cookies gesetzt.</p>
<p>Zur Reichweitenmessung nutzen wir <strong>Cloudflare Web Analytics</strong>
(Cloudflare, Inc.). Dieser Dienst arbeitet <strong>ohne Cookies</strong> und ohne
geräteübergreifende Wiedererkennung; erfasst werden aggregierte Angaben wie
aufgerufene Seite, Referrer und technische Rahmendaten.</p>
<p>Auf Spot-Seiten und im Spot-TV können <strong>Logos und Produkthinweise von
Sponsoren</strong> erscheinen. Dabei findet <strong>kein</strong> personalisiertes
Werbe-Tracking und kein Profilaufbau statt.</p>

<h2>6. Hosting und Datenbank</h2>
<p>Website und Anwendung werden bei <strong>Render Services, Inc.</strong> (USA)
betrieben, die Datenbank bei <strong>Neon, Inc.</strong> (USA). Vorgeschaltet ist
<strong>Cloudflare, Inc.</strong> (USA) für Auslieferung und Schutz der Domains.
Dabei werden technisch bedingt Verbindungsdaten (z. B. IP-Adresse) verarbeitet,
auch in den USA. Grundlage für die Übermittlung sind die
EU-Standardvertragsklauseln bzw. das EU-US Data Privacy Framework.</p>

<h2>7. Eingebundene Dienste Dritter</h2>
<p>Beim Aufruf einzelner Funktionen werden Daten an folgende Anbieter übermittelt –
in der Regel deine IP-Adresse, teils Koordinaten:</p>
<ul>
  <li><strong>Open-Meteo</strong> – Wetter und Vorhersage; übermittelt werden die
      Koordinaten des Spots.</li>
  <li><strong>OpenStreetMap</strong> (Kartenkacheln) und <strong>Esri</strong>
      (Satellitenbilder) – Kartenanzeige; dein Browser ruft die Bildkacheln
      direkt dort ab.</li>
  <li><strong>unpkg</strong> – Auslieferung der Kartenbibliothek.</li>
  <li><strong>Nominatim (OpenStreetMap)</strong> – Ermitteln eines Ortsnamens aus
      Koordinaten für neu entstehende Spots.</li>
  <li><strong>is-on-water</strong> – Prüfung, ob Koordinaten auf Wasser liegen
      (Plausibilitätsprüfung); übermittelt werden einzelne Koordinaten ohne
      Kontobezug.</li>
  <li><strong>Anthropic</strong> – Erstellen von Spot-Beschreibungen aus
      Spot- und Wetterangaben; es werden keine Kontodaten übermittelt.</li>
  <li><strong>E-Mail-Versand</strong> – für Bestätigungs-, Hinweis- und
      Check-in-Mails setzen wir einen Versanddienstleister ein; übermittelt wird
      die Empfängeradresse mit dem Nachrichteninhalt.</li>
  <li><strong>Polar</strong> – nur wenn du dein Polar-Konto ausdrücklich
      verbindest: Austausch von Zugriffstoken und Import deiner Aktivitäten.</li>
  <li><strong>Webcams</strong> – auf einzelnen Spot-Seiten binden wir die Livebilder
      der jeweiligen Betreiber ein. Beim Laden erhält der Betreiber deine
      IP-Adresse. Wir haben auf diese Inhalte keinen Einfluss.</li>
</ul>

<h2>8. Speicherdauer</h2>
<p>Konto-, Session- und Gruppendaten speichern wir, bis du dein Konto bzw. die
jeweiligen Einträge löschst oder die Löschung verlangst.</p>

<h2>9. Deine Rechte</h2>
<p>Du hast das Recht auf Auskunft, Berichtigung, Löschung, Einschränkung der
Verarbeitung, Datenübertragbarkeit sowie auf Widerruf erteilter Einwilligungen mit
Wirkung für die Zukunft. Außerdem steht dir ein Beschwerderecht bei einer
Datenschutz-Aufsichtsbehörde zu.</p>
<p>Zur Ausübung deiner Rechte oder zur Löschung deines Kontos genügt eine formlose
Nachricht an <a href="mailto:{email}">{email}</a>. Im eingeloggten Bereich kannst
du dein Konto und alle zugehörigen Daten außerdem selbst unter
„Konto &amp; Daten löschen" entfernen.</p>

<p class="stand">Stand: {stand}. Diese Erklärung wird angepasst, wenn sich die
Datenverarbeitung ändert.</p>
"""


def impressum_html():
    """Impressum als HTML-Fragment (ohne Seitenrahmen)."""
    return IMPRESSUM_HTML.format(**LEGAL_OPERATOR)


def datenschutz_html():
    """Datenschutzerklärung als HTML-Fragment (ohne Seitenrahmen)."""
    return DATENSCHUTZ_HTML.format(stand=LEGAL_STAND, **LEGAL_OPERATOR)
