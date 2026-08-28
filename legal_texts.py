"""Rechtstexte an EINER Stelle: Impressum und Datenschutzerklärung, DE und EN.

Warum diese Datei existiert: die Texte standen nur in der Streamlit-App und waren
darum von der Website aus nur über `app.mywatersessions.com/?seite=…` erreichbar.
Damit hingen zwei Pflichtseiten daran, dass ein Streamlit-Server hochfährt, und
sie waren für Suchmaschinen unsichtbar. Jetzt lesen BEIDE Seiten denselben Text:

  * die App          -> render_impressum() / render_datenschutz()   (deutsch)
  * die Website      -> gen_legal_pages.py schreibt vier statische Seiten
                        /impressum.html   /datenschutz.html   (deutsch)
                        /imprint.html     /privacy.html       (englisch)

Der Inhalt ist absichtlich HTML, nicht Markdown: so muss ihn niemand
konvertieren und die Ausgaben können nicht auseinanderlaufen. Streamlit gibt ihn
mit `unsafe_allow_html=True` aus, der Generator legt ihn in eine Seitenvorlage.

WICHTIG bei Änderungen: Text hier pflegen – und zwar BEIDE Sprachen –, dann
`python gen_legal_pages.py` laufen lassen und die Landing neu deployen.
LEGAL_STAND mit hochziehen.
"""

LEGAL_OPERATOR = {
    "name": "Jan Brinkman",
    "street": "Thorner Strasse 12",
    "city": "51469 Bergisch Gladbach",
    "country": "Deutschland",
    "email": "support@mywatersessions.com",
}

# Englische Fassung der Anschrift (nur das Land wird übersetzt).
LEGAL_OPERATOR_EN = dict(LEGAL_OPERATOR, country="Germany")

LEGAL_STAND = "August 2026"
LEGAL_UPDATED_EN = "August 2026"


# =====================================================================
#  Impressum (deutsch)
# =====================================================================
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


# =====================================================================
#  Imprint (englisch)
# =====================================================================
IMPRINT_HTML_EN = """
<p class="lead">Information pursuant to § 5 of the German Digital Services Act (DDG)</p>

<h2>Service provider</h2>
<p>{name}<br>
{street}<br>
{city}<br>
{country}</p>

<h2>Contact</h2>
<p>E-mail: <a href="mailto:{email}">{email}</a></p>

<h2>Responsible for the content</h2>
<p>pursuant to § 18 (2) MStV: {name}, address as above.</p>

<h2>Liability for content</h2>
<p>As a service provider we are responsible for our own content on these pages
under the general laws. We are not obliged, however, to monitor third-party
information that is transmitted or stored.</p>

<h2>Liability for links</h2>
<p>Our site contains links to external websites whose content is beyond our
control. Responsibility for that content always rests with the provider of the
respective site.</p>

<h2>User-generated content</h2>
<p>Rankings, spot descriptions, ratings and spot photos are partly submitted by
users. The person who submits such content is responsible for it. We remove
unlawful content once we become aware of it – a note to the e-mail address above
is enough.</p>

<h2>Not a rescue or safety service</h2>
<p>MyWaterSessions is a leisure application. The optional safety check-in is a
reminder feature and <strong>not a rescue system</strong>: it replaces neither an
arrangement with someone on the beach nor an emergency call. Information about
spots, hazards and conditions is guidance without warranty; judging your own
safety always remains with you.</p>
"""


# =====================================================================
#  Datenschutzerklärung (deutsch)
# =====================================================================
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
  <li><strong>Resend</strong> (Resend, Inc., USA) – Versand der Konto-E-Mails:
      Bestätigung der Registrierung sowie bei einer Änderung deiner E-Mail-Adresse
      die Bestätigung an die neue und der Sicherheitshinweis an die alte Adresse.
      Übermittelt werden Empfängeradresse und Nachrichteninhalt.</li>
  <li><strong>Brevo</strong> (Brevo SAS, Frankreich) – Versand der Mails des
      Sicherheits-Check-ins. Übermittelt werden die von dir angegebene
      Notfall-Adresse und der Nachrichteninhalt.</li>
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


# =====================================================================
#  Privacy policy (englisch)
# =====================================================================
PRIVACY_HTML_EN = """
<h2>1. Controller</h2>
<p>The controller for the processing described here is:<br>
{name}, {street}, {city}, {country}<br>
E-mail: <a href="mailto:{email}">{email}</a></p>

<h2>2. What data we process</h2>

<p><strong>Account:</strong> username, e-mail address and password. The password is
stored only as a salted hash (PBKDF2-HMAC-SHA256) – we never know it in plain
text. Optionally weight and height, which are used only for recommendations
(e.g. sail size) and for the fun rankings.</p>

<p><strong>Session and performance data:</strong> for each session we store the
date, sport, spot, gear (board, sail/kite/wing, fin) and the values we compute –
among them top speed over 2 s and 30 s, best 500 m and nautical mile, avg 5×10,
Alpha 500, longest run, total distance, duration and active time, jumps and
airtime, paddle strokes – plus the weather that matched the session (wind, gusts,
wind direction, temperature) and an internal plausibility score used to keep
clearly impossible values out of the rankings.</p>

<p><strong>GPS data:</strong> the file you upload, or the session your watch sends,
contains GPS points (a thinned route). This route is stored and used to compute
your values and to draw the map in your personal area.</p>
<p>Only a <strong>short section of it becomes public</strong>: on the Live page and
on Spot-TV we show the map of the <strong>fastest stretch</strong> of a session
(in the order of 500 m). The <strong>complete route is never shown
publicly</strong>, so neither where you started nor where you parked can be seen.
You can delete your sessions including their track at any time under
“Delete account &amp; data”.</p>

<p><strong>Watch link:</strong> to attribute sessions from your watch we store a
random device identifier (a token or a short connect code).</p>

<p><strong>Safety check-in (optional):</strong> if you switch the check-in on, we
store the <strong>emergency e-mail address</strong> you enter, the time you expect
to be back and the spot, and we send an e-mail to that address if you are overdue.
Please only enter addresses whose owner agrees to this – it is another person's
data. The entry can be deleted at any time.</p>

<p><strong>Spot photos and ratings:</strong> if you upload a spot photo or rate a
spot, we store it together with your username; photos are shown publicly in the
spot gallery. Please upload only your own, suitable photos, to which you hold the
rights. We may remove unsuitable images at any time, and we delete your uploads on
request.</p>

<p><strong>Contact form:</strong> your message and, if given, your name and e-mail
address so that we can reply.</p>

<p><strong>Guest entry:</strong> the QR code on Spot-TV lets people take part
without an account. We then store the name entered and the values from the
uploaded file. Please do not enter a name under which you do not wish to appear
publicly.</p>

<h2>3. What is publicly visible</h2>
<p>MyWaterSessions is a community leaderboard, and part of its content is visible
<strong>without an account and to search engines</strong>. Public are:
<strong>username, date, spot, sport, gear, the speed and distance values and the
weather</strong>, plus the short map section described above. This applies to the
rankings, the spot pages, the Live page, the record overview and Spot-TV.</p>
<p>In <strong>private groups</strong>, results are visible only to confirmed
members. If you prefer not to appear publicly, use a username unconnected to your
identity, or delete your sessions or your account.</p>

<h2>4. Purposes and legal bases</h2>
<ul>
  <li><strong>Account and sign-in</strong> to provide the service – Art. 6 (1) (b)
      GDPR (performance of a contract).</li>
  <li><strong>Publishing your results</strong> in the rankings and on the public
      pages – Art. 6 (1) (a) GDPR (consent, given at registration and
      withdrawable at any time with effect for the future).</li>
  <li><strong>Safety check-in</strong> – Art. 6 (1) (a) GDPR (consent; the feature
      is optional and switched on individually).</li>
  <li><strong>Operation, security and abuse prevention</strong>, including the
      plausibility check of sessions – Art. 6 (1) (f) GDPR (legitimate interest in
      a leaderboard that cannot be gamed).</li>
</ul>

<h2>5. Cookies, measurement and advertising</h2>
<p>If you tick “Stay signed in” at login, we store one functionally necessary
cookie (<code>surf_auth</code>) holding a random sign-in token. Without that
option no cookies are set.</p>
<p>For audience measurement we use <strong>Cloudflare Web Analytics</strong>
(Cloudflare, Inc.). The service works <strong>without cookies</strong> and without
cross-device recognition; what is collected is aggregated information such as the
page viewed, the referrer and technical parameters.</p>
<p>Spot pages and Spot-TV may show <strong>sponsor logos and product
listings</strong>. There is <strong>no</strong> personalised ad tracking and no
profiling.</p>

<h2>6. Hosting and database</h2>
<p>Website and application run at <strong>Render Services, Inc.</strong> (USA),
the database at <strong>Neon, Inc.</strong> (USA), with
<strong>Cloudflare, Inc.</strong> (USA) in front for delivery and protection of the
domains. Connection data (e.g. IP address) is technically necessarily processed,
including in the USA. Transfers are based on the EU standard contractual clauses
and/or the EU-US Data Privacy Framework.</p>

<h2>7. Third-party services</h2>
<p>When certain features are used, data is transmitted to the following
providers – usually your IP address, in some cases coordinates:</p>
<ul>
  <li><strong>Open-Meteo</strong> – weather and forecast; the spot's coordinates are
      transmitted.</li>
  <li><strong>OpenStreetMap</strong> (map tiles) and <strong>Esri</strong>
      (satellite imagery) – map display; your browser fetches the image tiles
      directly from them.</li>
  <li><strong>unpkg</strong> – delivery of the map library.</li>
  <li><strong>Nominatim (OpenStreetMap)</strong> – deriving a place name from
      coordinates for newly created spots.</li>
  <li><strong>is-on-water</strong> – checking whether coordinates are on water
      (plausibility check); single coordinates are transmitted, with no account
      reference.</li>
  <li><strong>Anthropic</strong> – generating spot descriptions from spot and
      weather information; no account data is transmitted.</li>
  <li><strong>Resend</strong> (Resend, Inc., USA) – sending the account e-mails:
      the registration confirmation and, when you change your e-mail address, the
      confirmation to the new address and the security notice to the old one.
      Recipient address and message content are transmitted.</li>
  <li><strong>Brevo</strong> (Brevo SAS, France) – sending the safety check-in
      e-mails. The emergency address you entered and the message content are
      transmitted.</li>
  <li><strong>Polar</strong> – only if you explicitly connect your Polar account:
      exchange of access tokens and import of your activities.</li>
  <li><strong>Webcams</strong> – on some spot pages we embed the live image of the
      respective operator. Loading it gives that operator your IP address. We have
      no influence over this content.</li>
</ul>

<h2>8. Retention</h2>
<p>We keep account, session and group data until you delete your account or the
individual entries, or ask us to delete them.</p>

<h2>9. Your rights</h2>
<p>You have the right of access, rectification, erasure, restriction of processing
and data portability, and the right to withdraw consent with effect for the
future. You also have the right to lodge a complaint with a data protection
supervisory authority.</p>
<p>To exercise your rights or delete your account, an informal message to
<a href="mailto:{email}">{email}</a> is enough. When signed in you can also remove
your account and all related data yourself under “Delete account &amp; data”.</p>

<p class="stand">Last updated: {stand}. This policy is amended whenever our
processing changes.</p>
"""


def impressum_html():
    """Impressum als HTML-Fragment (ohne Seitenrahmen), deutsch."""
    return IMPRESSUM_HTML.format(**LEGAL_OPERATOR)


def datenschutz_html():
    """Datenschutzerklärung als HTML-Fragment (ohne Seitenrahmen), deutsch."""
    return DATENSCHUTZ_HTML.format(stand=LEGAL_STAND, **LEGAL_OPERATOR)


def imprint_html_en():
    """Imprint as an HTML fragment (no page frame), English."""
    return IMPRINT_HTML_EN.format(**LEGAL_OPERATOR_EN)


def privacy_html_en():
    """Privacy policy as an HTML fragment (no page frame), English."""
    return PRIVACY_HTML_EN.format(stand=LEGAL_UPDATED_EN, **LEGAL_OPERATOR_EN)
