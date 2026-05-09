/**
 * Olgahaus Essensbestellung – Google Apps Script
 * ================================================
 * Tabs:
 *   1. Bewohner                      – Stammdaten (Apt/PIN/Name)
 *   2. Bestellungen                  – Rohdaten der App (Schreibziel)
 *   3. Wochenübersicht Küche         – Was bekommt jeder Bewohner Mo–So?
 *   4. Monatsübersicht Buchhaltung   – Wie viele Mahlzeiten/Bewohner/Monat?
 *
 * Architektur (NEU – robuster Ansatz):
 *   Statt zerbrechlicher Cross-Sheet-Formeln (FILTER, INDEX/MATCH, ARRAYFORMULA)
 *   werden die Übersichten direkt aus Apps Script befüllt:
 *   - aktualisiereUebersichten() liest Bewohner + Bestellungen, schreibt Werte
 *   - onEdit-Trigger ruft aktualisiereUebersichten() auf, wenn KW/Monat/Jahr in
 *     den Tabs geändert wird
 *   - doPost() ruft aktualisiereUebersichten() nach jeder neuen Bestellung auf
 *
 * EINRICHTUNG (einmalig):
 *   Erweiterungen → Apps Script → Code einfügen → Speichern → initAlles ausführen
 *
 * Web-App-Bereitstellung:
 *   Bereitstellen → Web-App → Ausführen als: Ich, Zugriff: Jeder
 */

// ── Konfiguration ─────────────────────────────────────────────────────────────
var FARBE_TUERKIS = "#03ACB3";
var FARBE_GRUEN   = "#8DC73F";
var FARBE_ANTHRA  = "#292C2D";
var FARBE_GELB    = "#FFF4D6";
var FARBE_HGRUEN  = "#EEF8E0";
var FARBE_HROT    = "#FBE4E4";

var WOCHE_NAME = "Wochenübersicht Küche";
var MONAT_NAME = "Monatsübersicht Buchhaltung";
var TAGE_DE    = ["Montag","Dienstag","Mittwoch","Donnerstag","Freitag","Samstag","Sonntag"];

// Spiegel von MAHLZEITEN_STANDARD aus der HTML-App.
// Wenn der Speiseplan sich ändert, müssen beide Stellen aktualisiert werden.
var MENUE_PRO_WOCHENTAG = {
  "Montag":     { vor: "Kartoffelsuppe",               nach: "Frisches Obst" },
  "Dienstag":   { vor: "Kraftbrühe mit Grießklößchen", nach: "Naturjoghurt" },
  "Mittwoch":   { vor: "Broccoli-Cremesuppe",          nach: "Kompott" },
  "Donnerstag": { vor: "Kraftbrühe mit Flädle",        nach: "Fruchtjoghurt" },
  "Freitag":    { vor: "Blumenkohl-Cremesuppe",        nach: "Frisches Obst" },
  "Samstag":    { vor: "Tomaten-Cremesuppe",           nach: "Schokoladenpudding" },
  "Sonntag":    { vor: "Kraftbrühe mit Backerbsen",    nach: "Frisches Obst" }
};

// ──────────────────────────────────────────────
// 1. INIT – einmalig alle Tabs anlegen & befüllen
// ──────────────────────────────────────────────
function initAlles() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Standard-Tab umbenennen falls noch "Tabellenblatt1"
  var erstesBl = ss.getSheets()[0];
  if (erstesBl.getName() === "Tabellenblatt1") erstesBl.setName("Bewohner");

  // Veraltete Tabs entfernen
  ["Tagesübersicht", "Bewohner KW"].forEach(function(name){
    var alt = ss.getSheetByName(name);
    if (alt) ss.deleteSheet(alt);
  });

  initBewohner(ss);
  initBestellungen(ss);
  repariereBestellungenStill(ss);
  initWochenuebersichtKueche(ss);
  initMonatsuebersichtBuchhaltung(ss);

  // Reihenfolge der Tabs
  var reihenfolge = ["Bewohner", "Bestellungen", WOCHE_NAME, MONAT_NAME];
  for (var i = 0; i < reihenfolge.length; i++) {
    var bl = ss.getSheetByName(reihenfolge[i]);
    if (bl) { ss.setActiveSheet(bl); ss.moveActiveSheet(i + 1); }
  }

  // Übersichten direkt mit Daten befüllen
  aktualisiereUebersichten();

  // onEdit-Trigger automatisch installieren
  installiereTrigger();

  ss.setActiveSheet(ss.getSheetByName(WOCHE_NAME));
  SpreadsheetApp.getUi().alert(
    "✅ Einrichtung abgeschlossen!\n\n" +
    "4 Tabs sind eingerichtet:\n" +
    "  1. Bewohner – Stammdaten eintragen\n" +
    "  2. Bestellungen – wird von der App befüllt\n" +
    "  3. Wochenübersicht Küche – KW oben wählen → aktualisiert sich automatisch\n" +
    "  4. Monatsübersicht Buchhaltung – Monat oben wählen → aktualisiert sich automatisch\n\n" +
    "Trigger installiert: Werte aktualisieren sich beim Ändern von KW/Monat/Jahr."
  );
}

// ──────────────────────────────────────────────
// 2. Bewohner-Tab
// ──────────────────────────────────────────────
function initBewohner(ss) {
  var b = ss.getSheetByName("Bewohner");
  if (!b) b = ss.insertSheet("Bewohner");

  if (b.getLastRow() === 0) {
    var header = ["Apartment", "PIN", "Name"];
    b.getRange(1, 1, 1, 3).setValues([header])
      .setFontWeight("bold")
      .setBackground(FARBE_TUERKIS)
      .setFontColor("#FFFFFF")
      .setHorizontalAlignment("center");
    var leereZeilen = [];
    for (var i = 0; i < 100; i++) leereZeilen.push(["", "", ""]);
    b.getRange(2, 1, 100, 3).setValues(leereZeilen);
    b.setColumnWidth(1, 120); b.setColumnWidth(2, 100); b.setColumnWidth(3, 250);
    b.setFrozenRows(1);
  }
  b.getRange("A:A").setNumberFormat("@");
  b.getRange("A1").setNote(
    "Bitte alle Bewohner eintragen:\n" +
    "Spalte A: Apartment-Nummer (z.B. 12 oder A3)\n" +
    "Spalte B: 5-stelliger PIN-Code\n" +
    "Spalte C: Vollständiger Name"
  );
}

// ──────────────────────────────────────────────
// 3. Bestellungen-Tab
// ──────────────────────────────────────────────
function initBestellungen(ss) {
  var b = ss.getSheetByName("Bestellungen");
  if (!b) b = ss.insertSheet("Bestellungen");

  var header = ["Datum","Wochentag","KW","Jahr","Apartment","Name",
                "Menütyp","Menüname","Vorspeise","Nachspeise","Preis (€)","Zeitstempel","Typ"];

  if (b.getLastRow() === 0) {
    b.getRange(1, 1, 1, header.length).setValues([header])
      .setFontWeight("bold")
      .setBackground(FARBE_ANTHRA)
      .setFontColor("#FFFFFF");
    b.setFrozenRows(1);
    var breiten = [100,100,50,60,100,180,130,280,90,90,80,140,130];
    for (var i = 0; i < breiten.length; i++) b.setColumnWidth(i + 1, breiten[i]);
  }

  b.getRange("A:A").setNumberFormat("@");
  b.getRange("E:E").setNumberFormat("@");
}

// ──────────────────────────────────────────────
// 4. Wochenübersicht Küche – nur Header + Dropdowns
// ──────────────────────────────────────────────
function initWochenuebersichtKueche(ss) {
  var b = ss.getSheetByName(WOCHE_NAME);
  if (!b) b = ss.insertSheet(WOCHE_NAME);
  b.clear();
  b.clearConditionalFormatRules();

  // Zeile 1: KW + Jahr-Auswahl
  b.getRange("A1").setValue("Kalenderwoche:");
  b.getRange("C1").setValue("Jahr:");
  b.getRange(1, 1, 1, 9).setBackground(FARBE_TUERKIS).setFontColor("#FFFFFF").setFontWeight("bold");
  b.getRange("A1").setHorizontalAlignment("right");
  b.getRange("C1").setHorizontalAlignment("right");

  // KW + Jahr als feste Werte (keine Formel) – Trigger reagiert auf Änderungen
  var heute = new Date();
  b.getRange("B1").setValue(isoWeekNumber(heute));
  b.getRange("D1").setValue(heute.getFullYear());
  b.getRange("B1:D1").setBackground(FARBE_GELB).setFontColor(FARBE_ANTHRA)
    .setFontWeight("bold").setHorizontalAlignment("center");

  var kwListe = []; for (var i = 1; i <= 53; i++) kwListe.push(i);
  setDropdown(b.getRange("B1"), kwListe);
  setDropdown(b.getRange("D1"), [2025, 2026, 2027, 2028, 2029, 2030]);
  b.getRange("B1").setNote("KW wechseln → Tabelle aktualisiert sich automatisch.");

  // Zeile 3: nur Apartment/Name fix – Wochentag-Header (mit Datum) wird in
  // aktualisiereWochenuebersicht() dynamisch gesetzt
  b.getRange(3, 1, 1, 2).setValues([["Apartment","Name"]])
    .setFontWeight("bold")
    .setBackground(FARBE_ANTHRA)
    .setFontColor("#FFFFFF")
    .setHorizontalAlignment("center");
  b.setFrozenRows(3); b.setFrozenColumns(2);
  b.getRange("A:A").setNumberFormat("@");

  for (var c = 3; c <= 9; c++) b.setColumnWidth(c, 200);
  b.setColumnWidth(1, 100); b.setColumnWidth(2, 200);

  // Bedingte Formatierung: "✗ Kein Essen" rot
  var menuRange = b.getRange(4, 3, 200, 7);
  var regelRot = SpreadsheetApp.newConditionalFormatRule()
    .whenTextContains("Kein Essen")
    .setBackground(FARBE_HROT)
    .setFontColor("#A33A3A")
    .setRanges([menuRange])
    .build();
  b.setConditionalFormatRules([regelRot]);

  b.getRange("A1").setNote(
    "WOCHENÜBERSICHT KÜCHE\n\n" +
    "B1: KW wählen\n" +
    "D1: Jahr wählen\n\n" +
    "→ Tabelle zeigt für jeden Bewohner das konkrete Menü pro Tag.\n" +
    "→ Bei Änderung der KW oder des Jahres aktualisiert sich die Tabelle automatisch."
  );
}

// ──────────────────────────────────────────────
// 5. Monatsübersicht Buchhaltung – nur Header + Dropdowns
// ──────────────────────────────────────────────
function initMonatsuebersichtBuchhaltung(ss) {
  var b = ss.getSheetByName(MONAT_NAME);
  if (!b) b = ss.insertSheet(MONAT_NAME);
  b.clear();
  b.clearConditionalFormatRules();

  // Zeile 1: Monat + Jahr-Auswahl
  b.getRange("A1").setValue("Monat:");
  b.getRange("C1").setValue("Jahr:");
  b.getRange(1, 1, 1, 10).setBackground(FARBE_TUERKIS).setFontColor("#FFFFFF").setFontWeight("bold");
  b.getRange("A1").setHorizontalAlignment("right");
  b.getRange("C1").setHorizontalAlignment("right");

  var heute = new Date();
  b.getRange("B1").setValue(heute.getMonth() + 1);
  b.getRange("D1").setValue(heute.getFullYear());
  b.getRange("B1:D1").setBackground(FARBE_GELB).setFontColor(FARBE_ANTHRA)
    .setFontWeight("bold").setHorizontalAlignment("center");

  var monatListe = []; for (var i = 1; i <= 12; i++) monatListe.push(i);
  setDropdown(b.getRange("B1"), monatListe);
  setDropdown(b.getRange("D1"), [2025, 2026, 2027, 2028, 2029, 2030]);
  b.getRange("B1").setNote("Monat wechseln → Tabelle aktualisiert sich automatisch.");

  // Zeile 3: Header
  var header = ["Apartment","Name","Hauptmenü","Alt 1","Alt 2","Kein Essen",
                "Vorspeisen","Nachspeisen","Gesamt Essen","Gesamtkosten (€)"];
  b.getRange(3, 1, 1, header.length).setValues([header])
    .setFontWeight("bold")
    .setBackground(FARBE_ANTHRA)
    .setFontColor("#FFFFFF")
    .setHorizontalAlignment("center");
  b.setFrozenRows(3); b.setFrozenColumns(2);
  b.getRange("A:A").setNumberFormat("@");

  b.setColumnWidth(1, 100); b.setColumnWidth(2, 200);
  for (var c = 3; c <= 8; c++) b.setColumnWidth(c, 100);
  b.setColumnWidth(9, 110); b.setColumnWidth(10, 130);

  b.getRange("A1").setNote(
    "MONATSÜBERSICHT BUCHHALTUNG\n\n" +
    "B1: Monat (1–12)\n" +
    "D1: Jahr\n\n" +
    "→ Pro Bewohner: Anzahl Mahlzeiten je Typ + Gesamt + Kosten.\n" +
    "→ Σ-Zeile unten: Summen über alle Bewohner."
  );
}

// ──────────────────────────────────────────────
// 6. AKTUALISIEREN – Übersichten direkt befüllen
// ──────────────────────────────────────────────
function aktualisiereUebersichten() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  aktualisiereWochenuebersicht(ss);
  aktualisiereMonatsuebersicht(ss);
}

function aktualisiereWochenuebersicht(ss) {
  var b = ss.getSheetByName(WOCHE_NAME);
  if (!b) return;

  // Datenbereich leeren (Zeile 4 abwärts)
  var alteZeilen = b.getMaxRows() - 3;
  if (alteZeilen > 0) b.getRange(4, 1, alteZeilen, 9).clearContent();

  var kw    = Number(b.getRange("B1").getValue());
  var jahr  = Number(b.getRange("D1").getValue());
  if (!kw || !jahr) return;

  // Header (Zeile 3) dynamisch mit Wochentag + Datum setzen
  var mo = montagVonKW(kw, jahr);
  var headerTage = TAGE_DE.map(function(tag, i){
    var d = new Date(mo); d.setDate(mo.getDate() + i);
    return tag + "\n" + fmtKurzDatum(d);
  });
  b.getRange(3, 1, 1, 9).setValues([["Apartment","Name"].concat(headerTage)])
    .setFontWeight("bold")
    .setBackground(FARBE_ANTHRA)
    .setFontColor("#FFFFFF")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setWrap(true);
  b.setRowHeight(3, 42);

  var bewohner = ladeBewohner(ss);
  if (bewohner.length === 0) {
    b.getRange(4, 1).setValue("Bitte zuerst Bewohner im Tab 'Bewohner' eintragen.")
      .setFontStyle("italic").setFontColor("#888888");
    return;
  }

  var bestellungen = ladeBestellungen(ss);

  // Index für schnelle Suche: key "Apt|Wochentag|KW|Jahr" → o
  var idx = {};
  bestellungen.forEach(function(o){
    var key = o.apartment + "|" + o.wochentag + "|" + o.kw + "|" + o.jahr;
    idx[key] = o;
  });

  // Zelleninhalt formatieren: Hauptgericht + ggf. Vor-/Nachspeise
  function formatZelle(o, tag) {
    if (!o) return "";
    if (o.menueTyp === "Kein Essen") return "✗ Kein Essen";
    var teile = [o.menueName || ""];
    var menue = MENUE_PRO_WOCHENTAG[tag];
    if (o.vorspeise  === "Ja" && menue) teile.push("✓ " + menue.vor);
    if (o.nachspeise === "Ja" && menue) teile.push("✓ " + menue.nach);
    return teile.join("\n");
  }

  // Werte-Matrix bauen
  var matrix = [];
  bewohner.forEach(function(bw){
    var zeile = [bw.apartment, bw.name];
    TAGE_DE.forEach(function(tag){
      var key = bw.apartment + "|" + tag + "|" + kw + "|" + jahr;
      zeile.push(formatZelle(idx[key], tag));
    });
    matrix.push(zeile);
  });

  if (matrix.length > 0) {
    b.getRange(4, 1, matrix.length, 9).setValues(matrix);
    b.getRange(4, 3, matrix.length, 7).setWrap(true).setVerticalAlignment("top");
    for (var r = 0; r < matrix.length; r++) {
      b.setRowHeight(4 + r, 70); // Platz für Hauptgericht + Vor + Nach
    }
  }
}

function aktualisiereMonatsuebersicht(ss) {
  var b = ss.getSheetByName(MONAT_NAME);
  if (!b) return;

  var alteZeilen = b.getMaxRows() - 3;
  if (alteZeilen > 0) b.getRange(4, 1, alteZeilen, 10).clearContent();

  var monat = Number(b.getRange("B1").getValue());
  var jahr  = Number(b.getRange("D1").getValue());
  if (!monat || !jahr) return;

  var bewohner = ladeBewohner(ss);
  if (bewohner.length === 0) {
    b.getRange(4, 1).setValue("Bitte zuerst Bewohner im Tab 'Bewohner' eintragen.")
      .setFontStyle("italic").setFontColor("#888888");
    return;
  }

  var bestellungen = ladeBestellungen(ss);

  // Pro Apartment aggregieren – nur Bestellungen, deren Datum zum Monat/Jahr passt
  var agg = {};
  bewohner.forEach(function(bw){
    agg[bw.apartment] = {haupt:0, alt1:0, alt2:0, kein:0, vor:0, nach:0, kosten:0};
  });

  bestellungen.forEach(function(o){
    if (!agg[o.apartment]) return; // Bewohner nicht in Liste
    var d = parseDatum(o.datum);
    if (!d) return;
    if (d.getMonth() + 1 !== monat) return;
    if (d.getFullYear() !== jahr) return;

    var a = agg[o.apartment];
    if      (o.menueTyp === "Hauptmenü")     a.haupt++;
    else if (o.menueTyp === "Alternative 1") a.alt1++;
    else if (o.menueTyp === "Alternative 2") a.alt2++;
    else if (o.menueTyp === "Kein Essen")    a.kein++;
    if (o.vorspeise === "Ja")  a.vor++;
    if (o.nachspeise === "Ja") a.nach++;
    a.kosten += Number(o.preis) || 0;
  });

  var matrix = [];
  var summen = {haupt:0, alt1:0, alt2:0, kein:0, vor:0, nach:0, gesamt:0, kosten:0};
  bewohner.forEach(function(bw){
    var a = agg[bw.apartment];
    var gesamt = a.haupt + a.alt1 + a.alt2;
    matrix.push([bw.apartment, bw.name, a.haupt, a.alt1, a.alt2, a.kein,
                 a.vor, a.nach, gesamt, a.kosten]);
    summen.haupt += a.haupt;  summen.alt1 += a.alt1;  summen.alt2 += a.alt2;
    summen.kein  += a.kein;   summen.vor  += a.vor;   summen.nach += a.nach;
    summen.gesamt += gesamt;  summen.kosten += a.kosten;
  });

  if (matrix.length > 0) {
    b.getRange(4, 1, matrix.length, 10).setValues(matrix);
    for (var c = 3; c <= 9; c++) {
      b.getRange(4, c, matrix.length, 1).setHorizontalAlignment("center");
    }
    b.getRange(4, 10, matrix.length, 1).setNumberFormat("0.00 €");
  }

  // Σ-Zeile direkt nach den Bewohnern
  var sumZeile = matrix.length + 4;
  b.getRange(sumZeile, 1, 1, 10).setValues([[
    "Σ Gesamt", "(alle Bewohner)",
    summen.haupt, summen.alt1, summen.alt2, summen.kein,
    summen.vor, summen.nach, summen.gesamt, summen.kosten
  ]]);
  b.getRange(sumZeile, 1, 1, 10).setBackground(FARBE_HGRUEN).setFontWeight("bold");
  b.getRange(sumZeile, 10, 1, 1).setNumberFormat("0.00 €");
  b.getRange(sumZeile, 2).setFontStyle("italic").setFontColor("#666666");
}

// ──────────────────────────────────────────────
// 7. Datenlader (Helper)
// ──────────────────────────────────────────────
function ladeBewohner(ss) {
  var sh = ss.getSheetByName("Bewohner");
  if (!sh || sh.getLastRow() < 2) return [];
  var rows = sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues();
  var out = [];
  rows.forEach(function(r){
    var apt  = String(r[0] || "").trim();
    var name = String(r[2] || "").trim();
    if (apt && name && apt.indexOf("←") < 0) {
      out.push({apartment: apt, name: name});
    }
  });
  return out;
}

function ladeBestellungen(ss) {
  var sh = ss.getSheetByName("Bestellungen");
  if (!sh || sh.getLastRow() < 2) return [];
  var rows = sh.getRange(2, 1, sh.getLastRow() - 1, 13).getValues();
  return rows.map(function(r){
    return {
      datum:      String(r[0] || "").trim(),
      wochentag:  String(r[1] || "").trim(),
      kw:         Number(r[2]) || 0,
      jahr:       Number(r[3]) || 0,
      apartment:  String(r[4] || "").trim(),
      name:       String(r[5] || "").trim(),
      menueTyp:   String(r[6] || "").trim(),
      menueName:  String(r[7] || "").trim(),
      vorspeise:  String(r[8] || "").trim(),
      nachspeise: String(r[9] || "").trim(),
      preis:      Number(r[10]) || 0
    };
  });
}

// ──────────────────────────────────────────────
// 8. onEdit-Trigger – aktualisiert beim Ändern von KW/Monat/Jahr
// ──────────────────────────────────────────────
function onEdit(e) {
  try {
    if (!e || !e.range) return;
    var sheet = e.range.getSheet().getName();
    var row   = e.range.getRow();
    var col   = e.range.getColumn();
    if (row !== 1) return;
    if (col !== 2 && col !== 4) return;

    if (sheet === WOCHE_NAME) aktualisiereWochenuebersicht(SpreadsheetApp.getActiveSpreadsheet());
    if (sheet === MONAT_NAME) aktualisiereMonatsuebersicht(SpreadsheetApp.getActiveSpreadsheet());
  } catch (err) {
    Logger.log("onEdit-Fehler: " + err.message);
  }
}

function installiereTrigger() {
  var existing = ScriptApp.getProjectTriggers();
  var hatOnEdit = existing.some(function(t){
    return t.getHandlerFunction() === "onEdit" && t.getEventType() === ScriptApp.EventType.ON_EDIT;
  });
  if (!hatOnEdit) {
    ScriptApp.newTrigger("onEdit")
      .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
      .onEdit()
      .create();
  }
}

// ──────────────────────────────────────────────
// 9. Helpers
// ──────────────────────────────────────────────
function setDropdown(range, werte) {
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(werte.map(String), true)
    .setAllowInvalid(false)
    .build();
  range.setDataValidation(rule);
}

// Datum "DD.MM.YYYY" → Date-Objekt
function parseDatum(s) {
  if (!s) return null;
  var m = String(s).match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})$/);
  if (!m) {
    // Fallback: vielleicht doch ein Date-Objekt
    var d = new Date(s);
    if (!isNaN(d.getTime())) return d;
    return null;
  }
  return new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1]));
}

// Montag einer beliebigen ISO-KW (Inverse von isoWeekNumber)
function montagVonKW(kw, jahr) {
  var jan4 = new Date(jahr, 0, 4);
  var dow  = jan4.getDay() || 7;
  var mo1  = new Date(jan4);
  mo1.setDate(jan4.getDate() - dow + 1);
  var mo = new Date(mo1);
  mo.setDate(mo1.getDate() + (kw - 1) * 7);
  return mo;
}

// Datum kurz: "11.05.2026"
function fmtKurzDatum(d) {
  return ("0"+d.getDate()).slice(-2) + "." +
         ("0"+(d.getMonth()+1)).slice(-2) + "." + d.getFullYear();
}

// ISO-Wochennummer (Donnerstag der Woche bestimmt das Jahr)
function isoWeekNumber(d) {
  var t = new Date(d.getTime());
  t.setHours(0, 0, 0, 0);
  t.setDate(t.getDate() + 4 - (t.getDay() || 7));
  var jahrStart = new Date(t.getFullYear(), 0, 1);
  return Math.ceil((((t - jahrStart) / 86400000) + 1) / 7);
}

// ──────────────────────────────────────────────
// 10. Blätter sichern (beim POST automatisch)
// ──────────────────────────────────────────────
function sichereBlaetter(ss) {
  if (!ss.getSheetByName("Bestellungen")) initBestellungen(ss);
  if (!ss.getSheetByName(WOCHE_NAME))     initWochenuebersichtKueche(ss);
  if (!ss.getSheetByName(MONAT_NAME))     initMonatsuebersichtBuchhaltung(ss);
}

// ──────────────────────────────────────────────
// 10b. PIN-Validierung – prüft Apartment+PIN gegen Bewohner-Tab
//      Rückgabe: Name bei Erfolg, leerer String bei Fehlschlag
// ──────────────────────────────────────────────
function validierePin(ss, apt, pin) {
  var sh = ss.getSheetByName("Bewohner");
  if (!sh || sh.getLastRow() < 2) return "";
  var rows = sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues();
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i][0]).trim() === apt &&
        String(rows[i][1]).trim() === pin) {
      return String(rows[i][2] || "").trim() || apt;
    }
  }
  return "";
}

// ──────────────────────────────────────────────
// 11a. GET-Endpunkt – Bewohnerliste ODER Bestellungen für Apartment
// ──────────────────────────────────────────────
function doGet(e) {
  try {
    var ss     = SpreadsheetApp.getActiveSpreadsheet();
    sichereBlaetter(ss);
    var action = String((e.parameter && e.parameter.action) || "").trim();
    var apt    = String((e.parameter && e.parameter.apartment) || "").trim();
    var pin    = String((e.parameter && e.parameter.pin) || "").trim();

    // ── Action: getBewohner → Apartment+Name (KEIN PIN!) zurückgeben
    if (action === "getBewohner") {
      var bwSheet = ss.getSheetByName("Bewohner");
      if (!bwSheet || bwSheet.getLastRow() < 2) return json_ok([]);
      var bwRows = bwSheet.getRange(2, 1, bwSheet.getLastRow() - 1, 3).getValues();
      var bwOut = [];
      bwRows.forEach(function(r) {
        var a = String(r[0] || "").trim();
        var n = String(r[2] || "").trim();
        if (a && n) bwOut.push({ apartment: a, name: n });
      });
      return json_ok(bwOut);
    }

    // ── Standard: Bestellungen für Apartment abrufen (PIN-Validierung)
    if (!apt) return json_err("Kein Apartment angegeben");
    if (!pin) return json_err("Kein PIN angegeben");
    var name = validierePin(ss, apt, pin);
    if (!name) return json_err("Ungültiges Apartment oder falscher PIN");

    var sheet = ss.getSheetByName("Bestellungen");
    var rows  = sheet.getDataRange().getValues();
    var out   = [];
    if (rows.length > 1) {
      for (var i = 1; i < rows.length; i++) {
        if (String(rows[i][4]).trim() !== apt) continue;
        out.push({
          datum:      rows[i][0],
          wochentag:  rows[i][1],
          kw:         rows[i][2],
          jahr:       rows[i][3],
          apartment:  rows[i][4],
          name:       rows[i][5],
          menueTyp:   rows[i][6],
          menueName:  rows[i][7],
          vorspeise:  rows[i][8],
          nachspeise: rows[i][9],
          preis:      rows[i][10]
        });
      }
    }
    return ContentService
      .createTextOutput(JSON.stringify({ok: true, name: name, data: out}))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return json_err(err.message);
  }
}

// ──────────────────────────────────────────────
// 11b. POST-Endpunkt – schreibt Bestellung + aktualisiert Übersichten
// ──────────────────────────────────────────────
function doPost(e) {
  try {
    var data  = JSON.parse(e.postData.contents);
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    sichereBlaetter(ss);

    // ── Sicherheit: PIN+Apartment validieren bevor etwas gespeichert wird
    var apt = String(data.apartment || "").trim();
    var pin = String(data.pin       || "").trim();
    if (!apt || !pin) return json_err("Apartment und PIN erforderlich");
    var name = validierePin(ss, apt, pin);
    if (!name) return json_err("Ungültiges Apartment oder falscher PIN");
    // Wichtig: Der Name kommt IMMER aus der Bewohner-Tabelle (server-side),
    // niemals aus data.name – damit kann der Client nichts Falsches schicken.

    var sheet = ss.getSheetByName("Bestellungen");
    var kw    = Number(data.kw);
    var jahr  = Number(data.jahr);
    var typ   = data.aenderung ? "Änderung" : "Erstbestellung";
    var ts    = Utilities.formatDate(new Date(), "Europe/Berlin", "dd.MM.yyyy HH:mm");

    // Upsert: alte Zeilen für diese Bewohner-KW löschen
    var alle = sheet.getDataRange().getValues();
    for (var i = alle.length - 1; i >= 1; i--) {
      if (String(alle[i][4]).trim() === apt &&
          Number(alle[i][2]) === kw &&
          Number(alle[i][3]) === jahr) {
        sheet.deleteRow(i + 1);
      }
    }

    var tage = data.tage || [];
    for (var t = 0; t < tage.length; t++) {
      var tag = tage[t];
      if (!tag.wahl) continue;
      sheet.appendRow([
        tag.datum, tag.wochentag, kw, jahr,
        apt, name,
        tag.menueTyp, tag.menueName,
        tag.vorspeise, tag.nachspeise,
        tag.preis, ts, typ
      ]);
    }

    // Übersichten direkt mit aktualisieren
    aktualisiereUebersichten();

    return json_ok({zeilen: tage.length, typ: typ});
  } catch (err) {
    return json_err(err.message);
  }
}

// ──────────────────────────────────────────────
// 12. JSON-Helpers
// ──────────────────────────────────────────────
function json_ok(data) {
  return ContentService
    .createTextOutput(JSON.stringify({ok: true, data: data}))
    .setMimeType(ContentService.MimeType.JSON);
}
function json_err(msg) {
  return ContentService
    .createTextOutput(JSON.stringify({ok: false, error: msg}))
    .setMimeType(ContentService.MimeType.JSON);
}

// ──────────────────────────────────────────────
// 13. Reparatur Bestellungen-Daten in Text-Format
// ──────────────────────────────────────────────
function repariereBestellungen() {
  var n = repariereBestellungenStill(SpreadsheetApp.getActiveSpreadsheet());
  if (n < 0)      SpreadsheetApp.getUi().alert("Tab 'Bestellungen' fehlt.");
  else if (n===0) SpreadsheetApp.getUi().alert("Keine Daten zum Reparieren.");
  else            SpreadsheetApp.getUi().alert("✅ " + n + " Zeilen repariert.");
}

function repariereBestellungenStill(ss) {
  var sheet = ss.getSheetByName("Bestellungen");
  if (!sheet) return -1;
  var lr = sheet.getLastRow();
  if (lr < 2) return 0;

  var rngA = sheet.getRange(2, 1, lr - 1, 1);
  var neuA = rngA.getValues().map(function(r){
    var v = r[0];
    if (v instanceof Date) return [Utilities.formatDate(v, "Europe/Berlin", "dd.MM.yyyy")];
    return [String(v).trim()];
  });
  sheet.getRange("A:A").setNumberFormat("@");
  rngA.setValues(neuA);

  var rngE = sheet.getRange(2, 5, lr - 1, 1);
  var neuE = rngE.getValues().map(function(r){ return [String(r[0]).trim()]; });
  sheet.getRange("E:E").setNumberFormat("@");
  rngE.setValues(neuE);

  return lr - 1;
}

// ──────────────────────────────────────────────
// 14. Manuelle Aktualisierung (im Editor ausführbar)
// ──────────────────────────────────────────────
function aktualisiereJetzt() {
  aktualisiereUebersichten();
  SpreadsheetApp.getUi().alert("✅ Übersichten wurden aktualisiert.");
}

// ──────────────────────────────────────────────
// 15. Reparatur: Namen in Bestellungen aus Bewohner-Tab auffüllen
//     (falls fehlerhafte Apartment-Nummer als Name gespeichert wurde)
// ──────────────────────────────────────────────
function repariereNamen() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Bestellungen");
  var bw    = ss.getSheetByName("Bewohner");
  if (!sheet || !bw) {
    SpreadsheetApp.getUi().alert("Tab 'Bestellungen' oder 'Bewohner' fehlt.");
    return;
  }
  var lr = sheet.getLastRow();
  if (lr < 2) { SpreadsheetApp.getUi().alert("Keine Bestellungen vorhanden."); return; }

  // Apartment → Name Mapping aus Bewohner-Tab
  var bwRows = bw.getRange(2, 1, bw.getLastRow() - 1, 3).getValues();
  var aptZuName = {};
  bwRows.forEach(function(r){
    var apt  = String(r[0] || "").trim();
    var name = String(r[2] || "").trim();
    if (apt && name) aptZuName[apt] = name;
  });

  var rng = sheet.getRange(2, 5, lr - 1, 2); // Spalten E (Apartment) + F (Name)
  var werte = rng.getValues();
  var korrigiert = 0;
  werte.forEach(function(r, i){
    var apt = String(r[0]).trim();
    var name = String(r[1]).trim();
    var richtigerName = aptZuName[apt];
    if (richtigerName && name !== richtigerName) {
      werte[i][1] = richtigerName;
      korrigiert++;
    }
  });
  if (korrigiert > 0) rng.setValues(werte);
  SpreadsheetApp.getUi().alert("✅ " + korrigiert + " Namen aus Bewohner-Tab korrigiert.");
}

// ──────────────────────────────────────────────
// 15. Test-Funktion
// ──────────────────────────────────────────────
function testEintrag() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  sichereBlaetter(ss);
  var sheet = ss.getSheetByName("Bestellungen");
  sheet.appendRow([
    "12.05.2026","Dienstag",20,2026,
    "12","Max Mustermann",
    "Hauptmenü","Curry aus Blumenkohl mit Reis",
    "Ja","Ja",7.50,
    Utilities.formatDate(new Date(),"Europe/Berlin","dd.MM.yyyy HH:mm"),
    "Erstbestellung"
  ]);
  aktualisiereUebersichten();
  SpreadsheetApp.getUi().alert("✅ Test-Zeile + Übersichten aktualisiert!");
}
