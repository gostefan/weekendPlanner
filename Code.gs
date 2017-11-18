FormularId      = "18hGQ2RvJmBbgjjGbbCJJeD_qxdWedKcywuQgyAqhRmI"
NamenDropdownId = "281467942"
INT_HUGE        = 99999999999
Freitag         = 5
Samstag         = 6
Sonntag         = 0

StatistikBlattName = "Assistenzärzte"
StatistikAnfangsSpalte = 6
StatistikNameSpalte = 2
StatistikProzentSpalte = 4
StatistikAbteilungSpalte = 5
StatistikTag = 0
StatistikNacht = 1
StatistikWE = 2
StatistikErsteZeile = 3

StatAbt = 3
StatProzent = 4

VerhindertBlattName = "Verhindert"
VerhindertErsteZeile = 2
VerhindertSpalteName = 2
VerhindertSpalteDatum = 3

AusgabeSpalteDatum = 1
AusgabeSpaltePerson = 2
AusgabeErsteZeile = 2

function test() {
}

function planen() {
  var monat = promptNumber("Monat als Zahl eingeben")
  var jahr  = promptNumber( "Jahr als Zahl eingeben")
  var sheet = tabFinden(monat, jahr)
  if (sheet == undefined) {
    Browser.msgBox("Das Tab zum Planungsmonat konnte nicht gefunden werden. Breche ab...")
    return
  }
  monatPlanen(sheet.getDataRange(), monat, jahr) 
}

function tabFinden(monat, jahr) {
  var ssa = SpreadsheetApp.getActive();
  var sheetName = monatJahrString(monat, jahr)
  var sheet = ssa.getSheetByName(sheetName);
  return sheet
}

function monatPlanen(area, monat, jahr) {
  var statistik  = statistikLesen (monat, jahr)
  var verhindert = verhindertLesen(monat, jahr)
  var nTage = tageImMonat(monat, jahr)
  var resultat = personenAussuchen(area, statistik, verhindert, monat, jahr, nTage)
  if (resultat != null) {
    schichtplanSchreiben(area, resultat, nTage)
    statistikSchreiben(statistik, monat, jahr)
    Browser.msgBox("Die Monatsplanung ist abgeschlossen")
  } else
    Browser.msgBox("Es konnte kein Schichtplan gefunden werden, der alle Bedingungen erfüllt.")
}

//------------------- Personen auswaehlen --------------------

function personenAussuchen(area, statistik, verhindert, monat, jahr, nTage) {
  var letzteNacht = letzterMonatNachtPerson(monat, jahr)
  var ausschliessen = []
  if (letzteNacht != null)
    aussliessen.push(letzteNacht)
  return schichtFuellen(area, statistik, ausschliessen, "", verhindert, 1, monat, jahr, nTage, StatistikTag)
}

function schichtFuellen(area, statistik, ausschliessen, abtAusschliessen, verhindert, tag, monat, jahr, nTage, tagNacht) {
  var statistikTag = istWochenende(tag, monat, jahr, tagNacht) ? StatistikWE : tagNacht
  if (tag > nTage) // recursion abortion
    return []

  // TODO: handle impossible selection for downstream
  var nextAbtAusschliessen = ausschliessen
  var person = area.getCell(tag + AusgabeErsteZeile - 1, AusgabeSpaltePerson + tagNacht).getValue();
  if (person == "") {
    var ausschluss = ausschlussSammeln(statistik, ausschliessen, abtAusschliessen, verhindert[new Date(tag, monat, jahr)]);
    person = personAuswaehlen(statistik, ausschluss, statistikTag)
  }
  if (person == undefined)
    return null;
  
  var resultat = schichtFuellenUndErrorHandling(area, statistik, ausschliessen, person, verhindert, tag, monat, jahr, nTage, tagNacht, statistikTag)
  
  if (resultat[tag] == undefined)
    resultat[tag] = []
  resultat[tag][tagNacht] = person
  return resultat
}

function schichtFuellenUndErrorHandling(area, statistik, ausschliessen, person, verhindert, tag, monat, jahr, nTage, tagNacht, statsTag) {
  if (statistik[person] != undefined)
    statistik[person][statsTag]++
  var naechsterTag = tag + (tagNacht == StatistikNacht ? 1 : 0)
  var naechsterTagNacht = tagNacht == StatistikNacht ? StatistikTag : StatistikNacht
  var naechsteAbtAusgeschl = ausschliessen.length > 0 ? ausschliessen[0] : ""
  var resultat = schichtFuellen(area, statistik, [person], naechsteAbtAusgeschl, verhindert, naechsterTag, monat, jahr, nTage, naechsterTagNacht)
  if (resultat == null) {
    if (statistik[person] != undefined)
      statistik[person][statsTag]--
    return null
  }
  return resultat
}

function istWochenende(tag, monat, jahr, tagNacht) {
  var datum = new Date(jahr, monat - 1, tag)
  var saSo = datum.getDay() == Samstag || datum.getDay() == Sonntag
  var freitagAbend = datum.getDay() == Freitag && tagNacht == StatistikNacht
  return saSo || freitagAbend
}

function ausschlussSammeln(statistik, ausschliessen, abtAusschliessen, verhinderte) {
  for (var verhinderter in verhinderte)
    ausschliessen.push(verhinderter)
  if (abtAusschliessen != "" && statistik[abtAusschliessen] != undefined) {
    var abteilung = statistik[abtAusschliessen][StatAbt]
    for (var eintrag in statistik)
      if (statistik[eintrag][StatAbt] == abteilung && eintrag != abtAusschliessen)
        ausschliessen.push(eintrag)
  }
  return ausschliessen
}

function personAuswaehlen(statistik, ausschliessen, tagNacht) {
  var moegliche = statistikVorbereiten(statistik, ausschliessen, tagNacht)
  
  for (n in moegliche) {
    var wenigsten = moegliche[n]
    var index = Math.floor(Math.random() * wenigsten.length)
    return wenigsten[index]
  }
}

function statistikVorbereiten(statistik, ausschliessen, tagNacht) {
  var namen = extractNames(statistik)
  var resultat = []
  for (var idx in namen) {
    var name = namen[idx]
    if (!wertIn(name, ausschliessen)) {
      var anzahlGeleistet = Math.round(statistik[name][tagNacht] / statistik[name][StatProzent])
      if (isNaN(anzahlGeleistet))
          anzahlGeleistet = 0
      if (resultat[anzahlGeleistet] == null)
        resultat[anzahlGeleistet] = []
      resultat[anzahlGeleistet].push(namen[idx])
    }
  }
  return resultat
}

function extractNames(statistik) {
  var result = []
  var minimum = INT_HUGE
  for (var name in statistik) {
    result.push(name)
  }
  return result
}

//---------------- Daten schreiben ------------------

function schichtplanSchreiben(area, resultat, nTage) {
  for (var i = 1; i <= nTage; i++) {
    for (var tagNacht = 0; tagNacht <= 1; tagNacht++) {
      var person = resultat[i][tagNacht]
      var cell = area.getCell(i + AusgabeErsteZeile - 1, AusgabeSpaltePerson + tagNacht)
      cell.setValue(person)
    }
  }
}

function statistikSchreiben(statistik, monat, jahr) {
  var spreadsheet = SpreadsheetApp.getActive()
  var sheet = spreadsheet.getSheetByName(StatistikBlattName)
  var range = neuerStatistikBereich(sheet)
  var startSpalte = range.getWidth() - 2
  range.getCell(1, startSpalte).setValue(monatJahrString(monat, jahr))
  range.getCell(2, startSpalte + StatistikTag  ).setValue("Tag")
  range.getCell(2, startSpalte + StatistikNacht).setValue("Nacht")
  range.getCell(2, startSpalte + StatistikWE   ).setValue("Wochenende")
  for (var i = StatistikErsteZeile; i <= range.getHeight(); i++) {
    var name = range.getCell(i, StatistikNameSpalte).getValue()
    if (name != "" && statistik[name] != undefined) {
      for (var tagNacht = 0; tagNacht <= 2; tagNacht++) {
        var wert = statistik[name][tagNacht]
        range.getCell(i, startSpalte + tagNacht).setValue(wert)
      }
    }
  }
}

function neuerStatistikBereich(sheet) {
  var oldRange = sheet.getDataRange()
  var newWidth = oldRange.getWidth() + 3
  return sheet.getRange(1, 1, oldRange.getHeight(), newWidth)
}

//---------------- Daten lesen -------------------

function letzterMonatNachtPerson(monat, jahr) {
  var spreadsheet = SpreadsheetApp.getActive()
  var lastMonth = letzterMonatString(monat, jahr)
  var sheet = spreadsheet.getSheetByName(lastMonth)
  if (sheet != null) {
    var range = sheet.getDataRange()
    return range.getCell(range.getHeight(), AusgabeSpaltePersonNacht).getValue()
  } else
    return null
}

function statistikLesen(monat, jahr) {
  var spreadsheet = SpreadsheetApp.getActive()
  var sheet = spreadsheet.getSheetByName(StatistikBlattName)
  var activeRange = sheet.getDataRange()
  var letzterMonat = letzterMonatString(monat, jahr)
  var aktiveSpalte = findeMonatsSpalte(activeRange, letzterMonat);
  if (aktiveSpalte == -1)
    aktiveSpalte = StatistikAnfangsSpalte;
  return statistikZeilenLesen(activeRange, aktiveSpalte)
}

function letzterMonatString(monat, jahr) {
  return monat == 1 ? monatJahrString(12, jahr - 1) : monatJahrString(monat - 1, jahr)
}

function findeMonatsSpalte(range, monatString) {
  for (var i = 1; i < range.getWidth(); i++)
    if (range.getCell(1, i).getValue() == monatString)
      return i;
  return -1;
}

function statistikZeilenLesen(activeRange, aktiveSpalte) {
  var result = {}
  for (i = StatistikErsteZeile; i <= activeRange.getHeight(); i++) {
    var name   = activeRange.getCell(i, StatistikNameSpalte).getValue()
    var abteilung = activeRange.getCell(i, StatistikAbteilungSpalte).getValue()
    var prozent = activeRange.getCell(i, StatistikProzentSpalte).getValue()
    var nTag   = statistikEintragLesen(activeRange, aktiveSpalte, i, StatistikTag)
    var nNacht = statistikEintragLesen(activeRange, aktiveSpalte, i, StatistikNacht)
    var nWE    = statistikEintragLesen(activeRange, aktiveSpalte, i, StatistikWE)
    result[name] = [ nTag, nNacht, nWE, abteilung, prozent ]
  }
  return result
}

function statistikEintragLesen(activeRange, aktiveSpalte, i, offset) {
  var resultat = activeRange.getCell(i, aktiveSpalte + offset).getValue()
  if (resultat == "")
    resultat = activeRange.getCell(i, StatistikAnfangsSpalte + offset).getValue()
  if (resultat == "" || isNaN(resultat))
    resultat = 0
  return resultat
}

function verhindertLesen(monat, jahr) {
  var spreadsheet = SpreadsheetApp.getActive()
  var sheet = spreadsheet.getSheetByName(VerhindertBlattName)
  var activeRange = sheet.getDataRange()
  
  var width = activeRange.getWidth()
  var result = {}
  for (i = VerhindertErsteZeile; i <= activeRange.getHeight(); i++) {
    var date = new Date(activeRange.getCell(i, VerhindertSpalteDatum).getValue());
    var name = activeRange.getCell(i, VerhindertSpalteName).getValue()
    if (date.getMonth() == monat - 1 && date.getYear() == jahr) {
      if (result[date] != undefined)
        result[date].push(name)
      else
        result[date] = [name]
    }
  }
  return result;
}

//---------------- Monats Tab kreieren --------------------

function monatsTab() {
  var monat = promptNumber("Monat als Zahl eingeben")
  var jahr  = promptNumber( "Jahr als Zahl eingeben")
  var sheet = tabKreieren(monat, jahr)
  datenFuellen(sheet, monat, jahr, tageImMonat(monat, jahr))
  monatTabDesign(sheet)
  Browser.msgBox("Der Monats-Tab wurde erstellt.")
}

function tabKreieren(monat, jahr) {
  var spreadsheet = SpreadsheetApp.getActive()
  var numSheets = spreadsheet.getNumSheets()
  var sheetName = monatJahrString(monat, jahr)
  var newSheet = spreadsheet.insertSheet(sheetName, numSheets + 1)
  newSheet.activate()
  return newSheet
}

function datenFuellen(sheet, monat, jahr, nTage) {
  sheet.getRange(1,1).setValue("Datum")
  for (var i = 1; i <= nTage; i++) {
    var zelle = sheet.getRange(i + AusgabeErsteZeile - 1, 1)
    zelle.setValue(datumFormatieren(i, monat, jahr))
  }
}

function monatTabDesign(sheet) {
  sheet.setFrozenColumns(1)
  sheet.setFrozenRows(1)
  sheet.getRange(1, 2).setValue("Tag")
  sheet.getRange(1, 3).setValue("Nacht")
  sheet.getRange(1, 1, 1, 3).setFontWeight("bold")
}

//---------------- Formular aktualisieren --------------------

function mitarbeiterAktualisieren() {
  var form = FormApp.openById(FormularId);
  var namesList = form.getItemById(NamenDropdownId).asListItem();

  var ssa = SpreadsheetApp.getActive();
  var namesSheet = ssa.getSheetByName(StatistikBlattName);
  var namesValues = namesSheet.getRange(StatistikErsteZeile, StatistikNameSpalte, namesSheet.getMaxRows() - 1).getValues();
  var names = [];
  for(var i = 0; i < namesValues.length; i++)    
    if(namesValues[i][0] != "")
      names[i] = namesValues[i][0];
  namesList.setChoiceValues(names);
  
  Browser.msgBox("Das Formular wurde aktualisiert.")
}

//----------------- Generelle funktionen ---------------------

function datumFormatieren(tag, monat, jahr) {
  return tag.toString() + "." + monat.toString() + "." + jahr.toString() 
}

function monatJahrString(monat, jahr) {
  return monat.toString() + "." + jahr.toString()
}

function wertIn(wert, liste) {
  for (idx in liste)
    if (wert == liste[idx])
      return true;
  return false;
}

function promptNumber(text) {
  var value = null
  while (isNaN(parseInt(value)) || !isFinite(value))
    value = Browser.inputBox(text)
  return parseInt(value)
}

function tageImMonat(monat, jahr) {
  for (var i = 28; i <= 31; i++) {
    var date = new Date(jahr, monat - 1, i)
    if (date.getMonth() != monat - 1 || date.getDate() != i)
      return i - 1
  }
  return 31
}
