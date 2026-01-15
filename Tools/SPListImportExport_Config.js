/**
 * SPListImportExport_Config.js
 *
 * Globale Konfigurationsdatei für das SPListImportExport-Tool.
 * Legt Listen fest, die beim Export ignoriert werden sollen, sowie
 * optionale Regex-Patterns zum Ignorieren nach Namen.
 *
 * Diese Datei muss das globale Objekt `window.SPListImportExport_Config`
 * definieren. Beispiel-Aufruf im Browser:
 *   console.log(window.SPListImportExport_Config);
 *
 * Hinweis:
 * - Die Permissions-API-Base ist im Tool fest auf
 *   "https://webapps01.ecm.bundeswehr.org/shareextension" gesetzt und
 *   muss hier nicht konfiguriert werden.
 * - Patterns sind reguläre Ausdrücke (als String). Sie werden mit dem
 *   Flag 'i' (case-insensitive) verwendet.
 */

window.SPListImportExport_Config = {
  // Liste von Listentiteln, die beim Export komplett übersprungen werden.
  // Beispiel:
  // ignoreLists: ["DoNotExport", "Archiv_Alt"]
  ignoreLists: [
    "*Anforderungsliste URM",
    "*Hörsaalliste URM",
    "BwMitP*",
    "BwSPCR*",
    "BwTask Id Provider",
    "Bw-Tasks",
    "appdata",
    "appfiles",
    "CustomAppSettings",
    "Designkatalog",
    "Durchkomponierte Looks",
    "Formatbibliothek",
    "Formularvorlagen",
    "Gestaltungsvorlagenkatalog",
    "K2 Settings",
    "K2Pages",
    "Liste der Projektrichtlinienelemente",
    "Listenvorlagenkatalog",
    "Lösungskatalog",
    "Suchkonfigurationsliste",
    "TaxonomyHiddenList",
    "Webpartkatalog",
    "wfpub",
    "Workflowverlauf",
    "Websiteobjekte",
    "WebSiteSeiten"
  ],

  // Liste von Regex-Strings; wenn ein Listenname zu einem Pattern passt,
  // wird die Liste beim Export übersprungen.
  // Beispiel: ignorePatterns: ["^Test_", ".*_Temp$"]
  ignorePatterns: [
    // "^Test_",
    // ".*_Temp$"
  ],

  // Optional: zusätzliche Kontrollflags (werden vom Hauptskript gelesen, falls vorhanden)
  options: {
    // Wenn true, werden standardmäßig auch ausgeblendete Listen (Hidden=true)
    // in den Export aufgenommen. Das Hauptskript berücksichtigt Hidden bereits,
    // diese Option dient nur zur Dokumentation / zukünftigen Verwendung.
    includeHiddenLists: true,

    // Wenn true, schreibt das Tool beim Export die verwendete Config ins Export-File.
    embedConfigInExport: true
  }
};