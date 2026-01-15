// Globale Konfig für SPMetaHarvest
// Hier feste Ersetzungen (alt -> neu) hinterlegen, die IMMER angewendet werden sollen.
window.SPHarvestConfig = window.SPHarvestConfig || {};

window.SPHarvestConfig.overrides = {
  "https://zsandstbw.ecm.bundeswehr.org/portale/P5603398863": "https://ustgber.ecm.bundeswehr.org/portale/P5619770115",

  // weitere Beispiele:
  // "/portale/P5603398863": "/portale/P5619770115",
  // "http://altes-internes-hostname": "http://neues-internes-hostname"
};

// Optional: Die Overrides auch direkt in die Textarea der Abgleich-Seite füllen.
// Auf true setzen, falls gewünscht.
window.SPHarvestConfig.prefillOverridesTextarea = true;