// =====================================================================================
// interner-bereich.js
// =====================================================================================
// Vollständige Implementierung für "Interner Bereich"
// - Lädt Liste "Interne_Seiten" per JSOM (Title, URL, Prio)
// - Verwendet das vom Nutzer gelieferte Search-Payload (mit P5619770115) für /_api/search/postquery
// - Filtert Listeneinträge auf Basis der von der Suche zurückgegebenen Sites (SPSiteUrl)
// - Lädt Site-Logos per /_api/web?$select=SiteLogoUrl (fetch, credentials: 'include') und cached Ergebnisse
// - Robust: ermittelt RequestDigest via __REQUESTDIGEST, _spPageContextInfo.formDigestValue oder /_api/contextinfo
// - Verbesserte, ausgiebig debuggende extractSPSiteUrlsFromSearchResult, die mehrere Shapes der Search-Antwort unterstützt
// =====================================================================================

// ==========================
// Konfiguration
// ==========================
var INTERNAL_LIST_NAME   = "Interne_Seiten";
var INTERNAL_TITLE_FIELD = "Title";
var INTERNAL_URL_FIELD   = "URL";
var INTERNAL_PRIO_FIELD  = "Prio";     // Nummernfeld für Reihenfolge
var SEARCH_ROW_LIMIT     = 500;        // max Ergebnisse von search/postquery (anpassen bei Bedarf)

// ==========================
// Search-Payload: exakter Payload (P5619770115)
// ==========================
var SEARCH_PAYLOAD = {
  "request": {
    "__metadata": { "type": "Microsoft.Office.Server.Search.REST.SearchRequest" },
    "ClientType": "PnPModernSearch",
    "Properties": {
      "results": [
        { "Name": "EnableDynamicGroups", "Value": { "BoolVal": true, "QueryPropertyValueTypeIndex": 3 } },
        { "Name": "EnableMultiGeoSearch", "Value": { "BoolVal": true, "QueryPropertyValueTypeIndex": 3 } }
      ]
    },
    "Querytext": "*",
    "TimeZoneId": 4,
    "EnableQueryRules": true,
    "QueryTemplate": "{searchTerms} contentclass=STS_Site RefinableString118=\"*P5619770115*\" WebTemplate:STS RefinableString112=1",
    "RowLimit": 10,
    "SelectProperties": {
      "results": [
        "Title","Path","Created","Filename","SiteLogo","PreviewUrl","PictureThumbnailURL",
        "ServerRedirectedPreviewURL","ServerRedirectedURL","HitHighlightedSummary","FileType",
        "contentclass","ServerRedirectedEmbedURL","ParentLink","DefaultEncodingURL",
        "owstaxidmetadataalltagsinfo","Author","AuthorOWSUSER","SPSiteUrl","SiteTitle","IsContainer",
        "IsListItem","HtmlFileType","SiteId","WebId","UniqueID","OriginalPath","FileExtension",
        "IsDocument","NormSiteID","NormListID","NormUniqueID","AssignedTo","RefinableDate10",
        "RefinableDate11","RefinableDate13","RefinableDate17","RefinableDate18","RefinableDate19",
        "IsAllDayEvent","RefinableString101","RefinableString102","RefinableString103",
        "RefinableString105","RefinableString106","RefinableString107","RefinableString109",
        "RefinableString110","RefinableString112","RefinableString113","RefinableString114",
        "RefinableString115","RefinableString116","RefinableString117","RefinableString118",
        "RefinableString119","RefinableString120","RefinableString121","RefinableString123",
        "RefinableString124"
      ]
    },
    "TrimDuplicates": false,
    "SortList": { "results": [ { "Property": "LastModifiedTime", "Direction": 1 } ] },
    "Culture": 1031,
    "Refiners": "RefinableString106,RefinableString115,RefinableString118,Created(discretize=manual/2024-12-09T05:43:28.156Z/2025-09-09T04:43:28.156Z/2025-11-09T05:43:28.156Z/2025-12-02T05:43:28.156Z/2025-12-08T05:43:28.156Z),LastModifiedTime(discretize=manual/2024-12-09T05:43:28.156Z/2025-09-09T04:43:28.156Z/2025-11-09T05:43:28.156Z/2025-12-02T05:43:28.156Z/2025-12-08T05:43:28.156Z)",
    "HitHighlightedProperties": { "results": [] },
    "RefinementFilters": { "results": [] },
    "ReorderingRules": { "results": [] }
  }
};

// ==========================
// DOM-Referenzen
// ==========================
var openBtn        = document.getElementById("openInternalAreaBtn");
var overlay        = document.getElementById("internalModalOverlay");
var modal          = document.getElementById("internalModal");
var closeBtn       = document.getElementById("internalModalCloseBtn");
var listContainer  = document.getElementById("internalListContainer");
var stateContainer = document.getElementById("internalStateContainer");
var searchInput    = document.getElementById("internalSearchInput");
var countBadge     = document.getElementById("internalCountBadge");

// ==========================
// Caches & Status
// ==========================
var internalItemsCache      = []; // Alle Listeneinträge (inkl. prio)
var internalAccessibleItems = []; // Gefiltert nach Search-Ergebnissen (SPSiteUrl)
var logoCache               = {}; // key = webUrl (no trailing slash) -> logoUrl|null
var accessCheckInProgress   = false;

// ==========================
// Modal-Steuerung
// ==========================
function openInternalModal() {
  if (overlay) overlay.classList.add("visible");
  if (modal) modal.classList.add("visible");
  if (overlay) overlay.setAttribute("aria-hidden", "false");

  if (!internalItemsCache || internalItemsCache.length === 0) {
    loadInternalList();
  } else {
    renderList();
  }
}
function closeInternalModal() {
  if (overlay) overlay.classList.remove("visible");
  if (modal) modal.classList.remove("visible");
  if (overlay) overlay.setAttribute("aria-hidden", "true");
}
if (openBtn) openBtn.addEventListener("click", openInternalModal);
if (closeBtn) closeBtn.addEventListener("click", closeInternalModal);
if (overlay) overlay.addEventListener("click", function(e){ if (e.target === overlay) closeInternalModal(); });
document.addEventListener("keydown", function(e){ if (e.key === "Escape" && overlay && overlay.classList.contains("visible")) closeInternalModal(); });

// ==========================
// Status Helpers
// ==========================
function setState(messageHtml, cssClass) {
  if (!stateContainer) return;
  stateContainer.style.display = "block";
  stateContainer.className = "internal-state " + (cssClass || "");
  stateContainer.innerHTML = messageHtml;
}
function clearState() {
  if (!stateContainer) return;
  stateContainer.style.display = "none";
  stateContainer.innerHTML = "";
  stateContainer.className = "internal-state";
}

// ==========================
// JSOM: Liste laden (Title, URL, Prio)
// ==========================
function loadInternalList() {
  setState('<span><div class="internal-spinner"></div> Laden der internen Seiten …</span>');

  if (typeof SP === "undefined" || !SP.ClientContext) {
    setState("Fehler: Die SharePoint-JSOM-Bibliotheken (sp.js) sind nicht verfügbar.", "internal-error");
    return;
  }

  try {
    var context = SP.ClientContext.get_current();
    var web = context.get_web();
    var list = web.get_lists().getByTitle(INTERNAL_LIST_NAME);

    var camlQuery = new SP.CamlQuery();
    camlQuery.set_viewXml(
      "<View>" +
        "<Query></Query>" +
        "<ViewFields>" +
          "<FieldRef Name='" + INTERNAL_TITLE_FIELD + "' />" +
          "<FieldRef Name='" + INTERNAL_URL_FIELD + "' />" +
          "<FieldRef Name='" + INTERNAL_PRIO_FIELD + "' />" +
        "</ViewFields>" +
      "</View>"
    );

    var items = list.getItems(camlQuery);
    context.load(items);

    context.executeQueryAsync(function(){
      var enumerator = items.getEnumerator();
      var results = [];
      while (enumerator.moveNext()) {
        var item = enumerator.get_current();
        var title = item.get_item(INTERNAL_TITLE_FIELD) || "";
        var urlField = item.get_item(INTERNAL_URL_FIELD);
        var url = "";
        var urlDescription = "";
        if (urlField && urlField.get_url) {
          url = urlField.get_url();
          urlDescription = urlField.get_description ? urlField.get_description() : title;
        }

        var prioRaw = item.get_item(INTERNAL_PRIO_FIELD);
        var prioVal = (prioRaw === null || prioRaw === undefined || prioRaw === "") ? 999999 : Number(prioRaw);
        if (isNaN(prioVal)) prioVal = 999999;

        if (url) {
          results.push({
            title: title,
            url: url,
            description: urlDescription || title,
            prio: prioVal,
            siteLogoUrl: null,
            webUrl: null
          });
        }
      }

      // sortiere nach Prio
      results.sort(function(a,b){ return a.prio - b.prio; });

      internalItemsCache = results;

      if (!results || results.length === 0) {
        clearState();
        setState("Es wurden keine Einträge in der Liste <strong>" + INTERNAL_LIST_NAME + "</strong> gefunden.", "internal-empty");
        internalAccessibleItems = [];
        renderList();
        return;
      }

      // benutze die Search-API mit dem gelieferten (HAR) payload
      checkAccessWithSearch(results);
    }, function(sender, args){
      setState("Fehler beim Laden der Liste <strong>" + INTERNAL_LIST_NAME + "</strong>: " + args.get_message(), "internal-error");
    });
  } catch (e) {
    setState("Unbekannter Fehler beim Laden der Liste: " + (e && e.message ? e.message : e), "internal-error");
  }
}

// ==========================
// RequestDigest: modern (fetch) mit Fallbacks
// ==========================
function getRequestDigest() {
  return new Promise(function(resolve, reject){
    try {
      var digestElem = document.getElementById("__REQUESTDIGEST");
      if (digestElem && digestElem.value) { resolve(digestElem.value); return; }

      if (window._spPageContextInfo && _spPageContextInfo.formDigestValue) { resolve(_spPageContextInfo.formDigestValue); return; }

      var baseUrl = (window._spPageContextInfo && _spPageContextInfo.webAbsoluteUrl) ? _spPageContextInfo.webAbsoluteUrl.replace(/\/$/, "") : (window.location.protocol + '//' + window.location.host);
      var ctxUrl = baseUrl + "/_api/contextinfo";

      fetch(ctxUrl, {
        method: 'POST',
        headers: { 'Accept': 'application/json;odata=verbose', 'Content-Type': 'application/json;odata=verbose;charset=utf-8' },
        credentials: 'include'
      }).then(function(resp){
        if (!resp.ok) throw new Error("ContextInfo antwortete mit HTTP " + resp.status);
        return resp.json();
      }).then(function(parsed){
        var digest = parsed && parsed.d && parsed.d.GetContextWebInformation && parsed.d.GetContextWebInformation.FormDigestValue;
        if (digest) resolve(digest);
        else reject(new Error("Kein FormDigest im Kontextinfo-Antwortkörper gefunden."));
      }).catch(function(err){
        reject(err);
      });
    } catch (ex) {
      reject(ex);
    }
  });
}

// ==========================
// postSearchQuery: verwendet das bereitgestellte SEARCH_PAYLOAD (HAR)
// ==========================
function postSearchQuery(payload) {
  return new Promise(function(resolve, reject){
    var baseUrl = (window._spPageContextInfo && _spPageContextInfo.webAbsoluteUrl) ? _spPageContextInfo.webAbsoluteUrl.replace(/\/$/, "") : (window.location.protocol + '//' + window.location.host);
    var url = baseUrl + "/_api/search/postquery";

    getRequestDigest().then(function(digest){
      fetch(url, {
        method: 'POST',
        headers: {
          'Accept': 'application/json;odata=nometadata',
          'Content-Type': 'application/json;odata=verbose;charset=utf-8',
          'X-RequestDigest': digest
        },
        credentials: 'include',
        body: JSON.stringify(payload)
      }).then(function(resp){
        if (!resp.ok) {
          return resp.text().then(function(txt){
            var serverMsg = "";
            try {
              var parsed = JSON.parse(txt);
              serverMsg = (parsed && parsed.error && parsed.error.message && parsed.error.message.value) || parsed.Message || "";
            } catch(e){}
            throw new Error("Search-API antwortete mit HTTP " + resp.status + (serverMsg ? (": " + serverMsg) : ""));
          });
        }
        return resp.json();
      }).then(function(json){
        resolve(json);
      }).catch(function(err){
        reject(err);
      });
    }).catch(function(err){
      reject(new Error("RequestDigest konnte nicht ermittelt werden: " + (err && err.message ? err.message : err)));
    });
  });
}

// =====================================================================================
// Erweiterte, robuste Funktion: extractSPSiteUrlsFromSearchResult
// - Unterstützt mehrere JSON-Shapes:
//   * RelevantResults.Table.Rows.results (klassisch nometadata/hierarchie-objekt)
//   * RelevantResults.Table.Rows (Array)
//   * RelevantResults.Table (falls Rows wrapped differently)
// - Sammelt umfangreiche Debug-Informationen in window.searchDebug
// - Gibt ein Array eindeutiger SPSiteUrl-Strings (ohne trailing slash) zurück
// =====================================================================================
function extractSPSiteUrlsFromSearchResult(searchResult) {
  var debug = {
    timestamp: new Date().toISOString(),
    ok: false,
    summary: {
      hasPrimaryQueryResult: false,
      hasRelevantResults: false,
      rowsLocated: 0,
      extractedUrls: 0,
      uniqueUrls: 0
    },
    shapesTried: [],
    errors: [],
    warnings: [],
    rowsInfo: [], // per-row info
    rawSnapshot: null
  };

  try {
    debug.rawSnapshot = searchResult;

    if (!searchResult || !searchResult.PrimaryQueryResult) {
      debug.warnings.push("PrimaryQueryResult fehlt.");
      window.searchDebug = debug;
      console.warn("extractSPSiteUrlsFromSearchResult - PrimaryQueryResult fehlt", debug);
      return [];
    }
    debug.summary.hasPrimaryQueryResult = true;

    var relevant = searchResult.PrimaryQueryResult.RelevantResults;
    if (!relevant) {
      debug.warnings.push("RelevantResults fehlt.");
      window.searchDebug = debug;
      console.warn("extractSPSiteUrlsFromSearchResult - RelevantResults fehlt", debug);
      return [];
    }
    debug.summary.hasRelevantResults = true;

    var table = relevant.Table;
    if (!table) {
      debug.warnings.push("RelevantResults.Table fehlt.");
      window.searchDebug = debug;
      console.warn("extractSPSiteUrlsFromSearchResult - RelevantResults.Table fehlt", debug);
      return [];
    }

    // Try multiple path shapes to locate rows array
    var rows = [];

    // Shape 1: Table.Rows.results (common for some odata shapes)
    if (table.Rows && table.Rows.results && Array.isArray(table.Rows.results)) {
      debug.shapesTried.push("Table.Rows.results (array)");
      rows = table.Rows.results;
    }
    // Shape 2: Table.Rows is array directly (as in your provided payload)
    else if (Array.isArray(table.Rows)) {
      debug.shapesTried.push("Table.Rows (array)");
      rows = table.Rows;
    }
    // Shape 3: Table has property 'Rows' but it's an object with different casing or nested
    else if (table.Rows && typeof table.Rows === "object") {
      // try to discover any array inside table.Rows
      debug.shapesTried.push("Table.Rows (object) - searching for array children");
      for (var p in table.Rows) {
        if (table.Rows.hasOwnProperty(p) && Array.isArray(table.Rows[p])) {
          debug.shapesTried.push("Found array at Table.Rows." + p);
          rows = table.Rows[p];
          break;
        }
      }
    }
    // Shape 4: Some responses include RelevantResults.Table (Rows property deeper)
    if (!rows || rows.length === 0) {
      // fallback: try RelevantResults.Table.Rows.results (case-insensitive)
      try {
        for (var key in relevant) {
          if (!relevant.hasOwnProperty(key)) continue;
          var val = relevant[key];
          if (val && typeof val === "object") {
            // look for Rows property inside
            if (val.Rows && val.Rows.results && Array.isArray(val.Rows.results)) {
              debug.shapesTried.push("Fallback: relevant." + key + ".Rows.results");
              rows = val.Rows.results;
              break;
            } else if (Array.isArray(val.Rows)) {
              debug.shapesTried.push("Fallback: relevant." + key + ".Rows (array)");
              rows = val.Rows;
              break;
            }
          }
        }
      } catch (e) {
        debug.errors.push("Fehler beim Fallback-Scan: " + (e && e.message ? e.message : String(e)));
      }
    }

    debug.summary.rowsLocated = (rows && rows.length) ? rows.length : 0;

    if (!rows || rows.length === 0) {
      debug.warnings.push("Keine Rows gefunden - shapesTried: " + JSON.stringify(debug.shapesTried));
      window.searchDebug = debug;
      console.groupCollapsed("extractSPSiteUrlsFromSearchResult - no rows found");
      console.debug("debug:", debug);
      console.debug("searchResult.PrimaryQueryResult.RelevantResults:", searchResult.PrimaryQueryResult.RelevantResults);
      console.groupEnd();
      return [];
    }

    // iterate rows and extract SPSiteUrl (supporting both cell shapes: array of Cells.results or Cells array)
    var foundUrls = [];
    for (var r = 0; r < rows.length; r++) {
      var row = rows[r];
      var rowInfo = { index: r, cellsCount: 0, foundSPSiteUrl: null, keys: [] };
      try {
        var cells = [];
        // common case: row.Cells.results is array
        if (row.Cells && Array.isArray(row.Cells.results)) {
          cells = row.Cells.results;
        } else if (row.Cells && Array.isArray(row.Cells)) {
          cells = row.Cells;
        } else if (Array.isArray(row)) {
          // in case row itself is an array of cells
          cells = row;
        } else if (row && typeof row === "object") {
          // try to locate array inside row (Cells or other keys)
          for (var k in row) {
            if (!row.hasOwnProperty(k)) continue;
            if (Array.isArray(row[k])) {
              // prefer Cells if present
              if (k.toLowerCase() === "cells" || k.toLowerCase() === "cells") {
                cells = row[k];
                break;
              } else if (!cells.length) {
                cells = row[k];
              }
            } else if (row[k] && row[k].results && Array.isArray(row[k].results)) {
              cells = row[k].results;
              break;
            }
          }
        }

        rowInfo.cellsCount = Array.isArray(cells) ? cells.length : 0;

        for (var c = 0; c < cells.length; c++) {
          var cell = cells[c];
          var key = (cell && typeof cell.Key !== "undefined") ? cell.Key : (cell && cell.k) ? cell.k : null;
          if (key) rowInfo.keys.push(key);
          // Some shapes may have {Key, Value} or {key, value} or direct map
          var cellKey = cell && (cell.Key || cell.key || cell.k);
          var cellVal = cell && (cell.Value || cell.value || cell.v);
          if (!cellKey && typeof cell === "object") {
            // if cell is a plain object with single key-value representing SPSiteUrl
            // iterate its props
            for (var ck in cell) {
              if (!cell.hasOwnProperty(ck)) continue;
              // skip internal ValueType etc
              if (ck === "Key" || ck === "Value" || ck === "key" || ck === "value") continue;
            }
          }
          if (cellKey === "SPSiteUrl" && cellVal) {
            var siteUrl = String(cellVal).replace(/\/$/, "");
            rowInfo.foundSPSiteUrl = siteUrl;
            foundUrls.push(siteUrl);
            debug.summary.extractedUrls++;
            break; // one SPSiteUrl per row is enough
          }
          // sometimes Keys are under "k" and value under "v"
          if (!cellKey && cell && typeof cell === "object" && ("Key" in cell || "key" in cell) ) {
            var ck2 = cell.Key || cell.key;
            var cv2 = cell.Value || cell.value;
            if (ck2 === "SPSiteUrl" && cv2) {
              var su = String(cv2).replace(/\/$/, "");
              rowInfo.foundSPSiteUrl = su;
              foundUrls.push(su);
              debug.summary.extractedUrls++;
              break;
            }
          }
        }

        if (!rowInfo.foundSPSiteUrl) {
          rowInfo.note = "Kein SPSiteUrl in dieser Zeile gefunden.";
        }
      } catch (rowErr) {
        var errMsg = rowErr && rowErr.message ? rowErr.message : String(rowErr);
        rowInfo.error = errMsg;
        debug.errors.push("Fehler beim Verarbeiten von Row " + r + ": " + errMsg);
      }
      debug.rowsInfo.push(rowInfo);
    }

    // deduplicate case-insensitive but preserve original casing of first occurrence
    var unique = [];
    var seen = {};
    for (var i = 0; i < foundUrls.length; i++) {
      var u = foundUrls[i];
      var keyLower = (u || "").toLowerCase();
      if (!seen[keyLower]) {
        seen[keyLower] = true;
        unique.push(u);
      }
    }

    debug.summary.uniqueUrls = unique.length;
    debug.summary.rowsLocated = rows.length;
    debug.ok = true;

    // expose debug globally
    window.searchDebug = debug;
    window.lastExtractedSPSiteUrls = unique;

    // console output (grouped)
    console.groupCollapsed("extractSPSiteUrlsFromSearchResult debug - " + debug.timestamp);
    console.debug("shapesTried:", debug.shapesTried);
    console.debug("summary:", debug.summary);
    if (debug.errors.length) console.error("errors:", debug.errors);
    if (debug.warnings.length) console.warn("warnings:", debug.warnings);
    console.table(debug.rowsInfo.map(function(r){ return { index: r.index, cells: r.cellsCount, found: r.foundSPSiteUrl || "(none)", note: r.note || "", error: r.error || "" }; }));
    console.debug("unique SPSiteUrls:", unique);
    // print small sample of raw RelevantResults for inspection (avoid huge dump)
    try {
      console.debug("RelevantResults sample:", (searchResult.PrimaryQueryResult && searchResult.PrimaryQueryResult.RelevantResults && (searchResult.PrimaryQueryResult.RelevantResults.Table || searchResult.PrimaryQueryResult.RelevantResults)));
    } catch (ignore){}
    console.groupEnd();

    return unique;
  } catch (e) {
    var em = (e && e.message) ? e.message : String(e);
    debug.errors.push("Unhandled exception: " + em);
    debug.ok = false;
    window.searchDebug = debug;
    console.error("extractSPSiteUrlsFromSearchResult unexpected error", e, debug);
    return [];
  }
}

// ==========================
// Zugriffsermittlung mittels Search (verwendet SEARCH_PAYLOAD)
// ==========================
function checkAccessWithSearch(items) {
  accessCheckInProgress = true;
  internalAccessibleItems = [];

  setState('<span><div class="internal-spinner"></div> Berechtigte Sites werden über die Suche ermittelt …</span>');

  // distinct webUrls aus Items ableiten
  var webUrlSet = {};
  items.forEach(function(it){
    var webUrl = getWebUrlFromPageUrl(it.url);
    it.webUrl = webUrl;
    if (webUrl) webUrlSet[webUrl.replace(/\/$/, "").toLowerCase()] = webUrl.replace(/\/$/, "");
  });

  var allWebUrls = Object.keys(webUrlSet);
  if (!allWebUrls || allWebUrls.length === 0) {
    clearState();
    setState("Aus den URLs konnten keine Sites ermittelt werden.", "internal-empty");
    renderList();
    return;
  }

  // Verwende das exakte SEARCH_PAYLOAD (HAR)
  postSearchQuery(SEARCH_PAYLOAD).then(function(searchResult){
    var accessibleWebs = extractSPSiteUrlsFromSearchResult(searchResult);
    if (!accessibleWebs || accessibleWebs.length === 0) {
      clearState();
      setState("Die Suche hat keine Sites zurückgegeben, auf die du zugreifen kannst.", "internal-empty");
      renderList();
      return;
    }

    var accessibleSet = {};
    accessibleWebs.forEach(function(u){ accessibleSet[u.toLowerCase()] = true; });

    // Filter Items nach gefundenen Sites
    var filtered = items.filter(function(it){
      if (!it.webUrl) return false;
      return !!accessibleSet[it.webUrl.replace(/\/$/, "").toLowerCase()];
    });

    internalAccessibleItems = filtered;

    // Logos nachladen (optional)
    return loadLogosForAccessibleItems(filtered);
  }).then(function(){
    accessCheckInProgress = false;
    clearState();

    if (!internalAccessibleItems || internalAccessibleItems.length === 0) {
      setState("Es konnten keine internen Seiten gefunden werden, auf die du Zugriff hast.", "internal-empty");
    }

    internalAccessibleItems.sort(function(a,b){ return a.prio - b.prio; });
    renderList();
  }).catch(function(err){
    accessCheckInProgress = false;
    clearState();
    setState("Fehler bei der Zugriffsermittlung über die Suche: " + (err && err.message ? err.message : err), "internal-error");
  });
}

// ==========================
// postSearchQuery & RequestDigest (unchanged)
// ==========================
function postSearchQuery(payload) {
  return new Promise(function(resolve, reject){
    var baseUrl = (window._spPageContextInfo && _spPageContextInfo.webAbsoluteUrl) ? _spPageContextInfo.webAbsoluteUrl.replace(/\/$/, "") : (window.location.protocol + '//' + window.location.host);
    var url = baseUrl + "/_api/search/postquery";

    getRequestDigest().then(function(digest){
      fetch(url, {
        method: 'POST',
        headers: {
          'Accept': 'application/json;odata=nometadata',
          'Content-Type': 'application/json;odata=verbose;charset=utf-8',
          'X-RequestDigest': digest
        },
        credentials: 'include',
        body: JSON.stringify(payload)
      }).then(function(resp){
        if (!resp.ok) {
          return resp.text().then(function(txt){
            var serverMsg = "";
            try {
              var parsed = JSON.parse(txt);
              serverMsg = (parsed && parsed.error && parsed.error.message && parsed.error.message.value) || parsed.Message || "";
            } catch(e){}
            throw new Error("Search-API antwortete mit HTTP " + resp.status + (serverMsg ? (": " + serverMsg) : ""));
          });
        }
        return resp.json();
      }).then(function(json){
        resolve(json);
      }).catch(function(err){
        reject(err);
      });
    }).catch(function(err){
      reject(new Error("RequestDigest konnte nicht ermittelt werden: " + (err && err.message ? err.message : err)));
    });
  });
}

function getRequestDigest() {
  return new Promise(function(resolve, reject){
    try {
      var digestElem = document.getElementById("__REQUESTDIGEST");
      if (digestElem && digestElem.value) { resolve(digestElem.value); return; }

      if (window._spPageContextInfo && _spPageContextInfo.formDigestValue) { resolve(_spPageContextInfo.formDigestValue); return; }

      var baseUrl = (window._spPageContextInfo && _spPageContextInfo.webAbsoluteUrl) ? _spPageContextInfo.webAbsoluteUrl.replace(/\/$/, "") : (window.location.protocol + '//' + window.location.host);
      var ctxUrl = baseUrl + "/_api/contextinfo";

      fetch(ctxUrl, {
        method: 'POST',
        headers: { 'Accept': 'application/json;odata=verbose', 'Content-Type': 'application/json;odata=verbose;charset=utf-8' },
        credentials: 'include'
      }).then(function(resp){
        if (!resp.ok) throw new Error("ContextInfo antwortete mit HTTP " + resp.status);
        return resp.json();
      }).then(function(parsed){
        var digest = parsed && parsed.d && parsed.d.GetContextWebInformation && parsed.d.GetContextWebInformation.FormDigestValue;
        if (digest) resolve(digest);
        else reject(new Error("Kein FormDigest im Kontextinfo-Antwortkörper gefunden."));
      }).catch(function(err){
        reject(err);
      });
    } catch (ex) {
      reject(ex);
    }
  });
}

// ==========================
// Logos laden (fetch) mit Cache
// ==========================
function loadLogosForAccessibleItems(items) {
  var webSet = {};
  items.forEach(function(it){
    if (it.webUrl) webSet[it.webUrl.replace(/\/$/, "").toLowerCase()] = it.webUrl.replace(/\/$/, "");
  });
  var distinctWebs = Object.keys(webSet).map(function(k){ return webSet[k]; });

  var promises = distinctWebs.map(function(webUrl){
    return fetchWebLogoUrl(webUrl);
  });

  return Promise.all(promises).then(function(){
    items.forEach(function(it){
      if (it.webUrl) {
        var key = it.webUrl.replace(/\/$/, "");
        it.siteLogoUrl = logoCache.hasOwnProperty(key) ? logoCache[key] : null;
      }
    });
  });
}

function fetchWebLogoUrl(webUrl) {
  if (!webUrl) return Promise.resolve(null);
  var key = webUrl.replace(/\/$/, "");
  if (logoCache.hasOwnProperty(key)) return Promise.resolve(logoCache[key]);

  var restUrl = key + "/_api/web?$select=SiteLogoUrl";
  return fetch(restUrl, { method: 'GET', headers: { 'Accept': 'application/json;odata=nometadata' }, credentials: 'include' })
    .then(function(resp){
      if (!resp.ok) { logoCache[key] = null; return null; }
      return resp.json().then(function(json){
        var logo = json && json.SiteLogoUrl ? json.SiteLogoUrl : null;
        logoCache[key] = logo;
        return logo;
      }).catch(function(){ logoCache[key] = null; return null; });
    }).catch(function(){ logoCache[key] = null; return null; });
}

// ==========================
// Utility: Web-URL aus Seiten-URL ableiten
// ==========================
function getWebUrlFromPageUrl(absUrl) {
  try {
    var u = new URL(absUrl);
    var path = u.pathname;
    var pageFolders = ["/SitePages/", "/Seiten/", "/Pages/", "/_layouts/15/"];
    for (var i=0;i<pageFolders.length;i++){
      var marker = pageFolders[i];
      var idx = path.toLowerCase().indexOf(marker.toLowerCase());
      if (idx !== -1) {
        var webPath = path.substring(0, idx);
        if (!webPath) webPath = "/";
        return u.protocol + "//" + u.host + webPath;
      }
    }
    var lastSlash = path.lastIndexOf("/");
    if (lastSlash > 0) {
      var webPath2 = path.substring(0, lastSlash);
      return u.protocol + "//" + u.host + webPath2;
    }
    return u.protocol + "//" + u.host;
  } catch(e){
    return null;
  }
}

// ==========================
// Darstellung / Rendering
// ==========================
function stripDomain(url) {
  if (!url) return "";
  try {
    var u = new URL(url);
    return u.pathname + (u.search || "") + (u.hash || "");
  } catch(e){
    return url.replace(/^https?:\/\/[^/]+/i, "");
  }
}
function getInitials(text) {
  var trimmed = (text || "").trim();
  if (!trimmed) return "I";
  var parts = trimmed.split(/\s+/);
  if (parts.length === 1) return parts[0].substring(0,2).toUpperCase();
  return (parts[0].charAt(0) + parts[1].charAt(0)).toUpperCase();
}

function renderList() {
  if (!listContainer) return;
  listContainer.innerHTML = "";

  var sourceItems = internalAccessibleItems || [];
  var filterText = (searchInput && searchInput.value ? searchInput.value : "").toLowerCase().trim();

  var filtered = sourceItems.filter(function(it){
    if (!filterText) return true;
    var hay = (it.title + " " + it.url + " " + it.description).toLowerCase();
    return hay.indexOf(filterText) !== -1;
  });

  if (countBadge) countBadge.textContent = filtered.length + " Eintrag" + (filtered.length === 1 ? "" : "e");

  if (!filtered || filtered.length === 0) {
    listContainer.innerHTML = '<div class="internal-state internal-empty" style="padding:14px"><span>Keine Treffer für deine Suche.</span></div>';
    return;
  }

  filtered.forEach(function(it){
    var row = document.createElement("div"); row.className = "internal-row";

    var logoCell = document.createElement("div"); logoCell.className = "internal-logo-cell";
    if (it.siteLogoUrl) {
      var img = document.createElement("img"); img.src = it.siteLogoUrl; img.alt = "Site-Logo"; img.className = "internal-logo-img"; logoCell.appendChild(img);
    } else {
      var logo = document.createElement("div"); logo.className = "internal-logo"; logo.textContent = getInitials(it.title || it.description || "I"); logoCell.appendChild(logo);
    }

    var titleCell = document.createElement("div"); titleCell.className = "internal-title-cell";
    var titleEl = document.createElement("div"); titleEl.className = "internal-title"; titleEl.textContent = it.title || "(Ohne Titel)";
    var urlEl = document.createElement("div"); urlEl.className = "internal-url";
    if (it.url) { var a = document.createElement("a"); a.href = it.url; a.target = "_blank"; a.rel = "noopener noreferrer"; a.textContent = stripDomain(it.url); urlEl.appendChild(a); } else { urlEl.textContent = "(Keine URL angegeben)"; }
    titleCell.appendChild(titleEl); titleCell.appendChild(urlEl);

    var infoCell = document.createElement("div"); infoCell.className = "internal-info-cell";
    var infoTag = document.createElement("div"); infoTag.className = "internal-info-tag";
    var dot = document.createElement("div"); dot.className = "internal-info-dot";
    var infoText = document.createElement("span"); infoText.textContent = "Interne Seite";
    infoTag.appendChild(dot); infoTag.appendChild(infoText); infoCell.appendChild(infoTag);

    var actionCell = document.createElement("div");
    var btn = document.createElement("button"); btn.type = "button"; btn.className = "internal-open-btn";
    btn.innerHTML = '<span class="internal-open-btn-icon"></span><span>Öffnen</span>';
    btn.addEventListener("click", function(){ if (it.url) window.open(it.url, "_blank", "noopener"); });
    actionCell.appendChild(btn);

    row.appendChild(logoCell);
    row.appendChild(titleCell);
    row.appendChild(infoCell);
    row.appendChild(actionCell);

    listContainer.appendChild(row);
  });
}

if (searchInput) {
  searchInput.addEventListener("input", function(){ renderList(); });
}

// =====================================================================================
// Ende Datei
// =====================================================================================