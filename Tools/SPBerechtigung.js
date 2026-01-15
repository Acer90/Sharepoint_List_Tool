/**
 * SPBerechtigung.js - Vollständige, bereinigte Implementierung
 *
 * Features:
 * - Mehrere Site-URLs (UI zum Hinzufügen/Entfernen)
 * - Baumansicht (Web -> Listen/Libs -> Folders -> Items) mit Toggle
 * - Actions rechts in der Node-Zeile (Öffnen, Reset)
 * - Permission-Einträge mit Aktionen (Kopieren, Entfernen) rechts ausgerichtet
 * - Übersichtstabelle mit Checkboxen + Bulk-Löschen
 * - PDF-Export des Baumes (grafisch via html2canvas + jsPDF, Fallback Text-PDF)
 * - Expand/Collapse-Button funktioniert korrekt
 *
 * Hinweis:
 * - Diese Datei erwartet jQuery auf der Seite.
 * - Für grafischen PDF-Export werden html2canvas und jsPDF empfohlen.
 */

var S6PermTool = (function () {
  // --------------------------
  // Konfiguration
  // --------------------------
  var API_BASE = "https://webapps01.ecm.bundeswehr.org/shareextension";
  var ENDPOINTS = {
    removeList: API_BASE + "/lists/removeusers",
    removeItem: API_BASE + "/listitems/removeusers",
    assignList: API_BASE + "/lists/assignusers",
    assignItem: API_BASE + "/listitems/assignusers",
    resetItem: API_BASE + "/listitems/resetroleinheritance",
  };

  var EXCLUDED_STRINGS = [
    "sharepoint\\system",
    "nt authority\\authenticated users",
    "gestaltende",
    "lesende",
    "mitwirkende mit löschen",
    "erstellende",
    "ändernde",
    "mitwirkende ohne löschen",
    "bearbeitende",
    "genehmigende",
    "[svc] all_sp_sscowner",
    "[svc] all_sp_fscowner",
  ];
  var IGNORE_ROLE = "Beschränkter Zugriff";

  // --------------------------
  // Zustand
  // --------------------------
  var currentSiteUrl = "";
  var currentUserLogin = "";
  var currentPermHigh = "0";
  var currentPermLow = "0";

  // key: userId@siteUrl => { user: MemberObject, contexts: [...] }
  var activeUsersMap = {};
  // key: pathLower@siteUrl => $li
  var nodeLookup = {};

  var treeExpanded = false;

  // --------------------------
  // Hilfsfunktionen
  // --------------------------
  function log(msg, pct) {
    var $t = $("#s6LoadingText");
    if ($t.length) $t.text(msg);
    if (pct !== undefined) $("#s6ProgressBar").css("width", pct + "%");
  }

  function getJson(url) {
    return $.ajax({
      url: url,
      type: "GET",
      headers: { Accept: "application/json;odata=verbose" },
    });
  }

  function normalizeUrl(url) {
    if (!url) return "";
    return url.replace(/\/+$/, "");
  }

  // Normalisiert Site-Eingabe: trimmen, Schema ergänzen, Trailing Slash entfernen
  function normalizeSiteInput(url) {
    if (!url) return "";
    var t = String(url).trim();
    if (!/^https?:\/\//i.test(t)) t = "https://" + t.replace(/^\/+/, "");
    return t.replace(/\/+$/, "");
  }

  function getAbsoluteUrl(serverRelativePath) {
    try {
      var u = new URL(currentSiteUrl);
      var origin = u.origin;
      var rel = serverRelativePath || "";
      if (!rel.startsWith("/")) rel = "/" + rel;
      return origin + rel;
    } catch (e) {
      var rel2 = serverRelativePath || "";
      if (!rel2.startsWith("/")) rel2 = "/" + rel2;
      var m = currentSiteUrl.match(/^https?:\/\/[^\/]+/);
      var origin2 = m ? m[0] : "";
      return origin2 + rel2;
    }
  }

  function getCleanLogin(loginName) {
    if (!loginName) return "";
    var parts = loginName.split("|");
    var last = parts[parts.length - 1];
    return last.indexOf("\\") > -1 ? last.split("\\")[1] : last;
  }

  function isExcludedMember(member) {
    var login = (member.LoginName || "").toLowerCase();
    var title = (member.Title || "").toLowerCase();
    return EXCLUDED_STRINGS.some(function (ex) {
      ex = ex.toLowerCase();
      return login.indexOf(ex) > -1 || title.indexOf(ex) > -1;
    });
  }

  function makeCustomApiCall(url, payload) {
    return $.ajax({
      url: url,
      type: "POST",
      data: JSON.stringify(payload),
      headers: {
        Accept: "application/json;odata=verbose",
        "Content-Type": "application/json;odata=verbose",
        spappidentifier: "App4Gw.ShareExtension",
        spappweburl: currentSiteUrl,
        sphosturl: currentSiteUrl,
        sppermissionhigh: String(currentPermHigh),
        sppermissionlow: String(currentPermLow),
        spuserloginname: currentUserLogin,
      },
    });
  }

  function siteUserKey(userId) {
    return userId + "@" + currentSiteUrl;
  }
  function nodeKey(path) {
    return (path || "").toLowerCase() + "@" + currentSiteUrl;
  }

  // --------------------------
  // PDF Export (grafisch via html2canvas + jsPDF, sonst Text-Fallback)
  // --------------------------
  function exportTreeToPdf() {
    try {
      var $container = $(".s6-tree-container").first();
      if ($container.length === 0) {
        alert("Kein Baum vorhanden.");
        return;
      }

      var prevState = treeExpanded;
      expandAll();

      // grafischer Export
      if (typeof html2canvas !== "undefined" && (window.jspdf && window.jspdf.jsPDF || window.jsPDF)) {
        html2canvas($container.get(0), { backgroundColor: "#ffffff", useCORS: true, scale: 2 })
          .then(function (canvas) {
            try {
              var imgData = canvas.toDataURL("image/png");
              var doc;
              if (window.jspdf && window.jspdf.jsPDF) doc = new window.jspdf.jsPDF({ unit: "pt", format: "a4" });
              else doc = new window.jsPDF();

              var pageWidth = doc.internal.pageSize.getWidth();
              var pageHeight = doc.internal.pageSize.getHeight();
              var margin = 20;

              var ratio = canvas.width / (pageWidth - margin * 2);
              var imgWidth = canvas.width / ratio;
              var imgHeight = canvas.height / ratio;

              if (imgHeight <= pageHeight - margin * 2) {
                doc.addImage(imgData, "PNG", margin, margin, imgWidth, imgHeight);
              } else {
                // slice vertical
                var sliceHeight = Math.floor((pageHeight - margin * 2) * ratio);
                var canvasPage = document.createElement("canvas");
                canvasPage.width = canvas.width;
                canvasPage.height = sliceHeight;
                var ctx = canvasPage.getContext("2d");

                var rendered = 0;
                var first = true;
                while (rendered < canvas.height) {
                  var h = Math.min(sliceHeight, canvas.height - rendered);
                  ctx.clearRect(0, 0, canvasPage.width, canvasPage.height);
                  ctx.drawImage(canvas, 0, rendered, canvasPage.width, h, 0, 0, canvasPage.width, h);
                  var sliceData = canvasPage.toDataURL("image/png");
                  var sliceImgHeight = h / ratio;

                  if (first) {
                    doc.addImage(sliceData, "PNG", margin, margin, imgWidth, sliceImgHeight);
                    first = false;
                  } else {
                    doc.addPage();
                    doc.addImage(sliceData, "PNG", margin, margin, imgWidth, sliceImgHeight);
                  }
                  rendered += sliceHeight;
                }
              }

              var filename = "Berechtigungsbaum_" + new Date().toISOString().slice(0, 19).replace(/[:T]/g, "_") + ".pdf";
              doc.save(filename);
            } catch (err) {
              console.error("PDF Export Fehler:", err);
              alert("Fehler beim PDF-Export: " + (err && err.message ? err.message : err));
            } finally {
              if (!prevState) collapseAll();
            }
          })
          .catch(function (err) {
            console.error("html2canvas Fehler:", err);
            exportTreeAsTextPdf();
            if (!prevState) collapseAll();
          });
        return;
      }

      // Fallback: Text-PDF
      exportTreeAsTextPdf();
      if (!prevState) collapseAll();
    } catch (ex) {
      console.error("Export Fehler", ex);
      alert("Fehler beim PDF-Export: " + (ex && ex.message ? ex.message : ex));
    }
  }

  function exportTreeAsTextPdf() {
    try {
      var doc;
      if (window.jspdf && window.jspdf.jsPDF) doc = new window.jspdf.jsPDF({ unit: "pt", format: "a4" });
      else if (window.jsPDF) doc = new window.jsPDF();
      else {
        alert("jsPDF nicht gefunden. Bitte lade jsPDF oder verwende grafischen Export.");
        return;
      }

      var lines = [];
      lines.push("Berechtigungs-Baum von " + (currentSiteUrl || window.location.hostname));
      lines.push("Export Datum: " + new Date().toLocaleString());
      lines.push("----------------------------------------");

      function traverseLiToLines($li, depth) {
        var icon = $li.find("> .s6-node-row > .s6-icon").first().text() || "";
        var title = $li.find("> .s6-node-row > .s6-node-title").first().text() || "";
        var combined = (icon ? icon + " " : "") + title;
        combined = combined.replace(/\s+/g, " ").trim();
        lines.push(indent(depth) + combined);

        $li.children("ul").children("li").each(function () {
          var $child = $(this);
          if ($child.children(".s6-node-row").length) traverseLiToLines($child, depth + 1);
          else {
            var text = $child.text().replace(/\s+/g, " ").trim();
            if (text) lines.push(indent(depth + 1) + text);
          }
        });
      }
      function indent(n) {
        return new Array(n + 1).join("    ");
      }

      $("#s6TreeRoot").children("li").each(function () {
        traverseLiToLines($(this), 0);
      });

      var marginLeft = 40;
      var marginTop = 50;
      var pageHeight = doc.internal.pageSize.getHeight();
      var pageWidth = doc.internal.pageSize.getWidth();
      var lineHeight = 12;
      var y = marginTop;
      var maxWidth = pageWidth - marginLeft - 40;

      doc.setFontSize(10);
      for (var i = 0; i < lines.length; i++) {
        var split = doc.splitTextToSize(lines[i], maxWidth);
        for (var j = 0; j < split.length; j++) {
          if (y + lineHeight > pageHeight - 40) {
            doc.addPage();
            y = marginTop;
          }
          doc.text(split[j], marginLeft, y);
          y += lineHeight;
        }
      }

      var filename = "Berechtigungsbaum_" + new Date().toISOString().slice(0, 19).replace(/[:T]/g, "_") + ".pdf";
      doc.save(filename);
    } catch (ex) {
      console.error("Text-PDF Export Fehler", ex);
      alert("Fehler beim Text-PDF-Export: " + (ex && ex.message ? ex.message : ex));
    }
  }

  // expose globally (in case inline handlers exist)
  window.exportTreeToPdf = exportTreeToPdf;

  // --------------------------
  // UI Initialisierung & Bindings
  // --------------------------
  $(document).ready(function () {
    // Bind main buttons
    $(document).on("click", "#btnStartLoad", startBatchAnalysis);
    $(document).on("click", "#btnFetchSubsites", fetchSubsitesForFirstUrl);
    $(document).on("click", "#s6AddUrlBtn", function (e) {
      e.preventDefault();
      addUrlRow();
    });

    // Export button
    $(document).on("click", "#s6ExportPdfBtn", function (e) {
      e.preventDefault();
      exportTreeToPdf();
    });

    // Expand/Collapse button (delegated)
    $(document).on("click", "#s6ExpandCollapseBtn", function (e) {
      e.preventDefault();
      toggleExpandCollapseAll();
    });

    // Node row click toggles
    $(document).on("click", ".s6-node-row", function (e) {
      // ignore clicks on buttons inside the row
      if ($(e.target).closest("button").length) return;
      var $li = $(this).closest("li");
      var $ul = $li.children("ul");
      if ($ul.length) {
        $ul.slideToggle(120);
        var $tog = $(this).children(".s6-toggle");
        if ($tog.length) $tog.text($tog.text() === "▶" ? "▼" : "▶");
      }
    });

    // Toggle icon only
    $(document).on("click", ".s6-toggle", function (e) {
      e.stopPropagation();
      var $li = $(this).closest("li");
      var $ul = $li.children("ul");
      if ($ul.length) {
        $ul.slideToggle(120);
        var $tog = $(this);
        $tog.text($tog.text() === "▶" ? "▼" : "▶");
      }
    });

    // Select all checkbox in user table
    $(document).on("change", "#s6UserTableSelectAll", function () {
      var checked = $(this).is(":checked");
      $("#s6UserTable tbody input.user-select").prop("checked", checked);
    });

    // Bulk delete selected users
    $(document).on("click", "#s6DeleteSelectedBtn", function () {
      var selected = $("#s6UserTable tbody input.user-select:checked")
        .map(function () {
          return $(this).data("userid");
        })
        .get();
      if (!selected || selected.length === 0) {
        alert("Bitte mindestens einen Benutzer auswählen.");
        return;
      }
      if (!confirm("Ausgewählte (" + selected.length + ") Benutzer löschen?")) return;

      var promises = selected.map(function (uid) {
        return removeUserFromAllFoundContexts(uid, null, true);
      });
      Promise.all(promises)
        .then(function () {
          alert("Ausgewählte Benutzer wurden entfernt (sofern API-Aufrufe erfolgreich).");
          startBatchAnalysis();
        })
        .catch(function () {
          alert("Bei einigen Löschungen ist ein Fehler aufgetreten. Bitte prüfen.");
          startBatchAnalysis();
        });
    });

    // Ensure UI elements exist / have classes
    ensureUrlInputsUi();
    ensureExpandCollapseButton();

    // prefer page context url
    if (typeof _spPageContextInfo !== "undefined" && _spPageContextInfo.webAbsoluteUrl) {
      $("#s6UrlInput").val(_spPageContextInfo.webAbsoluteUrl);
    }
  });

  // --------------------------
  // UI: Mehrfach-URL Eingaben
  // --------------------------
  function ensureUrlInputsUi() {
    var $baseInput = $("#s6UrlInput");
    if ($baseInput.length === 0) {
      // create base input if missing
      var $wrapper = $("<div class='s6-header'></div>");
      var $row = $("<div class='s6-url-row'></div>");
      $baseInput = $("<input id='s6UrlInput' type='text' placeholder='https://...' />");
      $row.append($baseInput);
      $wrapper.append($row);
      $("#s6-tool-wrapper").prepend($wrapper);
    }

    // label row
    var $labelRow = $("#s6UrlLabelRow");
    if ($labelRow.length === 0) {
      $labelRow = $("<div id='s6UrlLabelRow'></div>");
      var $label = $("<label id='s6UrlLabel' for='s6UrlInput'>Site URL:</label>");
      var $addBtn = $("<button id='s6AddUrlBtn' class='s6-btn s6-btn-sm'>URL hinzufügen</button>");
      var $fetchBtn = $("#btnFetchSubsites");
      if ($fetchBtn.length === 0) $fetchBtn = $("<button id='btnFetchSubsites' class='s6-btn s6-btn-sm'>Untergeordnete Seiten abrufen</button>");

      $labelRow.append($label).append($addBtn).append($fetchBtn);
      $labelRow.insertBefore($("#s6UrlList"));
    }
  }

  function addUrlRow() {
    var $wrapper = $("#s6UrlList");
    if ($wrapper.length === 0) {
      ensureUrlInputsUi();
      $wrapper = $("#s6UrlList");
    }
    var inputId = "s6UrlInput_" + Date.now();
    var $row = $("<div class='s6-url-row'></div>");
    var $input = $("<input type='text' class='s6-url-input' />").attr("id", inputId).attr("placeholder", "https://...");
    var $remove = $("<button class='s6-btn s6-btn-danger s6-btn-sm'>Entfernen</button>");
    $remove.on("click", function () {
      $row.remove();
    });
    $row.append($input).append($remove);
    $wrapper.append($row);
  }

  // --------------------------
  // Expand/Collapse Button
  // --------------------------
  function ensureExpandCollapseButton() {
    var $header = $("#s6TreeHeader");
    if ($header.length === 0) {
      $header = $("<div id='s6TreeHeader' class='s6-tree-header'></div>");
      var $title = $("<h3>Berechtigungs-Baum</h3>");
      var $exportBtn = $("<button id='s6ExportPdfBtn' class='s6-btn s6-btn-sm'>Als PDF exportieren</button>");
      var $toggleBtn = $("<button id='s6ExpandCollapseBtn' class='s6-btn s6-btn-sm'>Baum aufklappen</button>");

      $header.append($title).append($exportBtn).append($toggleBtn);
      var $root = $("#s6TreeRoot");
      if ($root.length) $root.before($header);
      else $("#s6-tool-wrapper").append($header);
    } else {
      // ensure the button exists
      if ($header.find("#s6ExpandCollapseBtn").length === 0) {
        var $toggleBtn2 = $("<button id='s6ExpandCollapseBtn' class='s6-btn s6-btn-sm'>Baum aufklappen</button>");
        $header.append($toggleBtn2);
      }
      if ($header.find("#s6ExportPdfBtn").length === 0) {
        var $exportBtn2 = $("<button id='s6ExportPdfBtn' class='s6-btn s6-btn-sm'>Als PDF exportieren</button>");
        $header.append($exportBtn2);
      }
    }
    updateExpandCollapseButtonState(treeExpanded);
  }

  function updateExpandCollapseButtonState(expanded) {
    treeExpanded = !!expanded;
    var $btn = $("#s6ExpandCollapseBtn");
    if ($btn.length) {
      if (treeExpanded) {
        $btn.text("Baum zuklappen");
        $btn.attr("aria-expanded", "true");
        $btn.attr("title", "Klappt den gesamten Berechtigungsbaum zu");
      } else {
        $btn.text("Baum aufklappen");
        $btn.attr("aria-expanded", "false");
        $btn.attr("title", "Klappt den gesamten Berechtigungsbaum auf");
      }
    }
  }

  function toggleExpandCollapseAll() {
    if (treeExpanded) {
      collapseAll();
      updateExpandCollapseButtonState(false);
    } else {
      expandAll();
      updateExpandCollapseButtonState(true);
    }
  }

  function expandAll() {
    $("#s6TreeRoot li").each(function () {
      var $li = $(this),
        $ul = $li.children("ul");
      if ($ul.length) {
        $ul.show();
        var $tog = $li.find("> .s6-node-row > .s6-toggle");
        if ($tog.length) $tog.text("▼").removeClass("hidden");
      }
    });
    treeExpanded = true;
    updateExpandCollapseButtonState(true);
  }

  function collapseAll() {
    $("#s6TreeRoot li").each(function () {
      var $li = $(this),
        $ul = $li.children("ul");
      if ($ul.length) {
        $ul.hide();
        var $tog = $li.find("> .s6-node-row > .s6-toggle");
        if ($tog.length) $tog.text("▶").removeClass("hidden");
      }
    });
    treeExpanded = false;
    updateExpandCollapseButtonState(false);
  }

  // --------------------------
  // Site / lists / items laden
  // --------------------------
  function loadContextInfo() {
    var p1 = getJson(currentSiteUrl + "/_api/web/currentuser?$select=LoginName");
    var p2 = getJson(currentSiteUrl + "/_api/web/effectivebasepermissions");
    return $.when(p1, p2).done(function (uData, pData) {
      if(uData[0].d !== undefined){
        currentUserLogin = uData[0].d.LoginName;
        var perms = pData[0].d.EffectiveBasePermissions;
        currentPermHigh = perms.High;
        currentPermLow = perms.Low;
      }
    });
  }

  function startBatchAnalysis() {
    var urls = [];
    var baseVal = $("#s6UrlInput").val();
    if (baseVal) urls.push(baseVal);
    $(".s6-url-input").each(function () {
      var v = $(this).val();
      if (v && v !== baseVal) urls.push(v);
    });
    if (urls.length === 0) {
      alert("Bitte mindestens eine URL eingeben.");
      return;
    }

    // UI Reset
    $("#s6ResultsArea").hide();
    $("#s6TreeRoot").empty();
    $("#s6UserTable tbody").empty();
    $("#s6LoadingArea").show();

    activeUsersMap = {};
    nodeLookup = {};
    updateExpandCollapseButtonState(false);

    processNextUrl(urls, 0);
  }

  function processNextUrl(urls, index) {
    if (index >= urls.length) {
      finalize();
      return;
    }
    currentSiteUrl = normalizeUrl(urls[index]);
    log("Verarbeite Site: " + currentSiteUrl, Math.round((index / urls.length) * 100));

    loadContextInfo()
      .always(function () {
        loadSite(currentSiteUrl).always(function () {
          processNextUrl(urls, index + 1);
        });
      });
  }

  function loadSite(siteUrl) {
    var def = $.Deferred();
    getJson(siteUrl + "/_api/web?$select=Title,Url,ServerRelativeUrl,Id,HasUniqueRoleAssignments")
      .done(function (d) {
        var web = d.d;
        if (!web.ServerRelativeUrl) {
          alert("ServerRelativeUrl nicht verfügbar: " + siteUrl);
          def.resolve();
          return;
        }

        var $webNode = createTreeNode("[" + siteUrl + "] " + web.Title, "web", web.ServerRelativeUrl, null, null, web.ServerRelativeUrl);
        $("#s6TreeRoot").append($webNode);
        nodeLookup[nodeKey(web.ServerRelativeUrl)] = $webNode;

        if (web.HasUniqueRoleAssignments) {
          loadPermissionsForNode($webNode, "web", null, web.Title, null);
        }
        loadLists(web.ServerRelativeUrl).always(function () {
          def.resolve();
        });
      })
      .fail(function (e) {
        alert("Fehler beim Laden der Site '" + siteUrl + "': " + e.statusText);
        def.resolve();
      });

    return def.promise();
  }

  function loadLists(webServerRelativeUrl) {
    var def = $.Deferred();
    log("Lade Listen...", 15);
    var ep =
      "/_api/web/lists?$select=Id,Title,Hidden,BaseType,HasUniqueRoleAssignments,RootFolder/ServerRelativeUrl&$expand=RootFolder&$filter=Hidden eq false";

    getJson(currentSiteUrl + ep)
      .done(function (d) {
        var lists = d.d.results;
        var queue = lists.slice();

        function next() {
          if (queue.length === 0) {
            def.resolve();
            return;
          }
          var list = queue.shift();
          if (!list.RootFolder || !list.RootFolder.ServerRelativeUrl) {
            next();
            return;
          }

          var listUrl = list.RootFolder.ServerRelativeUrl;
          var type = list.BaseType === 1 ? "lib" : "list";

          var $listNode = createTreeNode(list.Title, type, listUrl, list.Id, list.Id, listUrl);
          nodeLookup[nodeKey(listUrl)] = $listNode;

          var $webNode = nodeLookup[nodeKey(webServerRelativeUrl)];
          if ($webNode) getOrCreateUl($webNode).append($listNode);

          var needsDisplay = false;
          if (list.HasUniqueRoleAssignments) {
            loadPermissionsForNode($listNode, "list", list.Title, list.Title, list.Id);
            needsDisplay = true;
          }

          findUniqueItemsInList(list, function (found) {
            if (needsDisplay || found) $listNode.show();
            else $listNode.hide();
            next();
          });
        }

        next();
      })
      .fail(function () {
        def.resolve();
      });

    return def.promise();
  }

  function findUniqueItemsInList(list, cb) {
    var selectFields = "Id,Title,FileRef,FileDirRef,FileLeafRef,FileSystemObjectType,HasUniqueRoleAssignments";
    var url = currentSiteUrl + "/_api/web/lists(guid'" + list.Id + "')/items?$select=" + selectFields + "&$top=2000";

    getJson(url)
      .done(function (d) {
        var items = d && d.d && d.d.results ? d.d.results : [];
        var uniqueItems = items.filter(function (i) {
          return i.HasUniqueRoleAssignments;
        });
        buildTreeFromItems(list, uniqueItems);
        cb(uniqueItems.length > 0);
      })
      .fail(function () {
        cb(false);
      });
  }

  function buildTreeFromItems(list, items) {
    items.forEach(function (item) {
      if (!item.FileRef || !item.FileDirRef) return;

      var fullPath = item.FileRef;
      var parentPath = item.FileDirRef;

      ensureParentPathExists(list, parentPath);

      var $parentNode = nodeLookup[nodeKey(parentPath)];
      if ($parentNode) {
        var isFolder = item.FileSystemObjectType === 1;
        var type = isFolder ? "folder" : "file";
        var name = item.FileLeafRef || item.Title || "Item " + item.Id;

        var $itemNode = createTreeNode(name, type, fullPath, item.Id, list.Id, fullPath);
        getOrCreateUl($parentNode).append($itemNode);
        nodeLookup[nodeKey(fullPath)] = $itemNode;

        loadPermissionsForNode($itemNode, "item", list.Title, name, list.Id);
        revealParents($itemNode);
      }
    });
  }

  function ensureParentPathExists(list, path) {
    if (!path) return;
    var key = nodeKey(path);
    if (nodeLookup[key]) return;

    var root = list.RootFolder && list.RootFolder.ServerRelativeUrl ? list.RootFolder.ServerRelativeUrl.toLowerCase() : "";
    if (!root || path.toLowerCase() === root || path.length < root.length) return;

    var last = path.lastIndexOf("/");
    if (last === -1) return;

    var parent = path.substring(0, last);
    var folderName = path.substring(last + 1);

    ensureParentPathExists(list, parent);

    var $parent = nodeLookup[nodeKey(parent)];
    if ($parent) {
      var $folderNode = createTreeNode(folderName, "folder", path, null, list.Id, path);
      getOrCreateUl($parent).append($folderNode);
      nodeLookup[key] = $folderNode;
      $folderNode.hide();
    }
  }

  function revealParents($node) {
    if (!$node || $node.length === 0) return;
    $node.show();
    var $parentLi = $node.parent().closest("li");
    if ($parentLi.length > 0) {
      revealParents($parentLi);
      $parentLi.children("ul").show();
      var $tog = $parentLi.find("> .s6-node-row > .s6-toggle");
      if ($tog.length) $tog.removeClass("hidden").text("▼");
    }
  }

  // --------------------------
  // Tree Node Creation
  // --------------------------
  function createTreeNode(title, type, path, itemId, listId, serverRelativePath) {
    var $li = $("<li></li>");
    var $row = $("<div class='s6-node-row'></div>");
    var $toggle = $("<span class='s6-toggle hidden'>▶</span>");
    var iconChar = "📄";
    if (type === "web") iconChar = "🌐";
    else if (type === "list") iconChar = "📋";
    else if (type === "lib") iconChar = "📚";
    else if (type === "folder") iconChar = "📁";
    else if (type === "file") iconChar = "📄";
    var $icon = $("<span class='s6-icon'></span>").text(iconChar);
    var $title = $("<span class='s6-node-title'></span>").text(title);
    var $pathSpan = $("<span class='s6-node-path'></span>").text(serverRelativePath || path || "");

    var $actions = $("<div class='s6-node-actions'></div>");
    var $openBtn = $("<button class='s6-btn s6-btn-sm s6-node-action-btn' title='In neuem Tab öffnen'>🌐</button>");
    $openBtn.on("click", function (e) {
      e.stopPropagation();
      var url = getAbsoluteUrl(serverRelativePath || path || "");
      window.open(url, "_blank");
    });

    var $resetBtn = null;
    if (itemId) {
      $resetBtn = $("<button class='s6-btn s6-btn-sm s6-node-action-btn' title='Berechtigungen zurücksetzen'>♻️</button>");
      $resetBtn.on("click", function (e) {
        e.stopPropagation();
        var lId = listId || $li.data("listId");
        var iId = itemId;
        confirmAndResetItem(lId, iId, $li, serverRelativePath || path);
      });
    }

    var toggleHandler = function (e) {
      e.stopPropagation();
      var $ul = $li.children("ul");
      if ($ul.length) {
        $ul.slideToggle(200);
        $toggle.text($toggle.text() === "▶" ? "▼" : "▶");
      }
    };
    $toggle.on("click", toggleHandler);
    $row.on("click", toggleHandler);

    $row.append($toggle).append($icon).append($title).append($pathSpan);
    $actions.append($openBtn);
    if ($resetBtn) $actions.append($resetBtn);
    $row.append($actions);

    $li.append($row);

    $li.data("id", itemId);
    $li.data("listId", listId);
    $li.data("path", serverRelativePath || path);

    return $li;
  }

  function getOrCreateUl($li) {
    var $ul = $li.children("ul");
    if ($ul.length === 0) {
      $ul = $("<ul></ul>");
      $li.append($ul);
    }
    $li.children(".s6-node-row").find(".s6-toggle").removeClass("hidden");
    return $ul;
  }

  // --------------------------
  // Permissions laden & rendern
  // --------------------------
  function loadPermissionsForNode($li, type, listTitle, displayName, listId) {
    var id = $li.data("id");
    if (!id && type !== "web") return;

    var endpoint = "";
    if (type === "web") endpoint = "/_api/web/roleassignments";
    else if (type === "list") endpoint = "/_api/web/lists(guid'" + id + "')/roleassignments";
    else endpoint = "/_api/web/lists/getbytitle('" + encodeURIComponent(listTitle) + "')/items(" + id + ")/roleassignments";
    endpoint += "?$expand=Member,RoleDefinitionBindings";

    getJson(currentSiteUrl + endpoint).done(function (d) {
      var roles = d.d.results;
      if (roles.length > 0) {
        var $ul = getOrCreateUl($li);

        roles.reverse().forEach(function (role) {
          var member = role.Member;
          if (isExcludedMember(member)) return;

          var validRoleNames = (role.RoleDefinitionBindings.results || []).map(function (r) {
            return r.Name;
          }).filter(function (name) {
            return name !== IGNORE_ROLE && name !== "Limited Access";
          });

          if (validRoleNames.length === 0) return;

          if (member.PrincipalType === 1) {
            var key = siteUserKey(member.Id);
            if (!activeUsersMap[key]) activeUsersMap[key] = { user: member, contexts: [] };
            activeUsersMap[key].contexts.push({
              siteUrl: currentSiteUrl,
              type: type,
              listId: listId || id,
              itemId: type === "item" ? id : null,
              path: $li.data("path"),
              roles: validRoleNames,
            });
          }

          var roleString = validRoleNames.join(", ");
          var cleanLogin = getCleanLogin(member.LoginName);

          var $permLi = $("<li class='s6-perm-li'></li>");
          var $permRow = $("<div class='s6-perm-item'></div>");
          var $permText = $("<div class='s6-perm-text'></div>").text("- " + member.Title + " (" + cleanLogin + ") — " + roleString);
          var $permActions = $("<div class='s6-perm-actions'></div>");

          var $delBtn = $("<button class='s6-btn s6-btn-danger s6-btn-sm s6-node-action-btn' title='Berechtigung löschen'>🗑️</button>");
          $delBtn.on("click", function (e) {
            e.stopPropagation();
            var lId = listId || id;
            var iId = type === "item" ? id : null;
            deletePermissionCustom(type, lId, iId, member.LoginName, $permLi, $li);
          });

          var $copyBtn = $("<button class='s6-btn s6-btn-sm s6-node-action-btn s6-btn-copy' title='Rollen kopieren'>📋</button>");
          $copyBtn.on("click", function (e) {
            e.stopPropagation();
            copySingleContextRoles(validRoleNames, listId || id, type === "item" ? id : null, $li, type, listId || id);
          });

          $permActions.append($copyBtn).append($delBtn);
          $permRow.append($permText).append($permActions);
          $permLi.append($permRow);
          $ul.prepend($permLi);
        });
      }
    });
  }

  // --------------------------
  // Aktionen: delete/copy/reset
  // --------------------------
  function deletePermissionCustom(type, listId, itemId, loginName, $domLi, $contextNodeLi) {
    if (!confirm("Berechtigung löschen?")) return;

    var payload = {
      ListId: listId,
      LoginNames: [loginName],
      DataContext: null,
      MailSettings: { SendMail: false, MailType: "", AdditionalMessage: "" },
    };
    var url = "";
    if (type === "item" && itemId) {
      url = ENDPOINTS.removeItem;
      payload.ListItemId = itemId;
    } else if (type === "list") {
      url = ENDPOINTS.removeList;
    } else {
      alert("Web-Berechtigungen über diese API nicht unterstützt.");
      return;
    }

    makeCustomApiCall(url, payload)
      .done(function () {
        $domLi.slideUp(function () {
          $(this).remove();
        });
        refreshContextNode($contextNodeLi, type, listId);
        rebuildUserTable();
      })
      .fail(function (xhr) {
        alert("Fehler: " + xhr.statusText);
      });
  }

  function copySingleContextRoles(roles, listId, itemId, $contextNodeLi, type, typeListId) {
    var persNr = prompt("Ziel-Personalnummer eingeben:");
    if (!persNr) return;
    var targetLogin = "i:0#.w|itbw\\" + persNr;

    var payload = {
      ListId: listId,
      LoginNames: [targetLogin],
      Roles: roles,
      DataContext: null,
      MailSettings: { SendMail: false, MailType: "", AdditionalMessage: "" },
    };
    var url = itemId ? ENDPOINTS.assignItem : ENDPOINTS.assignList;
    if (itemId) payload.ListItemId = itemId;

    makeCustomApiCall(url, payload)
      .done(function () {
        alert("Rollen kopiert.");
        refreshContextNode($contextNodeLi, itemId ? "item" : "list", typeListId);
        rebuildUserTable();
      })
      .fail(function (xhr) {
        alert("Fehler beim Kopieren: " + xhr.statusText);
      });
  }

  function confirmAndResetItem(listId, itemId, $contextNodeLi, path) {
    if (!confirm("Berechtigungen für dieses Element zurücksetzen?")) return;

    var payload = { ListId: listId, ListItemId: itemId };

    makeCustomApiCall(ENDPOINTS.resetItem, payload)
      .done(function () {
        var key = nodeKey(path || $contextNodeLi.data("path") || "");
        if (key && nodeLookup[key]) delete nodeLookup[key];

        Object.keys(activeUsersMap).forEach(function (uid) {
          var entry = activeUsersMap[uid];
          entry.contexts = entry.contexts.filter(function (ctx) {
            return !(ctx.siteUrl === currentSiteUrl && ctx.type === "item" && ctx.listId === listId && ctx.itemId === itemId);
          });
          if (entry.contexts.length === 0) delete activeUsersMap[uid];
        });

        $contextNodeLi.slideUp(200, function () {
          var $parent = $contextNodeLi.parent().closest("li");
          $(this).remove();
          if ($parent.length) {
            var $ul = $parent.children("ul");
            if ($ul.length && $ul.children().length === 0) {
              $ul.remove();
              $parent.find("> .s6-node-row > .s6-toggle").addClass("hidden").text("▶");
            }
          }
        });

        rebuildUserTable();
      })
      .fail(function (xhr) {
        alert("Fehler beim Zurücksetzen: " + xhr.statusText);
      });
  }

  // --------------------------
  // Refresh / User-Table
  // --------------------------
  function refreshContextNode($li, type, listId) {
    if (!$li || $li.length === 0) return;

    $li.children(".s6-perm-li").remove();
    $li.children("ul").remove();

    var nodePathKey = nodeKey($li.data("path") || "");
    Object.keys(activeUsersMap).forEach(function (uid) {
      var entry = activeUsersMap[uid];
      entry.contexts = entry.contexts.filter(function (ctx) {
        var ctxKey = ctx.path ? ctx.path.toLowerCase() + "@" + ctx.siteUrl : "";
        return ctxKey !== nodePathKey;
      });
      if (entry.contexts.length === 0) delete activeUsersMap[uid];
    });

    var titleText = $li.find("> .s6-node-row .s6-node-title").text();
    loadPermissionsForNode($li, type, titleText, titleText, listId);
  }

  function rebuildUserTable() {
    var $tbody = $("#s6UserTable tbody");
    if ($tbody.length === 0) return;
    $tbody.empty();

    var uniqueUserIds = Object.keys(activeUsersMap);
    if (uniqueUserIds.length === 0) {
      $tbody.append("<tr><td>Keine direkten Benutzer-Berechtigungen gefunden.</td></tr>");
      return;
    }

    uniqueUserIds.forEach(function (uid) {
      var entry = activeUsersMap[uid];
      var u = entry.user;
      var clean = getCleanLogin(u.LoginName);
      var $tr = $("<tr></tr>");
      var $chk = $("<input type='checkbox' class='user-select' />").data("userid", uid);
      $tr.append($("<td></td>").append($chk));
      $tr.append("<td>" + (u.Title || "") + "</td>");
      $tr.append("<td>" + (clean || "") + "</td>");
      var $actions = $("<td></td>");

      var $btnCopy = $("<button class='s6-btn s6-btn-sm s6-btn-copy'>Kopieren</button>");
      $btnCopy.on("click", function () {
        copyUser(uid);
      });

      var $btnDel = $("<button class='s6-btn s6-btn-danger s6-btn-sm'>Löschen</button>");
      $btnDel.on("click", function () {
        if (!confirm("Alle gefundenen Berechtigungen für '" + (u.Title || "") + "' löschen?")) return;
        removeUserFromAllFoundContexts(uid, $tr, false)
          .then(function () {
            $tr.remove();
          })
          .catch(function () {
            alert("Fehler beim Löschen.");
          });
      });

      $actions.append($btnCopy).append(" ").append($btnDel);
      $tr.append($actions);
      $tbody.append($tr);
    });
  }

  function removeUserFromAllFoundContexts(userId, $tr, skipConfirm) {
    return new Promise(function (resolve, reject) {
      var entry = activeUsersMap[userId];
      if (!entry) {
        resolve();
        return;
      }

      if (!skipConfirm && !confirm("Alle gefundenen Berechtigungen für '" + entry.user.Title + "' löschen?")) {
        reject("abgebrochen");
        return;
      }

      $("#s6LoadingArea").show();
      log("Führe API Calls aus...", 0);

      var promises = [];
      var loginName = entry.user.LoginName;

      entry.contexts.forEach(function (ctx) {
        if (ctx.type !== "list" && ctx.type !== "item") return;
        currentSiteUrl = ctx.siteUrl;

        var payload = {
          ListId: ctx.listId,
          LoginNames: [loginName],
          DataContext: null,
          MailSettings: { SendMail: false, MailType: "", AdditionalMessage: "" },
        };
        var url = ctx.type === "item" ? ENDPOINTS.removeItem : ENDPOINTS.removeList;
        if (ctx.type === "item") payload.ListItemId = ctx.itemId;

        promises.push(makeCustomApiCall(url, payload));
      });

      if (promises.length === 0) {
        $("#s6LoadingArea").hide();
        delete activeUsersMap[userId];
        if ($tr) $tr.remove();
        resolve();
        return;
      }

      $.when.apply($, promises)
        .done(function () {
          delete activeUsersMap[userId];
          if ($tr) $tr.remove();
          $("#s6LoadingArea").hide();
          resolve();
        })
        .fail(function () {
          $("#s6LoadingArea").hide();
          reject("api-fail");
        });
    });
  }

  function copyUser(sourceUserId) {
    var entry = activeUsersMap[sourceUserId];
    if (!entry || !entry.contexts || entry.contexts.length === 0) {
      alert("Keine Berechtigungskontexte für diesen Benutzer gefunden.");
      return;
    }

    var persNr = prompt("Ziel-Personalnummer eingeben:");
    if (!persNr) return;
    var targetLogin = "i:0#.w|itbw\\" + persNr;

    if (!confirm("Berechtigungen von '" + entry.user.Title + "' an " + entry.contexts.length + " Orten auf '" + targetLogin + "' übertragen?")) return;

    log("Kopiere Berechtigungen...", 0);
    $("#s6LoadingArea").show();

    var promises = [];
    entry.contexts.forEach(function (ctx) {
      if (ctx.type !== "list" && ctx.type !== "item") return;
      currentSiteUrl = ctx.siteUrl;

      var payload = {
        ListId: ctx.listId,
        LoginNames: [targetLogin],
        Roles: ctx.roles,
        DataContext: null,
        MailSettings: { SendMail: false, MailType: "", AdditionalMessage: "" },
      };
      var url = ctx.type === "item" ? ENDPOINTS.assignItem : ENDPOINTS.assignList;
      if (ctx.type === "item") payload.ListItemId = ctx.itemId;

      promises.push(makeCustomApiCall(url, payload));
    });

    $.when.apply($, promises)
      .done(function () {
        alert("Kopiervorgang erfolgreich.");
        startBatchAnalysis();
      })
      .fail(function (xhr) {
        alert("Fehler beim Kopieren: " + (xhr && xhr.statusText ? xhr.statusText : "unknown"));
      })
      .always(function () {
        $("#s6LoadingArea").hide();
      });
  }

  // --------------------------
    // Untergeordnete Seiten abrufen (Search API) - FIXED PAYLOAD
    // --------------------------
    function fetchSubsitesForFirstUrl() {
        var firstUrl = $('#s6UrlInput').val();
        if (!firstUrl) {
            alert("Bitte geben Sie zuerst eine Site URL in das erste Feld ein.");
            return;
        }
        var siteCode = extractSiteCodeFromUrl(firstUrl);
        if (!siteCode) {
            alert("Die Site-Kennung (z. B. P5619770115) konnte aus der URL nicht ermittelt werden.\nBitte prüfen Sie die URL.");
            return;
        }

        var payload = buildSearchPayloadForSiteCode(siteCode);

        $('#s6LoadingArea').show();
        log("Rufe untergeordnete Seiten ab für " + siteCode + " …", 10);

        postSearchQuery(payload).then(function(json){
            var subsites = extractSPSiteUrlsFromSearchResult(json);
            if (!subsites || subsites.length === 0) {
                alert("Keine untergeordneten Seiten gefunden.");
                $('#s6LoadingArea').hide();
                return;
            }

            // Füge die gefundenen Subsite-URLs in die URL-Liste hinzu (als weitere Eingabezeilen)
            var $wrapper = $('#s6UrlList');
            subsites.forEach(function(u){
                // Duplikate vermeiden
                var existsAlready = false;
                $('.s6-url-input').each(function(){
                    if ($(this).val().replace(/\/+$/,'') === String(u).replace(/\/+$/,'')) existsAlready = true;
                });
                if (existsAlready) return;

                var $row = $('<div class="s6-url-row" style="display:flex; gap:6px; align-items:flex-start;"></div>');
                var inputId = 's6UrlInput_' + Date.now() + Math.floor(Math.random()*1000);
                var $input = $('<input type="text" class="s6-url-input" id="'+inputId+'" style="flex:1;" />');
                $input.val(u);
                var $remove = $('<button class="s6-btn s6-btn-danger s6-btn-sm" title="Zeile entfernen" style="padding:4px 8px; font-size:13px;">Entfernen</button>');
                $remove.on('click', function(){ $row.remove(); });
                $row.append($input).append($remove);
                $wrapper.append($row);
            });

            alert("Untergeordnete Seiten wurden hinzugefügt: " + subsites.length);
        }).catch(function(err){
            console.error(err);
            alert("Fehler beim Abrufen der untergeordneten Seiten: " + (err && err.message ? err.message : err));
        }).finally(function(){
            $('#s6LoadingArea').hide();
        });
    }

    // Extrahiert die Seitenkennung (z. B. P5619770115) aus der URL
    function extractSiteCodeFromUrl(url) {
        try {
            var u = new URL(url);
            var path = u.pathname; // z. B. /portale/P5619770115/Libs/Tools/...
            var m = path.match(/\/([Pp]\d{10,})\b/); // Muster: P + Ziffern (mind. 10)
            if (m && m[1]) return m[1];
            // Alternativ: nach "arbeitsbereiche/<code>" Muster
            var m2 = path.match(/\/arbeitsbereiche\/([A-Za-z0-9]+)/);
            if (m2 && m2[1]) return m2[1];
            return null;
        } catch (e) {
            return null;
        }
    }

    // BUILD FIXED SEARCH PAYLOAD (konform zum funktionierenden Beispiel)
    function buildSearchPayloadForSiteCode(siteCode) {
        return {
            request: {
                __metadata: { type: "Microsoft.Office.Server.Search.REST.SearchRequest" },
                ClientType: "PnPModernSearch",
                Properties: {
                    results: [
                        { Name: "EnableDynamicGroups", Value: { BoolVal: true, QueryPropertyValueTypeIndex: 3 } },
                        { Name: "EnableMultiGeoSearch", Value: { BoolVal: true, QueryPropertyValueTypeIndex: 3 } }
                    ]
                },
                Querytext: "*",
                TimeZoneId: 4,
                EnableQueryRules: true,
                QueryTemplate: "{searchTerms} contentclass=STS_Site RefinableString118=\"*" + siteCode + "*\" WebTemplate:STS RefinableString112=1",
                RowLimit: 500,
                SelectProperties: {
                    results: [
                        "Title","Path","Created","Filename","SiteLogo","PreviewUrl","PictureThumbnailURL",
                        "ServerRedirectedPreviewURL","ServerRedirectedURL","HitHighlightedSummary","FileType",
                        "contentclass","ServerRedirectedEmbedURL","ParentLink","DefaultEncodingURL",
                        "owstaxidmetadataalltagsinfo","Author","AuthorOWSUSER","SPSiteUrl","SiteTitle","IsContainer",
                        "IsListItem","HtmlFileType","SiteId","WebId","UniqueID","OriginalPath","FileExtension",
                        "IsDocument","NormSiteID","NormListID","NormUniqueID","AssignedTo"
                    ]
                },
                TrimDuplicates: false,
                SortList: { results: [ { Property: "LastModifiedTime", Direction: 1 } ] },
                Culture: 1031,
                Refiners: "",
                HitHighlightedProperties: { results: [] },
                RefinementFilters: { results: [] },
                ReorderingRules: { results: [] }
            }
        };
    }

    // POST SEARCH QUERY (konform)
    function postSearchQuery(payload) {
        var baseUrl = (window._spPageContextInfo && _spPageContextInfo.webAbsoluteUrl)
            ? _spPageContextInfo.webAbsoluteUrl.replace(/\/$/, "")
            : (window.location.protocol + '//' + window.location.host);

        var url = baseUrl + "/_api/search/postquery";

        return getRequestDigest().then(function(digest){
            return fetch(url, {
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
                        throw new Error("Search-API HTTP " + resp.status + (serverMsg ? (": " + serverMsg) : ""));
                    });
                }
                return resp.json();
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
                    if (!resp.ok) throw new Error("ContextInfo HTTP " + resp.status);
                    return resp.json();
                }).then(function(parsed){
                    var digest = parsed && parsed.d && parsed.d.GetContextWebInformation && parsed.d.GetContextWebInformation.FormDigestValue;
                    if (digest) resolve(digest);
                    else reject(new Error("Kein FormDigest in ContextInfo."));
                }).catch(function(err){ reject(err); });
            } catch (ex) { reject(ex); }
        });
    }

    // Extrahiert SPSiteUrl aus dem Search-Ergebnis (kompakte Variante)
    function extractSPSiteUrlsFromSearchResult(searchResult) {
        try {
            var rows = [];
            var rel = searchResult && searchResult.PrimaryQueryResult && searchResult.PrimaryQueryResult.RelevantResults;
            var table = rel && rel.Table;
            if (!table) return [];
            if (Array.isArray(table.Rows)) rows = table.Rows;
            else if (table.Rows && Array.isArray(table.Rows.results)) rows = table.Rows.results;
            var urls = [];
            (rows || []).forEach(function(row){
                var cells = [];
                if (row.Cells && Array.isArray(row.Cells)) cells = row.Cells;
                else if (row.Cells && Array.isArray(row.Cells.results)) cells = row.Cells.results;
                (cells || []).forEach(function(cell){
                    var key = cell.Key || cell.key || cell.k;
                    var val = cell.Value || cell.value || cell.v;
                    if (key === "SPSiteUrl" && val) {
                        var u = String(val).replace(/\/+$/,'');
                        if (urls.indexOf(u) === -1) urls.push(u);
                    }
                });
            });
            return urls;
        } catch(e){ return []; }
    }

  // --------------------------
  // finalize / show results
  // --------------------------
  function finalize() {
    rebuildUserTable();
    log("Fertig.", 100);
    setTimeout(function () {
      $("#s6LoadingArea").hide();
    }, 500);
    $("#s6ResultsArea").fadeIn();
  }

  // --------------------------
  // Public API
  // --------------------------
  return {
    addUrlRow: addUrlRow,
    startBatchAnalysis: startBatchAnalysis,
    fetchSubsitesForFirstUrl: fetchSubsitesForFirstUrl,
    exportTreeToPdf: exportTreeToPdf,
  };
})();