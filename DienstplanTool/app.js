(function () {
  "use strict";

  var LIST_DIENTE = "Dienste";
  var LIST_CONFIG = "DienstplanKonfiguration";
  var LIST_ADMINS = "DienstplanAdmins";

  var ctx = {
    currentUserId: _spPageContextInfo.userId,
    isAdmin: false,
    holidays: [],
    departments: [],
    currentFrom: null,
    currentTo: null,
    mailTemplate: "",
    admins: []
  };

  // ---------- Helper ----------

  function logDebug(msg, obj) {
    try {
      if (obj !== undefined) {
        console.log("[Dienstplan DEBUG] " + msg, obj);
      } else {
        console.log("[Dienstplan DEBUG] " + msg);
      }
    } catch (e) { /* ignore */ }
  }

  function formatDateISO(d) {
    return d.toISOString().substring(0, 10);
  }

  function formatDateDE(d) {
    var dd = String(d.getDate()).padStart(2, "0");
    var mm = String(d.getMonth() + 1).padStart(2, "0");
    var yyyy = d.getFullYear();
    return dd + "." + mm + "." + yyyy;
  }

  function getWeekdayNameDE(d) {
    var days = ["Sonntag","Montag","Dienstag","Mittwoch","Donnerstag","Freitag","Samstag"];
    return days[d.getDay()];
  }

  function isWeekend(d) {
    var day = d.getDay();
    return day === 0 || day === 6;
  }

  function isSpecialDate(d) {
    var month = d.getMonth() + 1;
    var day = d.getDate();
    return (month === 12 && (day === 24 || day === 31));
  }

  function isHoliday(d) {
    var iso = formatDateISO(d);
    return ctx.holidays.indexOf(iso) >= 0;
  }

  function getMonthStartEnd(date) {
    var start = new Date(date.getFullYear(), date.getMonth(), 1);
    var end = new Date(date.getFullYear(), date.getMonth() + 1, 0);
    return { start: start, end: end };
  }

  // ---------- Init ----------

  $(document).ready(function () {
    logDebug("Dokument bereit, initialisiere UI. currentUserId=" + ctx.currentUserId);
    initUI();

    logDebug("Starte Initialisierung ohne SP.SOD.executeFunc");

    checkIsAdmin()
      .then(function () {
        logDebug("checkIsAdmin fertig, ctx.isAdmin=" + ctx.isAdmin);
        return loadConfiguration();
      })
      .then(function () {
        logDebug("Konfiguration geladen", {
          holidays: ctx.holidays,
          departments: ctx.departments,
          mailTemplateLength: (ctx.mailTemplate || "").length
        });

        var now = new Date();
        var range = getMonthStartEnd(now);
        ctx.currentFrom = range.start;
        ctx.currentTo = range.end;
        $("#dp-from-date").val(formatDateISO(range.start));
        $("#dp-to-date").val(formatDateISO(range.end));
        updateMonthLabel(range.start);
        // Abteilungsfilter ist raus, daher kein fillDepartmentFilter mehr nötig
        if (ctx.isAdmin) {
          $("#dp-mail-template").val(ctx.mailTemplate || "");
        }
        logDebug("Starte initiales loadAndRenderServices");
        return loadAndRenderServices();
      })
      .fail(function (msg) {
        console.error("Fehler beim Initialisieren des Dienstplans: " + msg);
      });
  });

  // ---------- UI ----------

  function initUI() {
    logDebug("initUI aufgerufen");

    $(".dp-tab").on("click", function () {
      $(".dp-tab").removeClass("dp-tab-active");
      $(this).addClass("dp-tab-active");

      var tab = $(this).data("tab");
      logDebug("Tab-Klick: " + tab + ", ctx.isAdmin=" + ctx.isAdmin);

      if (tab === "intern") {
        if (ctx.isAdmin) {
          logDebug("Öffne internen Bereich");
          $("#dp-config-panel").show();
          $(".dp-filters").hide();
          $(".dp-table-wrapper").hide();
          initAdminPeoplePicker();
          loadAdmins();
        } else {
          logDebug("Verweigere Zugriff auf internen Bereich, ctx.isAdmin=false");
          alert("Nur Administratoren dürfen den internen Bereich öffnen.");
        }
      } else {
        $("#dp-config-panel").hide();
        $(".dp-filters").show();
        $(".dp-table-wrapper").show();
        loadAndRenderServices();
      }
    });

    $("#dp-from-date, #dp-to-date").on("change", function () {
      var fromVal = $("#dp-from-date").val();
      var toVal = $("#dp-to-date").val();
      logDebug("Filter geändert", { fromVal: fromVal, toVal: toVal });

      if (fromVal) ctx.currentFrom = new Date(fromVal);
      if (toVal) ctx.currentTo = new Date(toVal);
      updateMonthLabel(ctx.currentFrom);
      loadAndRenderServices();
    });

    $("#dp-prev-month").on("click", function () { logDebug("Prev Month Klick"); shiftMonth(-1); });
    $("#dp-next-month").on("click", function () { logDebug("Next Month Klick"); shiftMonth(1); });

    $("#dp-create-lists-btn").on("click", function () { logDebug("createLists-Button Klick"); createLists(); });
    $("#dp-save-config-btn").on("click", function () { logDebug("saveConfiguration-Button Klick"); saveConfiguration(); });

    $("#dp-refresh-admins-btn").on("click", function () {
      logDebug("dp-refresh-admins-btn Klick, ctx.isAdmin=" + ctx.isAdmin);
      if (!ctx.isAdmin) { alert("Nur Administratoren dürfen Admins verwalten."); return; }
      loadAdmins();
    });

    $("#dp-add-admin-btn").on("click", function () {
      logDebug("dp-add-admin-btn Klick, ctx.isAdmin=" + ctx.isAdmin);
      if (!ctx.isAdmin) { alert("Nur Administratoren dürfen Admins verwalten."); return; }
      addAdminPrompt();
    });

    $("#dp-save-mail-template-btn").on("click", function () {
      logDebug("dp-save-mail-template-btn Klick");
      saveMailTemplate();
    });
  }

  function updateMonthLabel(date) {
    var months = ["Januar","Februar","März","April","Mai","Juni","Juli","August","September","Oktober","November","Dezember"];
    var text = months[date.getMonth()] + " " + date.getFullYear();
    $("#dp-current-month-label").text(text);
  }

  function shiftMonth(delta) {
    var current = ctx.currentFrom || new Date();
    var newDate = new Date(current.getFullYear(), current.getMonth() + delta, 1);
    var range = getMonthStartEnd(newDate);
    ctx.currentFrom = range.start;
    ctx.currentTo = range.end;
    $("#dp-from-date").val(formatDateISO(range.start));
    $("#dp-to-date").val(formatDateISO(range.end));
    updateMonthLabel(range.start);
    logDebug("shiftMonth, neuer Bereich", { from: ctx.currentFrom, to: ctx.currentTo });
    loadAndRenderServices();
  }

  // ---------- Admin-Prüfung ----------

  function checkIsAdmin() {
    var deferred = $.Deferred();

    logDebug("checkIsAdmin gestartet, LIST_ADMINS=" + LIST_ADMINS + ", currentUserId=" + ctx.currentUserId);

    var ctxJsom = SP.ClientContext.get_current();
    var web = ctxJsom.get_web();
    var lists = web.get_lists();
    var adminList;

    try {
      adminList = lists.getByTitle(LIST_ADMINS);
      ctxJsom.load(adminList);
    } catch (e) {
      logDebug("getByTitle(LIST_ADMINS) wirft Fehler, wird im executeQueryAsync-Fail behandelt", e);
    }

    ctxJsom.executeQueryAsync(
      function () {
        logDebug("Adminliste '" + LIST_ADMINS + "' existiert, starte CAML-Abfrage");
        var query = new SP.CamlQuery();
        query.set_viewXml(
          "<View><Query>" +
          "<Where><Eq>" +
          "<FieldRef Name='Benutzer' LookupId='TRUE' />" +
          "<Value Type='Integer'>" + ctx.currentUserId + "</Value>" +
          "</Eq></Where>" +
          "</Query><RowLimit>1</RowLimit></View>"
        );
        var items = adminList.getItems(query);
        ctxJsom.load(items);
        ctxJsom.executeQueryAsync(
          function () {
            var count = items.get_count();
            logDebug("Adminliste gelesen, Trefferanzahl für aktuellen Benutzer: " + count);
            ctx.isAdmin = (count > 0);
            deferred.resolve();
          },
          function (s, a) {
            console.warn("Fehler beim Lesen der Admin-Liste (CAML): " + a.get_message());
            ctx.isAdmin = false;
            deferred.resolve();
          }
        );
      },
      function (s, a) {
        console.warn("Admin-Liste '" + LIST_ADMINS + "' nicht gefunden oder nicht lesbar. Zugriff wird erlaubt.", a.get_message());
        ctx.isAdmin = true;
        deferred.resolve();
      }
    );

    return deferred.promise();
  }

  // ---------- Konfiguration ----------

  function loadConfiguration() {
    var deferred = $.Deferred();
    logDebug("loadConfiguration gestartet, LIST_CONFIG=" + LIST_CONFIG);

    var ctxJsom = SP.ClientContext.get_current();
    var web = ctxJsom.get_web();
    var lists = web.get_lists();
    var listConfig;

    try {
      listConfig = lists.getByTitle(LIST_CONFIG);
    } catch (e) {
      console.warn("Konfigurationsliste '" + LIST_CONFIG + "' existiert noch nicht.");
      ctx.holidays = [];
      ctx.departments = [];
      ctx.mailTemplate = "";
      deferred.resolve();
      return deferred.promise();
    }

    var query = new SP.CamlQuery();
    query.set_rowLimit(1);
    var items = listConfig.getItems(query);
    ctxJsom.load(items);

    ctxJsom.executeQueryAsync(
      function () {
        if (items.get_count() > 0) {
          var itEnum = items.getEnumerator();
          itEnum.moveNext();
          var item = itEnum.get_current();
          var feiertage = item.get_item("Feiertage");
          var abteilungen = item.get_item("Abteilungen");
          var mailVorlage = item.get_item("MailVorlage");

          logDebug("Konfigurations-Item gefunden", { feiertage: feiertage, abteilungen: abteilungen, mailVorlageLength: (mailVorlage || "").length });

          if (feiertage) {
            try { ctx.holidays = JSON.parse(feiertage); } catch (e) { console.warn("Feiertage JSON-Fehler", e); ctx.holidays = []; }
          } else ctx.holidays = [];

          if (abteilungen) {
            try { ctx.departments = JSON.parse(abteilungen); } catch (e) { console.warn("Abteilungen JSON-Fehler", e); ctx.departments = []; }
          } else ctx.departments = [];

          ctx.mailTemplate = mailVorlage || "";
        } else {
          logDebug("Keine Konfigurationseinträge vorhanden");
          ctx.holidays = [];
          ctx.departments = [];
          ctx.mailTemplate = "";
        }
        deferred.resolve();
      },
      function (s, a) {
        console.warn("Konfiguration konnte nicht geladen werden: " + a.get_message());
        ctx.holidays = [];
        ctx.departments = [];
        ctx.mailTemplate = "";
        deferred.resolve();
      }
    );

    return deferred.promise();
  }

  function saveConfiguration() {
    if (!ctx.isAdmin) {
      alert("Nur Administratoren können die Konfiguration ändern.");
      return;
    }

    logDebug("saveConfiguration gestartet");

    var holidaysText = $("#dp-holidays-json").val();
    var holidays;
    try {
      holidays = JSON.parse(holidaysText || "[]");
      ctx.holidays = holidays;
    } catch (e) {
      alert("Fehler: Feiertage müssen ein gültiges JSON-Array sein.");
      return;
    }

    var ctxJsom = SP.ClientContext.get_current();
    var web = ctxJsom.get_web();
    var listConfig;

    try {
      listConfig = web.get_lists().getByTitle(LIST_CONFIG);
    } catch (e) {
      alert("Konfigurationsliste '" + LIST_CONFIG + "' existiert noch nicht.");
      return;
    }

    var query = new SP.CamlQuery();
    query.set_rowLimit(1);
    var items = listConfig.getItems(query);
    ctxJsom.load(items);

    ctxJsom.executeQueryAsync(
      function () {
        var item;
        if (items.get_count() > 0) {
          var en = items.getEnumerator();
          en.moveNext();
          item = en.get_current();
        } else {
          var ci = new SP.ListItemCreationInformation();
          item = listConfig.addItem(ci);
          item.set_item("Title", "Standard");
        }

        item.set_item("Feiertage", JSON.stringify(holidays));
        item.set_item("Abteilungen", JSON.stringify(ctx.departments || []));
        item.set_item("MailVorlage", ctx.mailTemplate || "");
        item.update();

        ctxJsom.executeQueryAsync(
          function () {
            logDebug("Konfiguration gespeichert");
            $("#dp-config-status").text("Konfiguration gespeichert.");
            loadAndRenderServices();
          },
          function (s2, a2) {
            console.error("Fehler beim Speichern der Konfiguration: " + a2.get_message());
            $("#dp-config-status").text("Fehler beim Speichern der Konfiguration.");
          }
        );
      },
      function (s, a) {
        console.error("Fehler beim Laden der Konfigurationsliste: " + a.get_message());
        $("#dp-config-status").text("Fehler beim Speichern der Konfiguration.");
      }
    );
  }

  function saveMailTemplate() {
    if (!ctx.isAdmin) {
      alert("Nur Administratoren können die Mailvorlage ändern.");
      return;
    }
    ctx.mailTemplate = $("#dp-mail-template").val() || "";
    logDebug("saveMailTemplate, Länge=" + ctx.mailTemplate.length);
    saveConfiguration();
  }

  // ---------- PeoplePicker für Admins ----------

  function initAdminPeoplePicker() {
    logDebug("initAdminPeoplePicker gestartet");

    var pickerElementId = "dp-admin-peoplepicker";

    if ($("#" + pickerElementId + "_TopSpan").length > 0) {
      logDebug("PeoplePicker bereits initialisiert");
      return;
    }

    var schema = {};
    schema.PrincipalAccountType = "User";
    schema.SearchPrincipalSource = 15;
    schema.ResolvePrincipalSource = 15;
    schema.AllowMultipleValues = false;
    schema.MaximumEntitySuggestions = 50;
    schema.Width = "100%";

    SPClientPeoplePicker_InitStandaloneControlWrapper(pickerElementId, null, schema);
  }

  function getSelectedAdminFromPeoplePicker() {
    var pickerElementId = "dp-admin-peoplepicker";
    var picker = SPClientPeoplePicker.SPClientPeoplePickerDict[pickerElementId + "_TopSpan"];

    if (!picker) {
      logDebug("getSelectedAdminFromPeoplePicker: Picker nicht gefunden");
      return null;
    }

    var users = picker.GetAllUserInfo();
    logDebug("PeoplePicker ausgewählte Benutzer", users);

    if (!users || users.length === 0) return null;
    return users[0];
  }

  function clearAdminPeoplePicker() {
    var pickerElementId = "dp-admin-peoplepicker";
    var picker = SPClientPeoplePicker.SPClientPeoplePickerDict[pickerElementId + "_TopSpan"];
    if (picker) {
      picker.ClearResolvedUsers();
      picker.UnselectAll();
    }
  }

  // ---------- Dienste laden ----------

  function loadAndRenderServices() {
    var deferred = $.Deferred();

    if (!ctx.currentFrom || !ctx.currentTo) {
      logDebug("loadAndRenderServices abgebrochen, kein Datum gesetzt");
      deferred.resolve();
      return deferred.promise();
    }

    logDebug("loadAndRenderServices gestartet", { from: ctx.currentFrom, to: ctx.currentTo });

    var ctxJsom = SP.ClientContext.get_current();
    var web = ctxJsom.get_web();
    var list;

    try {
      list = web.get_lists().getByTitle(LIST_DIENTE);
    } catch (e) {
      console.warn("Liste für Dienste '" + LIST_DIENTE + "' existiert noch nicht.");
      $("#dp-tbody").empty().append(
        $("<tr>").append(
          $("<td>").attr("colspan", 4).text("Liste '" + LIST_DIENTE + "' existiert noch nicht.")
        )
      );
      deferred.resolve();
      return deferred.promise();
    }

    var fromIso = formatDateISO(ctx.currentFrom);
    var toIso = formatDateISO(ctx.currentTo);

    var whereXml =
      "<And>" +
        "<Geq><FieldRef Name='DienstDatum' />" +
          "<Value IncludeTimeValue='FALSE' Type='DateTime'>" + fromIso + "</Value>" +
        "</Geq>" +
        "<Leq><FieldRef Name='DienstDatum' />" +
          "<Value IncludeTimeValue='FALSE' Type='DateTime'>" + toIso + "</Value>" +
        "</Leq>" +
      "</And>";

    var activeTab = $(".dp-tab-active").data("tab");
    logDebug("loadAndRenderServices, activeTab=" + activeTab);

    if (activeTab === "offen") {
      whereXml =
        "<And>" + whereXml +
          "<Eq><FieldRef Name='Status' /><Value Type='Text'>Offen</Value></Eq>" +
        "</And>";
    } else if (activeTab === "meine") {
      whereXml =
        "<And>" + whereXml +
          "<Eq><FieldRef Name='Dienstfuehrer' LookupId='TRUE' />" +
            "<Value Type='Integer'>" + ctx.currentUserId + "</Value>" +
          "</Eq>" +
        "</And>";
    }

    var caml = "<View><Query><Where>" + whereXml + "</Where></Query></View>";
    logDebug("CAML für Dienste", caml);

    var query = new SP.CamlQuery();
    query.set_viewXml(caml);

    var items = list.getItems(query);
    ctxJsom.load(items);

    ctxJsom.executeQueryAsync(
      function () {
        var results = [];
        var enumerator = items.getEnumerator();
        while (enumerator.moveNext()) {
          var it = enumerator.get_current();
          var vals = it.get_fieldValues();
          results.push({
            Id: it.get_id(),
            DienstDatum: vals["DienstDatum"],
            Status: vals["Status"],
            Dienstfuehrer: vals["Dienstfuehrer"]
          });
        }
        logDebug("Dienste geladen, Anzahl=" + results.length);
        renderTable(results);
        deferred.resolve();
      },
      function (s, a) {
        console.warn("Dienste konnten nicht geladen werden: " + a.get_message());
        $("#dp-tbody").empty().append(
          $("<tr>").append(
            $("<td>").attr("colspan", 4).text("Fehler beim Laden der Daten aus '" + LIST_DIENTE + "'.")
          )
        );
        deferred.reject(a.get_message());
      }
    );

    return deferred.promise();
  }

  function renderTable(items) {
    var tbody = $("#dp-tbody");
    tbody.empty();

    if (!ctx.currentFrom || !ctx.currentTo) return;

    var dayItemsByDate = {};
    items.forEach(function (it) {
      var d = new Date(it.DienstDatum);
      var key = formatDateISO(d);
      if (!dayItemsByDate[key]) dayItemsByDate[key] = [];
      dayItemsByDate[key].push(it);
    });

    logDebug("renderTable, Tage mit Einträgen=" + Object.keys(dayItemsByDate).length);

    var current = new Date(ctx.currentFrom.getTime());
    while (current <= ctx.currentTo) {
      var dateKey = formatDateISO(current);
      var dayEntries = dayItemsByDate[dateKey] || [];

      // alle Tage anzeigen – auch ohne Eintrag
      if (dayEntries.length === 0) {
        appendRow(tbody, current, null);
      } else {
        dayEntries.forEach(function (entry) {
          appendRow(tbody, current, entry);
        });
      }
      current.setDate(current.getDate() + 1);
    }
  }

  function appendRow(tbody, date, entry) {
    var weekdayName = getWeekdayNameDE(date);
    var dateStr = formatDateDE(date);

    var tr = $("<tr>");
    if (isHoliday(date)) tr.addClass("dp-row-holiday");
    else if (isSpecialDate(date)) tr.addClass("dp-row-special-date");
    else if (isWeekend(date)) tr.addClass("dp-row-weekend");

    $("<td>").text(weekdayName).appendTo(tr);
    $("<td>").text(dateStr).appendTo(tr);

    var dienstfuehrerCell = $("<td>");
    if (entry && entry.Dienstfuehrer) {
      if (entry.Dienstfuehrer.get_lookupValue) {
        dienstfuehrerCell.text(entry.Dienstfuehrer.get_lookupValue());
      } else {
        dienstfuehrerCell.text(entry.Dienstfuehrer);
      }
    } else {
      dienstfuehrerCell.html("<span class='dp-tag dp-tag-open'>offen</span>");
    }
    dienstfuehrerCell.appendTo(tr);

    var actionCell = $("<td>");
    var btn = $("<button>").addClass("dp-btn-primary");

    if (!entry) btn.text("Dienst anlegen");
    else btn.text(entry.Dienstfuehrer ? "ändern" : "eintragen");

    if (!ctx.isAdmin) {
      btn.addClass("dp-btn-disabled");
      btn.prop("disabled", true);
      btn.attr("title", "Nur Administratoren können Dienste eintragen oder ändern.");
    } else {
      btn.on("click", function () {
        logDebug("Eintragen-Button Klick", { date: date, entryId: entry ? entry.Id : null });
        onClickEintragen(date, entry);
      });
    }

    actionCell.append(btn).appendTo(tr);
    tbody.append(tr);
  }

  // ---------- Eintragen / Mail ----------

  function onClickEintragen(date, entry) {
    if (!ctx.isAdmin) {
      alert("Nur Administratoren können Dienste eintragen oder ändern.");
      return;
    }

    logDebug("onClickEintragen", { date: date, entryId: entry ? entry.Id : null });

    // Nur Nutzer abfragen
    var defaultUserText = "";
    if (entry && entry.Dienstfuehrer && entry.Dienstfuehrer.get_lookupValue) {
      defaultUserText = entry.Dienstfuehrer.get_email() || entry.Dienstfuehrer.get_lookupValue();
    }
    var userInput = prompt("Dienstführer (E-Mail oder Name) für diesen Dienst:", defaultUserText);
    if (userInput === null) return;
    if (!userInput) { alert("Dienstführer darf nicht leer sein."); return; }

    logDebug("onClickEintragen: Nutzer-Eingabe", { userInput: userInput });

    var ctxJsom = SP.ClientContext.get_current();
    var web = ctxJsom.get_web();

    var people = SP.Utilities.Utility.resolvePrincipal(
      ctxJsom,
      web,
      userInput,
      SP.Utilities.PrincipalType.user,
      SP.Utilities.PrincipalSource.all,
      null,
      true
    );

    ctxJsom.executeQueryAsync(
      function () {
        var principalInfo = people.get_value();
        if (!principalInfo) {
          alert("Benutzer konnte nicht gefunden werden.");
          return;
        }

        var user = principalInfo.get_User();
        logDebug("resolvePrincipal erfolgreich", { userId: user.get_id(), title: user.get_title(), email: user.get_email() });
        ctxJsom.load(user);
        ctxJsom.executeQueryAsync(
          function () {
            saveDienstWithUser(date, entry, user);
          },
          function (s1, a1) {
            console.error("Fehler beim Laden des Benutzers: " + a1.get_message());
            alert("Fehler beim Auflösen des Benutzers.");
          }
        );
      },
      function (s, a) {
        console.error("Fehler beim Auflösen des Benutzers: " + a.get_message());
        alert("Benutzer konnte nicht gefunden werden.");
      }
    );
  }

  function saveDienstWithUser(date, entry, user) {
    logDebug("saveDienstWithUser", { date: date, entryId: entry ? entry.Id : null, userId: user.get_id() });

    var ctxJsom = SP.ClientContext.get_current();
    var web = ctxJsom.get_web();
    var list = web.get_lists().getByTitle(LIST_DIENTE);

    var item;
    if (entry && entry.Id) {
      item = list.getItemById(entry.Id);
    } else {
      var createInfo = new SP.ListItemCreationInformation();
      item = list.addItem(createInfo);
    }

    item.set_item("Title", "Dienst");
    item.set_item("DienstDatum", date);
    item.set_item("Status", "Belegt");

    var userVal = new SP.FieldUserValue();
    userVal.set_lookupId(user.get_id());
    item.set_item("Dienstfuehrer", userVal);

    item.update();

    ctxJsom.executeQueryAsync(
      function () {
        logDebug("Dienst gespeichert, versende ggf. Mail");
        sendMailToUser(user, date);
        loadAndRenderServices();
      },
      function (s, a) {
        console.error("Fehler beim Speichern des Dienstes: " + a.get_message());
        alert("Fehler beim Speichern.");
      }
    );
  }

  function sendMailToUser(user, date) {
    var email = user.get_email();
    if (!email) {
      console.warn("Benutzer hat keine E-Mail, Mail wird nicht gesendet.");
      return;
    }

    var template = ctx.mailTemplate || "Hallo {Name},\n\nSie wurden für einen Dienst am {Datum} eingetragen.\n\nViele Grüße\nIhr Dienstplan-Team";

    var body = template
      .replace(/{Name}/g, user.get_title())
      .replace(/{Datum}/g, formatDateDE(date));

    var subject = "Neuer Dienst am " + formatDateDE(date);

    logDebug("Sende Mail", { to: email, subject: subject, bodyPreview: body.substring(0, 100) });

    var url = _spPageContextInfo.webAbsoluteUrl + "/_vti_bin/SendEmail";

    $.ajax({
      url: url,
      type: "POST",
      data: JSON.stringify({
        'properties': {
          '__metadata': { 'type': 'SP.Utilities.EmailProperties' },
          'To': { 'results': [email] },
          'Subject': subject,
          'Body': body
        }
      }),
      headers: {
        'Accept': 'application/json;odata=verbose',
        'content-type': 'application/json;odata=verbose',
        'X-RequestDigest': $("#__REQUESTDIGEST").val()
      },
      success: function () {
        logDebug("Mail an " + email + " gesendet.");
      },
      error: function (err) {
        console.error("Fehler beim Senden der Mail: ", err);
      }
    });
  }

  // ---------- Admin-Liste ----------

  function loadAdmins() {
    logDebug("loadAdmins gestartet, LIST_ADMINS=" + LIST_ADMINS);

    var ctxJsom = SP.ClientContext.get_current();
    var web = ctxJsom.get_web();
    var list;

    try {
      list = web.get_lists().getByTitle(LIST_ADMINS);
    } catch (e) {
      console.warn("Adminliste '" + LIST_ADMINS + "' existiert noch nicht.");
      $("#dp-admin-list").html("<p>Noch keine Adminliste vorhanden.</p>");
      return;
    }

    var query = new SP.CamlQuery();
    query.set_viewXml("<View><RowLimit>200</RowLimit></View>");
    var items = list.getItems(query);
    ctxJsom.load(items);

    ctxJsom.executeQueryAsync(
      function () {
        var html = "<table class='dp-table'><thead><tr><th>Admin</th><th>Aktion</th></tr></thead><tbody>";
        var enumerator = items.getEnumerator();
        ctx.admins = [];
        var count = 0;

        while (enumerator.moveNext()) {
          var it = enumerator.get_current();
          var id = it.get_id();
          var userVal = it.get_item("Benutzer");
          var display = userVal ? userVal.get_lookupValue() : "(unbekannt)";
          var email = userVal ? userVal.get_email() : "";
          ctx.admins.push({ itemId: id, user: userVal });
          count++;

          html += "<tr>";
          html += "<td>" + display + (email ? " &lt;" + email + "&gt;" : "") + "</td>";
          html += "<td><button class='dp-btn-secondary dp-remove-admin-btn' data-id='" + id + "'>entfernen</button></td>";
          html += "</tr>";
        }

        html += "</tbody></table>";
        logDebug("Adminliste geladen, Anzahl=" + count);
        $("#dp-admin-list").html(html);

        $(".dp-remove-admin-btn").on("click", function () {
          var itemId = parseInt($(this).data("id"), 10);
          logDebug("removeAdmin Klick", { itemId: itemId });
          removeAdmin(itemId);
        });
      },
      function (s, a) {
        console.error("Fehler beim Laden der Adminliste: " + a.get_message());
        $("#dp-admin-list").html("<p>Fehler beim Laden der Adminliste.</p>");
      }
    );
  }

  function addAdminPrompt() {
    logDebug("addAdminPrompt (PeoplePicker) gestartet");

    var selected = getSelectedAdminFromPeoplePicker();
    if (!selected) {
      alert("Bitte zuerst einen Benutzer im PeoplePicker auswählen.");
      return;
    }

    logDebug("addAdminPrompt: ausgewählter Benutzer", selected);

    var loginName = selected.Key;

    var ctxJsom = SP.ClientContext.get_current();
    var web = ctxJsom.get_web();

    var user = web.ensureUser(loginName);
    ctxJsom.load(user);

    ctxJsom.executeQueryAsync(
      function () {
        logDebug("ensureUser OK", { userId: user.get_id(), title: user.get_title(), email: user.get_email() });

        var list = web.get_lists().getByTitle(LIST_ADMINS);
        var ci = new SP.ListItemCreationInformation();
        var item = list.addItem(ci);

        var userVal = new SP.FieldUserValue();
        userVal.set_lookupId(user.get_id());
        item.set_item("Benutzer", userVal);
        item.update();

        ctxJsom.executeQueryAsync(
          function () {
            logDebug("Admin-Eintrag gespeichert");
            clearAdminPeoplePicker();
            loadAdmins();
          },
          function (s2, a2) {
            console.error("Fehler beim Speichern des Admins: " + a2.get_message());
            alert("Fehler beim Speichern des Admins.");
          }
        );
      },
      function (s, a) {
        console.error("Fehler bei ensureUser: " + a.get_message());
        alert("Fehler beim Auflösen des Benutzers.");
      }
    );
  }

  function removeAdmin(itemId) {
    if (!confirm("Diesen Admin wirklich entfernen?")) return;

    logDebug("removeAdmin gestartet", { itemId: itemId });

    var ctxJsom = SP.ClientContext.get_current();
    var web = ctxJsom.get_web();
    var list = web.get_lists().getByTitle(LIST_ADMINS);
    var item = list.getItemById(itemId);
    item.deleteObject();

    ctxJsom.executeQueryAsync(
      function () { logDebug("Admin gelöscht"); loadAdmins(); },
      function (s, a) {
        console.error("Fehler beim Entfernen des Admins: " + a.get_message());
        alert("Fehler beim Entfernen des Admins.");
      }
    );
  }

  // ---------- Listen anlegen ----------

  function createLists() {
    if (!ctx.isAdmin) {
      alert("Nur Administratoren können Listen erstellen.");
      return;
    }

    LIST_DIENTE = $("#dp-list-dienste-name").val() || "Dienste";
    LIST_CONFIG = $("#dp-list-config-name").val() || "DienstplanKonfiguration";
    LIST_ADMINS = $("#dp-list-admins-name").val() || "DienstplanAdmins";

    logDebug("createLists gestartet", {
      LIST_DIENTE: LIST_DIENTE,
      LIST_CONFIG: LIST_CONFIG,
      LIST_ADMINS: LIST_ADMINS
    });

    $("#dp-create-lists-status").text("Listen werden erstellt ...");

    ensureListJsom(LIST_CONFIG, 100, [
      { Name: "Feiertage", Type: "Note" },
      { Name: "Abteilungen", Type: "Note" },
      { Name: "MailVorlage", Type: "Note" }
    ])
      .then(function () {
        return ensureListJsom(LIST_DIENTE, 100, [
          { Name: "DienstDatum", Type: "DateTime" },
          { Name: "Dienstfuehrer", Type: "User" },
          { Name: "Status", Type: "Choice", Choices: ["Offen", "Belegt", "Storniert"] }
        ]);
      })
      .then(function () {
        return ensureListJsom(LIST_ADMINS, 100, [
          { Name: "Benutzer", Type: "User" }
        ]);
      })
      .then(function () {
        logDebug("createLists erfolgreich");
        $("#dp-create-lists-status").text("Listen vorhanden oder erfolgreich erstellt.");
      })
      .fail(function (msg) {
        console.error("Fehler beim Erstellen der Listen: " + msg);
        $("#dp-create-lists-status").text("Fehler beim Erstellen der Listen. Details in der Konsole.");
      });
  }

  function ensureListJsom(title, baseTemplate, fields) {
    var deferred = $.Deferred();

    logDebug("ensureListJsom gestartet", { title: title, baseTemplate: baseTemplate });

    var ctxJsom = SP.ClientContext.get_current();
    var web = ctxJsom.get_web();
    var lists = web.get_lists();
    var list;

    try {
      list = lists.getByTitle(title);
      ctxJsom.load(list);
      ctxJsom.executeQueryAsync(
        function () {
          logDebug("Liste '" + title + "' existiert, prüfe Felder");
          ensureFieldsJsom(list, fields)
            .then(function () { deferred.resolve(); })
            .fail(function (e) { deferred.reject(e); });
        },
        function () {
          logDebug("Liste '" + title + "' existiert nicht, lege an");
          var info = new SP.ListCreationInformation();
          info.set_title(title);
          info.set_templateType(baseTemplate);
          list = lists.add(info);
          ctxJsom.load(list);
          ctxJsom.executeQueryAsync(
            function () {
              ensureFieldsJsom(list, fields)
                .then(function () { deferred.resolve(); })
                .fail(function (e) { deferred.reject(e); });
            },
            function (s2, a2) {
              deferred.reject("Fehler beim Anlegen der Liste '" + title + "': " + a2.get_message());
            }
          );
        }
      );
    } catch (e) {
      deferred.reject("Fehler beim Zugriff auf Listen '" + title + "': " + e.message);
    }

    return deferred.promise();
  }

  function ensureFieldsJsom(list, fields) {
    var deferred = $.Deferred();

    if (!fields || !fields.length) {
      deferred.resolve();
      return deferred.promise();
    }

    var ctxJsom = list.get_context();
    var fieldsCollection = list.get_fields();

    var index = 0;
    function addNext() {
      if (index >= fields.length) {
        deferred.resolve();
        return;
      }

      var f = fields[index++];
      var fieldXml;

      switch (f.Type) {
        case "Note":
          fieldXml = "<Field Type='Note' DisplayName='" + f.Name + "' Name='" + f.Name + "' />";
          break;
        case "DateTime":
          fieldXml = "<Field Type='DateTime' DisplayName='" + f.Name + "' Name='" + f.Name + "' Format='DateOnly' />";
          break;
        case "User":
          fieldXml = "<Field Type='User' DisplayName='" + f.Name + "' Name='" + f.Name + "' UserSelectionMode='PeopleAndGroups' />";
          break;
        case "Choice":
          var choices = (f.Choices || []).map(function (c) { return "<CHOICE>" + c + "</CHOICE>"; }).join("");
          fieldXml =
            "<Field Type='Choice' DisplayName='" + f.Name + "' Name='" + f.Name + "'>" +
            "<CHOICES>" + choices + "</CHOICES></Field>";
          break;
        default:
          fieldXml = "<Field Type='Text' DisplayName='" + f.Name + "' Name='" + f.Name + "' />";
      }

      logDebug("ensureFieldsJsom: versuche Feld anzulegen", { field: f });

      fieldsCollection.addFieldAsXml(fieldXml, true, SP.AddFieldOptions.defaultValue);
      ctxJsom.executeQueryAsync(
        function () { addNext(); },
        function (s, a) {
          console.warn("Feld '" + f.Name + "' konnte evtl. nicht angelegt werden (existiert schon?): " + a.get_message());
          addNext();
        }
      );
    }

    addNext();
    return deferred.promise();
  }

})();