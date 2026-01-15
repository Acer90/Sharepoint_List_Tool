var now = new Date();
var startDate = new Date(now.getFullYear(), now.getMonth(), 1); //start datum 1 des aktuellen Monats
var endDate = new Date();
endDate = new Date(endDate.setMonth(endDate.getMonth() + 5)); // 6 monat in die zukunft
endDate =  new Date(endDate.getFullYear(), endDate.getMonth() + 1, 0); //letzen tag des monats

//datum unten befüllen
$('#startDate').val(startDate.toLocaleDateString('en-CA'));
$('#endDate').val(endDate.toLocaleDateString('en-CA'));

var siteUrl = '/portale/P5603398863';
var calenders = [];
var events = [];
var calendarEl = document.getElementById('Kalender');
var calendar = new FullCalendar.Calendar(calendarEl, {
    initialView: 'dayGridMonth',
    weekNumbers: true,
    locale: 'de',
    headerToolbar: {
        start: 'prev,next today', // will normally be on the left. if RTL, will be on the right
        center: 'title',
        end: 'dayGridMonth,multiMonthYear,listYear' // will normally be on the right. if RTL, will be on the left
    },
    textColor: 'black',
    color: 'yellow',
    //eventColor: '#000000',
    eventClick: function(event) {
        if (event.url) {
            window.open(event.url, "_blank");
            return false;
        }
    },
    eventMouseEnter: function (info) {
        // Display description when hovering over an event
        var position = $("#Kalender").offset();
        $('#Kalender_Tooltip_List').html(info.event.extendedProps.list);
        $('#Kalender_Tooltip_List').css({"background-color": info.event.backgroundColor, "color": info.event.extendedProps.Font_Color});

        $('#Kalender_Tooltip_Name').html(info.event.title);
        if(info.event.extendedProps.Location != "") {
            $('#Kalender_Tooltip_Location_String').html(info.event.extendedProps.Location);
            $('#Kalender_Tooltip_Location').show();
        }else{
            $('#Kalender_Tooltip_Location_String').html("");
            $('#Kalender_Tooltip_Location').hide();
        }
        
        //zeitraum
        //debugger;
        var diff = info.event.end - info.event.start;
        if(info.event.allDay){
            if(diff > 86400000){
                //länger als ein Tag
                $('#Kalender_Tooltip_Date').html(info.event.start.format( "dd.MM.yyyy" ) + " - " + info.event.end.format( "dd.MM.yyyy" ));
            }else{
                $('#Kalender_Tooltip_Date').html(info.event.start.format( "dd.MM.yyyy" ));
            }
        }else{
            if(diff > 86400000){
                //länger als ein Tag
                $('#Kalender_Tooltip_Date').html(info.event.start.format( "dd.MM.yyyy HH:mm" ) + " - " + info.event.end.format( "dd.MM.yyyy HH:mm" ));
            }else{
                $('#Kalender_Tooltip_Date').html(info.event.start.format( "dd.MM.yyyy HH:mm" ) + " - " + info.event.end.format( "HH:mm" ));
            }
        }
        
        /*event.extendedProps = {
            list: data.Title, 
            default_display: data.display,
            Veranstaltungen: data.Veranstaltungen,
            Location: oListItem.get_item("BwVMLocationLookup").get_lookupValue(),
            Modified: oListItem.get_item("Modified"),
            Editor: oListItem.get_item("Editor"),
            CurTeilnehmer: oListItem.get_item("BwVMEventSlotOccupied"),
            MaxTeilnehmer: oListItem.get_item("BwVMEventSlotAmount"),
            Veranstalter: oListItem.get_item("BwVMEventOrganiserName"),
        };*/

        if(info.event.extendedProps.Veranstaltungen){
            if(info.event.extendedProps.MaxTeilnehmer>0){
                $('#Kalender_Tooltip_Teilnehmer').show();
                var color = "#008000";
                if(info.event.extendedProps.CurTeilnehmer >= info.event.extendedProps.MaxTeilnehmer) color = "#FF0000";
                var Teilnehmer = "<font style='color:"+color+"; font-weight:800;'>"+info.event.extendedProps.CurTeilnehmer + "/" + info.event.extendedProps.MaxTeilnehmer+ "</font>";
                $('#Kalender_Tooltip_Teilnehmer_String').html(Teilnehmer);
            }else{
                $('#Kalender_Tooltip_Teilnehmer').hide();
            }
            if(info.event.extendedProps.Veranstalter){
                $('#Kalender_Tooltip_Veranstalter_String').html(info.event.extendedProps.Veranstalter);
                $('#Kalender_Tooltip_Veranstalter').show();
            }else{
                $('#Kalender_Tooltip_Veranstalter').hide();
            }

        }else{
            $('#Kalender_Tooltip_Veranstalter').hide();
            $('#Kalender_Tooltip_Teilnehmer').hide();
        }

        
        $('#Kalender_Tooltip').css({"left": (info.jsEvent.pageX - 50 - position.left) + 'px', "top": (info.jsEvent.pageY + 90 - position.top) + 'px'}); 
        $('#Kalender_Tooltip').slideDown(200);
    },
    eventMouseLeave: function () {
        // Hide description when mouse leaves an event
        $('#Kalender_Tooltip').css({"display": "none"});
    }
});
calendar.render();
Load_KalenderList();

function Load_KalenderList() {

    var clientContext = new SP.ClientContext(siteUrl);
    var oList = clientContext.get_web().get_lists().getByTitle('Liste_Kalender');
        
    var camlQuery = new SP.CamlQuery();
    camlQuery.set_viewXml('<View><Query><Where></Where><OrderBy><FieldRef Name="Prio" Ascending="FALSE" /></OrderBy></Query><RowLimit>50</RowLimit></View>');  //<Eq><FieldRef Name="FSObjType" /><Value Type="Integer">0</Value></Eq>
    this.collListItem = oList.getItems(camlQuery);
        
    clientContext.load(collListItem);
    clientContext.executeQueryAsync(function () { 
        var listItemEnumerator = collListItem.getEnumerator();

        while (listItemEnumerator.moveNext()) {
            var oListItem = listItemEnumerator.get_current();
            if(!oListItem.get_item("Enable")) continue;

            var data = {};
            data.guid = oListItem.get_item("GUID0");
            data.Title = oListItem.get_item("Title");
            data.URL = oListItem.get_item("URL").get_url();
            data.Background_Color = oListItem.get_item("Background_Color");
            data.Font_Color = oListItem.get_item("Font_Color");
            data.Active = oListItem.get_item("Active_onStart");
            data.group = oListItem.get_id();
            data.display = oListItem.get_item("Display");
            data.Display_Context = oListItem.get_item("Display_Context");
            data.Veranstaltungen =  oListItem.get_item("Veranstaltungen");
            data.OnClick_Link =  oListItem.get_item("OnClick_Link");

            if(data.OnClick_Link != null){
                data.OnClick_Link = data.OnClick_Link.get_url();
            } 

            if(!data.URL.toString().includes("/Lists/")) continue; //nur listen verwenden
            var url_arr = data.URL.split("/Lists/");
            var site_arr = url_arr[0].split("/");
            data.siteUrl = "/"+site_arr[site_arr.length-2]+"/"+site_arr[site_arr.length-1]+"/";
            data.list_name = url_arr[1].split("/")[0];

            calenders.push(data);
            Load_Kalender(data);
            Display_Context(data);
        }
    }, function (sender, args) {
        alert('Laden der Kalender Informationen fehlgeschlagen: ' + args.get_message() + '\n' + args.get_stackTrace());
    });         
}

function Load_Kalender(data) {
    var clientContext = new SP.ClientContext(data.siteUrl);
    var oList
    if(data.Veranstaltungen){
        oList = clientContext.get_web().get_lists().getById(data.guid);
    }else{
        oList = clientContext.get_web().get_lists().getByTitle(data.list_name);
    }
    

    var str_start_date = startDate.toISOString(); 
    var str_end_date = endDate.toISOString(); 

    var query = "";    
    var camlQuery = new SP.CamlQuery();
    query = query +'<View><Query><Where>';
    query = query +'<And>';
        //nur einträge laden wenn wie kleinder sind nicht älter sind als 180 tage
        query = query +'<Leq><FieldRef Name="EventDate" /><Value Type="DateTime">'+ str_end_date +'</Value></Leq>';

        //und
        query = query +'<Or>';
            //einträge laden die nach dem start datum beginnen
            query = query +'<Geq><FieldRef Name="EventDate" /><Value Type="DateTime">'+ str_start_date +'</Value></Geq>';
            //oder wenn das Endatum größer das Startdatum ist und 
            query = query +'<Geq><FieldRef Name="EndDate" /><Value Type="DateTime">'+ str_start_date +'</Value></Geq>';
        query = query +'</Or>';
    query = query +'</And>';
    query = query +'</Where>';
    if(data.Veranstaltungen){
        query = query +'<OrderBy><FieldRef Name="BwVMEventDate" Ascending="True" /></OrderBy>';
    }else{
        query = query +'<OrderBy><FieldRef Name="EventDate" Ascending="True" /></OrderBy>';
    }
    query = query +'</Query>';
    query = query +'<RowLimit>0</RowLimit></View>';

    camlQuery.set_viewXml(query);
    var collListItem = oList.getItems(camlQuery); 
    clientContext.load(collListItem);
    clientContext.executeQueryAsync(function () { 
        var listItemEnumerator = collListItem.getEnumerator();

        while (listItemEnumerator.moveNext()) {
            
            var oListItem = listItemEnumerator.get_current();

            if(!data.Veranstaltungen){
                var RecurrenceData =  oListItem.get_item("RecurrenceData");

                if(RecurrenceData == null){
                    var event = {};
                    event.id = data.Title;
                    event.title = oListItem.get_item("Title");
                    event.start = oListItem.get_item("EventDate").toISOString();
                    event.end = oListItem.get_item("EndDate").toISOString();
                    event.allDay = oListItem.get_item("fAllDayEvent"); 
                    event.description = oListItem.get_item("Description"); 

                    if(data.Active) event.display = data.display; else event.display = "none";
                    if(data.guid != "")event.url = data.siteUrl+"_layouts/15/Event.aspx?ListGuid="+data.guid+"&ItemId="+oListItem.get_id();
                    
                    //event.groupId = data.group; 
                    //event.color = data.Font_Color;
                    event.backgroundColor = data.Background_Color; 
                    event.borderColor = data.Background_Color; 
                    event.textColor = "#000000"; 
                    
                    var diff = oListItem.get_item("EndDate") - oListItem.get_item("EventDate");
                    if(diff >= 86400000 || event.allDay) event.textColor = data.Font_Color; 

                    event.extendedProps = {
                        list: data.Title, 
                        default_display: data.display,
                        Veranstaltungen: data.Veranstaltungen,
                        Location: oListItem.get_item("Location"),
                        Modified: oListItem.get_item("Modified"),
                        Editor: oListItem.get_item("Editor"),
                        Font_Color: data.Font_Color,
                    };

                    console.info(event);
                    calendar.addEvent(event, true);
                }else{
                    let xmlDoc = $.parseXML(RecurrenceData);
                    let $xml = $(xmlDoc);
                    let weekly = $xml.find("weekly");
                    if(weekly.length > 0){
                        //wöchentlich
                        var event = {};
                        event.id = data.Title;
                        event.title = oListItem.get_item("Title");
                        event.start = oListItem.get_item("EventDate").toISOString();
                        event.end = oListItem.get_item("EndDate").toISOString();
                        event.allDay = oListItem.get_item("fAllDayEvent"); 
                        event.description = oListItem.get_item("Description");

                        if(data.Active) event.display = data.display; else event.display = "none";
                        if(data.guid != "")event.url = data.siteUrl+"_layouts/15/Event.aspx?ListGuid="+data.guid+"&ItemId="+oListItem.get_id();
                        
                        //event.groupId = data.group; 
                        //event.color = data.Font_Color;
                        event.backgroundColor = data.Background_Color; 
                        event.borderColor = data.Background_Color; 
                        event.textColor = "#000000"; 
                        //if((event.end-event.start)<86400000) event.textColor = data.Font_Color; 


                        event.extendedProps = {
                            list: data.Title, 
                            default_display: data.display,
                            Veranstaltungen: data.Veranstaltungen,
                            Location: oListItem.get_item("Location"),
                            Modified: oListItem.get_item("Modified"),
                            Editor: oListItem.get_item("Editor"),
                            Font_Color: data.Font_Color,
                        };

                        event.startTime = oListItem.get_item("EventDate").format( "HH:mm:ss" );
                        event.endTime = oListItem.get_item("EndDate").format( "HH:mm:ss" );

                        var days = [];
                        if(weekly.attr("mo") != null) days.push(1);
                        if(weekly.attr("di") != null) days.push(2);
                        if(weekly.attr("mi") != null) days.push(3);
                        if(weekly.attr("do") != null) days.push(4);
                        if(weekly.attr("fr") != null) days.push(5);
                        if(weekly.attr("sa") != null) days.push(6);
                        if(weekly.attr("so") != null) days.push(0);

                        event.daysOfWeek = days

                        console.info(event);
                        calendar.addEvent(event, true);
                    }else{
                        debugger;
                    }
                }
            }else{
                //veranstaltungen
                var event = {};
                event.id = data.Title;
                event.title = oListItem.get_item("Title");
                event.start = oListItem.get_item("BwVMEventDate").toISOString();
                event.end = oListItem.get_item("BwVMEndDate").toISOString();
                event.allDay = oListItem.get_item("BwVMIsAllDayEvent"); 
                event.description = oListItem.get_item("BwVMDescription"); 

                if(data.Active) event.display = data.display; else event.display = "none";
                if(data.OnClick_Link != null)event.url =data.OnClick_Link+"#/navs/all/items/"+oListItem.get_id()+"/views/cardView/detail";
                
                //event.groupId = data.group; 
                //event.color = data.Font_Color;
                event.backgroundColor = data.Background_Color; 
                event.borderColor = data.Background_Color; 
                event.textColor = "#000000"; 

                var diff = oListItem.get_item("BwVMEndDate") - oListItem.get_item("BwVMEventDate");
                if(diff >= 86400000 || event.allDay) event.textColor = data.Font_Color; 


                event.extendedProps = {
                    list: data.Title, 
                    default_display: data.display,
                    Veranstaltungen: data.Veranstaltungen,
                    Location: oListItem.get_item("BwVMLocationLookup").get_lookupValue(),
                    Modified: oListItem.get_item("Modified"),
                    Editor: oListItem.get_item("Editor"),
                    CurTeilnehmer: oListItem.get_item("BwVMEventSlotOccupied"),
                    MaxTeilnehmer: oListItem.get_item("BwVMEventSlotAmount"),
                    Veranstalter: oListItem.get_item("BwVMEventOrganiserName"),
                    Font_Color: data.Font_Color,
                };

                console.info(event);
                calendar.addEvent(event, true);
            }
        }
        calendar.render();
    }, function (sender, args) {
                alert('Laden der Kalender Informationen fehlgeschlagen: ' + args.get_message() + '\n' + args.get_stackTrace());
    }); 
}

function Display_Context(data){
    var subclass = "";
    if(data.Active) subclass = " c_active";

    $("#context_"+data.Display_Context).append('<div id="context_'+data.guid+'" onclick="Context_Toggle(this, \''+data.guid+'\');" class="context_news_item'+subclass+'">'+data.Title+'</div>');
    //debugger;
}

function Context_Toggle(item, guid){
    var list_id = calenders.map((el) => el.guid).indexOf(guid);
    if(list_id < 0) return;
    if(calenders[list_id].Active) {
        calenders[list_id].Active = false; 
        $("#context_"+guid.replace(" ", '')).removeClass("c_active");
    } else {
        calenders[list_id].Active = true;
        $("#context_"+guid.replace(" ", '')).addClass("c_active");
    }


    Update_Events();
}

function Select_All(context){
    $("#"+context+" > div").each((index, elem) => {
        var guid = elem.id.replace("context_", "");
        var list_id = calenders.map((el) => el.guid).indexOf(guid);
        if(!calenders[list_id].Active) {
             Context_Toggle(null, guid);
        }
    });
}

var search_str = "";
function Search(){
    var text = $('#context_tb_input').val();
    if(text.length > 2){
        search_str = text.toLowerCase();
        Update_Events();
    }else{
        if(search_str !== "") {
            search_str = "";
            Update_Events();
        }else{
            search_str = "";
        }
    }
}

function Update_Events(){
    var activ_data = {};
    calenders.forEach(element => {
        activ_data[element.Title] = element.Active;
    });

    calendar.getEvents().forEach(event => {
        var list_name = event.extendedProps.list;
        var title = event.title;
        var description = event.extendedProps.description;
        if(description == null ) description = "";

        title = title.toLowerCase();
        description = description.toLowerCase();

        if(!activ_data[list_name]){
            event.setProp("display", "none");
            console.info("("+event.id +") " +title + " - Display => none");
            return;
        }
        if(search_str !== "" && !title.includes(search_str) && !description.includes(search_str)){
            event.setProp("display", "none");
            console.info("("+event.id +") " +title + " - Display => none");
            return;
        }

        console.info("("+event.id +") " +title + " - Display => " + event.extendedProps.default_display);
        event.setProp("display", event.extendedProps.default_display); 
    });
}

function Refresh_All(){
    $("#context_Top").empty();
    $("#context_Bottom").empty();
    $("#context_Bottom2").empty();
    $("#context_Bottom3").empty();

    startDate = new Date($('#startDate').val());
    endDate = new Date($('#endDate').val());
    
    var listEvent = calendar.getEvents();
    listEvent.forEach(event => { 
        event.remove()
    });

    Load_KalenderList();
}