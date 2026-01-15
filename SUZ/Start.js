$(".ControlZone").css("padding", "0px");
$(".ControlZone").css("margin", "0px");

var siteUrl = '/portale/P5619770115/';
var guid_ereignisse = "4dea2140-f5e6-4ad6-9f01-fe6a70d0c2b2";
var slides = [];

function Update_Termin() {
    var clientContext = new SP.ClientContext(siteUrl);
    var oList = clientContext.get_web().get_lists().getByTitle('Ereignisse');
    var now = new Date().toISOString(); 
    var query = "";    
    var camlQuery = new SP.CamlQuery();

    //var test = SP.ListOperation.Selection.getSelectedList();
    //alert(test.pageListId);
    
    query = query +'<View><Query><Where>';
    query = query +'<Or>';
    query = query +'<Geq><FieldRef Name="EventDate" /><Value Type="DateTime">'+ now +'</Value></Geq>';
    query = query +'<And>';
    query = query +'<Geq><FieldRef Name="EndDate" /><Value Type="DateTime">'+ now +'</Value></Geq>';
    query = query +'<Leq><FieldRef Name="EventDate" /><Value Type="DateTime">'+ now +'</Value></Leq>';
    query = query +'</And>';
    query = query +'</Or>';
    query = query +'</Where>';
    query = query +'<OrderBy><FieldRef Name="EventDate" Ascending="True" /></OrderBy>';
    query = query +'</Query><RowLimit>6</RowLimit></View>'; 

    camlQuery.set_viewXml(query);
    this.collListItem5 = oList.getItems(camlQuery);
        
    clientContext.load(collListItem5);
    clientContext.executeQueryAsync(Function.createDelegate(this, this.onQuerySucceeded_Termin), Function.createDelegate(this, this.onQueryFailed_Termin));         
}

function onQuerySucceeded_Termin(sender, args) {
    var html_Termin = '';
    var url;
    

    var listItemEnumerator = collListItem5.getEnumerator();

    var monthList = ["Jan", "Feb", "Mär", "Apr", "Mai", "Jun", "Jul", "Aug", "Sep", "Okt", "Nov", "Dez"];
    var c = 0;
    while (listItemEnumerator.moveNext()) {
        c++;
        var oListItem = listItemEnumerator.get_current();
        var d_start = oListItem.get_item('EventDate');
        var d_end = oListItem.get_item('EndDate');
        var allday = oListItem.get_item("fAllDayEvent"); 
        var timeString = "";
        url = siteUrl + "/_layouts/15/Event.aspx?ListGuid="+guid_ereignisse+"&ItemId="+oListItem.get_item('ID')



        var day_str_start = d_start.format("dd.MM.yyyy");
        var day_str_end = d_end.format("dd.MM.yyyy");

        var time_str_start = d_start.format("HH:ss");
        var time_str_end = d_end.format("HH:ss");

        var diff =  d_end - d_start;
        debugger;
        if(allday) {
            //ganztätig
            if(diff <= 86400000){
                //ganztätig
                timeString = day_str_start + " Ganztags";
            }else {
                //zeitraum
                timeString = day_str_start + " - " + day_str_end;
            }
        }else {
            //zeitraum + uhrzeiten
            if(diff <= 86400000){
                //ganztags
                timeString = day_str_start + " " + time_str_start + " - " + time_str_end;
            }else{
                //zeitraum + uhrzeiten
                timeString = day_str_start + " " + time_str_start + " - " + day_str_end + " " + time_str_end;
            }
        }
        
        var style = 'style="object-fit: scale-down; padding: 20px auto;"';
        var pic = "https://ustgber.ecm.bundeswehr.org/portale/P5619770115/Libs/img/kalender-100.png";
        if(oListItem.get_item('BannerUrl') !== null) { 
            pic = oListItem.get_item('BannerUrl').get_url(); 
            style = "";
        }

        html_Termin = html_Termin + '<div class="event_item" onClick="location.href=\''+url+'\';">';
        html_Termin = html_Termin + '<img class="event_img" '+style+' src="'+pic+'">';
        html_Termin = html_Termin + '<div class="event_bottom">';
        html_Termin = html_Termin + '<div class="event_DatumBox">'

        if(diff <= 86400000) {
            html_Termin = html_Termin + d_start.format("MMM") + '</br>';
            html_Termin = html_Termin + '<b class="event_Day">'+d_start.format("dd")+'</b>'; //
        }else{
            html_Termin = html_Termin + d_start.format("dd.MM.") + '</br>';
            html_Termin = html_Termin + '<div class="event_Line_Box"></div>';
            html_Termin = html_Termin + d_end.format("dd.MM.");
        }

        html_Termin = html_Termin + '</div>';	
        html_Termin = html_Termin + '<div class="event_TextBox">';
        html_Termin = html_Termin + '<font class="event_text">'+oListItem.get_item('Title')+'</font></br>';
        html_Termin = html_Termin + timeString+'</br>';
        html_Termin = html_Termin + oListItem.get_item('Location');

        html_Termin = html_Termin + '</div>';
        html_Termin = html_Termin + '</div>';
        html_Termin = html_Termin + '</div>';
    }

    if(c === 0) {
        //kein Termin gefunden
        pic = "https://ustgber.ecm.bundeswehr.org/portale/P5619770115/Libs/img/kalender-no_event-100.png";
        url = siteUrl + "_layouts/15/Events.aspx?ListGuid="+guid_ereignisse
        style = 'style="object-fit: scale-down; padding: 20px auto;"';

        var now = new Date();

        s_day = now.getDate();
        s_month = now.getMonth();

        s_day_str = s_day.toString();
        if(s_day_str.length == 1) s_day_str = "0" + s_day_str;


        html_Termin = "";
        html_Termin = html_Termin + '<div class="event_item" onClick="location.href=\''+url+'\';">';
        html_Termin = html_Termin + '<img class="event_img" '+style+' src="'+pic+'">';
        html_Termin = html_Termin + '<div class="event_bottom">';
        html_Termin = html_Termin + '<div class="event_DatumBox">'

        html_Termin = html_Termin + monthList[s_month] + '</br>';
        html_Termin = html_Termin + '<b class="event_Day">'+s_day_str+'</b>'; //

        html_Termin = html_Termin + '</div>';	
        html_Termin = html_Termin + '<div class="event_TextBox">';
        html_Termin = html_Termin + '<font class="event_text">Kein Event geplant</font></br>';
        html_Termin = html_Termin + ' '
        html_Termin = html_Termin + '</div>';
        html_Termin = html_Termin + '</div>';
        html_Termin = html_Termin + '</div>';

    }

    $("#events").html(html_Termin);
}

function onQueryFailed_Termin(sender, args) {
    alert('Termin failed. ' + args.get_message() + '\n' + args.get_stackTrace());
}

function Update_Links() {
    var clientContext = new SP.ClientContext(siteUrl);
    var oList = clientContext.get_web().get_lists().getByTitle('Links');
        
    var camlQuery = new SP.CamlQuery();
    camlQuery.set_viewXml('<View><Query><Where></Where></Query><RowLimit>25</RowLimit></View>'); // <eq><FieldRef Name=\'aufStartseiteAnzeigen\'/><Value Type=\'Checkbox\'>Yes</Value></eq>
    this.collListItem3 = oList.getItems(camlQuery);
        
    clientContext.load(collListItem3);
    clientContext.executeQueryAsync(Function.createDelegate(this, this.onQuerySucceeded_Links), Function.createDelegate(this, this.onQueryFailed_Links));         
}

function onQuerySucceeded_Links(sender, args) {
    var html_links = '';
    var listItemEnumerator = collListItem3.getEnumerator();
    
    var n = 0    
    while (listItemEnumerator.moveNext()) {
        var oListItem = listItemEnumerator.get_current();
        var class_img = "link_img";
        var class_txt = "link_text";

        if(n > 0) { 
            n = 0; 
            class_img = "link_img_right";
            class_txt = "link_text_right";
        } else {
            n = 1;
        }
        //oListItem.get_item('Link').get_url()

        var pic = "https://ustgber.ecm.bundeswehr.org/portale/P5619770115/Libs/img/b-externer-link-80.png";
        if(oListItem.get_item('Bild') !== null) { pic = oListItem.get_item('Bild').get_url(); }

        html_links = html_links + '<div class="link_item" onclick="window.open(\'' + oListItem.get_item('URL').get_url() + '\', \'_blank\');">';
        html_links = html_links + '<div class="'+class_img+'"><img max-width="100%" max-height="100%" src="' + pic + '" alt="' + oListItem.get_item('Kurzbeschreibung') + '"></div>';
        html_links = html_links + '<div class="'+class_txt+'">' + oListItem.get_item('Kurzbeschreibung') + '</div>';
        html_links = html_links + '</div>';

        //html_links = html_links + '<tr class="' + class_box + '" onclick="location.href=\'' + oListItem.get_item('Link').get_url() + '\';" style="cursor:pointer;">';
        //html_links = html_links + '<td width="50px" style="vertical-align: middle;"><img class="linkimg" src="' + oListItem.get_item('Bild').get_url() + '" alt="' + oListItem.get_item('Title') + '"></td>';
        //html_links = html_links + '<td width="275px" class="linktext">' + oListItem.get_item('Title') + '</td></tr>';
    }

    //html_links = html_links + '</table>';
    $("#links").html(html_links);
    //alert(html_links);
}

function onQueryFailed_Links(sender, args) {
    alert('Links failed. ' + args.get_message() + '\n' + args.get_stackTrace());
}

function Update_Menu() {
    var clientContext = new SP.ClientContext(siteUrl);
    var oList = clientContext.get_web().get_lists().getByTitle('Start_Menu');
        
    var camlQuery = new SP.CamlQuery();
    camlQuery.set_viewXml('<View><Query><Where></Where><OrderBy><FieldRef Name="Reihenfolge" Ascending="True" /></OrderBy></Query><RowLimit>25</RowLimit></View>');  //<Eq><FieldRef Name="FSObjType" /><Value Type="Integer">0</Value></Eq>
    this.collListItem4 = oList.getItems(camlQuery);
        
    clientContext.load(collListItem4);
    clientContext.executeQueryAsync(Function.createDelegate(this, this.onQuerySucceeded_Menu), Function.createDelegate(this, this.onQueryFailed_Menu));         
}

function onQuerySucceeded_Menu(sender, args) {
    var html_menu = '';
    var listItemEnumerator = collListItem4.getEnumerator();
    var n = 0;   
    
    while (listItemEnumerator.moveNext()) {
        var oListItem = listItemEnumerator.get_current();
		html_menu = html_menu + '<li class="list-item">';
        html_menu = html_menu + '<div onclick="location.href = \''+oListItem.get_item('Link').get_url()+'\';" class="menu_item" style="background-image: url('+oListItem.get_item('Bild').get_url()+');">'+oListItem.get_item('Title')+'</div>';
        html_menu = html_menu + '</li>'
    }

    //html_links = html_links + '</table>';
    $("#menu").html(html_menu);
    //if(n <= 5) $(".btn-next").addClass("hidden");
    //alert(html_files);
}

function onQueryFailed_Menu(sender, args) {
    alert('Menu failed. ' + args.get_message() + '\n' + args.get_stackTrace());
}

function Update_News() {
    var clientContext = new SP.ClientContext(siteUrl);
    var oList = clientContext.get_web().get_lists().getById('f77f9bc8-8b5c-4153-9922-fd1681c3aff6') ;//.getByTitle('Ankndigungen');
    var now = new Date().toISOString(); 
    var query = "";    
    var camlQuery = new SP.CamlQuery();

    //var test = SP.ListOperation.Selection.getSelectedList();
    //alert(test.pageListId);
    


    query = query +'<View><Query><Where>';
    //query = query +'<Or>';
    //query = query +'<Geq><FieldRef Name="EventDate" /><Value Type="DateTime">'+ now +'</Value></Geq>';
    //query = query +'<And>';
    query = query +'<Eq><FieldRef Name="Anzeigen" /><Value Type="Boolean">1</Value></Eq>';
    //query = query +'<Leq><FieldRef Name="EventDate" /><Value Type="DateTime">'+ now +'</Value></Leq>';
    //query = query +'</And>';
    //query = query +'</Or>';
    query = query +'</Where>';
    query = query +'<OrderBy><FieldRef Name="Modified" Ascending="True" /></OrderBy>';
    query = query +'</Query><RowLimit>10</RowLimit></View>'; 

    camlQuery.set_viewXml(query);
    this.collListItem15 = oList.getItems(camlQuery);
        
    clientContext.load(collListItem15);
    clientContext.executeQueryAsync(Function.createDelegate(this, this.onQuerySucceeded_News), Function.createDelegate(this, this.onQueryFailed_News));         
}

function onQuerySucceeded_News(sender, args) {
    var html_News = '';
    var listItemEnumerator = collListItem15.getEnumerator();
    var slides = [];

    //html_News = html_News + '<a class="aktuelles_prev" id="prev">&#10094;</a>';
    //html_News = html_News + '<a class="aktuelles_next" id="next">&#10095;</a>';
    html_News = html_News + '<div class="cycle-pager"></div>';
    html_News = html_News + '<div class="cycle-overlay"></div>';

    var c = 0;
    while (listItemEnumerator.moveNext()) {
	    c++;
        var oListItem = listItemEnumerator.get_current();

        url = "https://ustgber.ecm.bundeswehr.org/portale/P5619770115/SitePages/Aktuelles.aspx?news="+oListItem.get_id();
        if(oListItem.get_item('Alternativer_Link') != null) url = oListItem.get_item('Alternativer_Link').get_url();
        url = 'onClick="location.href=\''+url+'\';"';
        html_News = html_News + '<img '+url+' src="'+oListItem.get_item('Bild').get_url()+'" width="100%" height="100%" style="object-fit: cover; object-position: center; cursor: pointer;" data-title="'+oListItem.get_item('Title')+'" data-cycle-desc="'+oListItem.get_item('Kurzbeschreibung')+'">';
        //slides.push('<img '+url+' src="'+oListItem.get_item('Bild').get_url()+'" width="100%" height="100%" style="object-fit: cover; object-position: center;" data-title="'+oListItem.get_item('Title')+'" data-cycle-desc="'+oListItem.get_item('Kurzbeschreibung')+'">');
        //slides.push(html_News);
        
    }


    $("#slideshow").html(html_News);
    $('#slideshow').cycle({
        fx: 'cover', direction: 'right', timeout: 3000, speed: 300, next: '#next', prev: '#prev' //, pager: '#news_dots'
    });
}

function onQueryFailed_News(sender, args) {
    alert('News failed. ' + args.get_message() + '\n' + args.get_stackTrace());
}

//menu schieberegel

function resizeMenu() {
    let wrapW = $("#horizontal-nav .menu-wrap").width(),
        menuW = $("#horizontal-nav .menu").width();

    let itemsToScroll = 3,
        widthToScroll = 0,
        scrollX = parseFloat($("#horizontal-nav .menu-wrap .menu").css("left"));

    if ($(this).hasClass("btn-prev")) {
        let prevItemIndex,
        prevItemsWidth = 0;

        $("#horizontal-nav .list-item").each((i, el) => {
        if (prevItemIndex !== undefined) return;
        prevItemsWidth += $(el).outerWidth() + 14;
        if (Math.ceil(prevItemsWidth) > Math.abs(scrollX)) prevItemIndex = i;
        });

        for (
        let i = prevItemIndex;
        i >= 0 && i > prevItemIndex - itemsToScroll;
        i--
        )
        prevItemsWidth -=
            $(`#horizontal-nav .list-item:eq(${i})`).outerWidth() + 14;

        widthToScroll = scrollX - prevItemsWidth;
        let newScrollX = Math.abs(scrollX) + widthToScroll;
        $("#horizontal-nav .menu-wrap .menu").css({ left: newScrollX });

        $(this).toggleClass("hidden", !Math.floor(newScrollX));
        $(".btn-next").removeClass("hidden");
    } else {
        let nextItemIndex,
        prevItemsWidth = 0;

        $("#horizontal-nav .list-item").each((i, el) => {
        if (nextItemIndex !== undefined) return;
        prevItemsWidth += $(el).outerWidth() + 14;
        if (Math.floor(prevItemsWidth - 14) > Math.abs(scrollX) + wrapW)
            nextItemIndex = i;
        });

        if (scrollX + wrapW >= menuW) {
        if (!$(this).hasClass("hidden")) $(this).addClass("hidden");
        return;
        }
        $(this).removeClass("hidden");

        for (
        let i = nextItemIndex + 1;
        i < nextItemIndex + itemsToScroll &&
        nextItemIndex + itemsToScroll <= $("#horizontal-nav .list-item").length;
        i++
        )
        prevItemsWidth +=
            $(`#horizontal-nav .list-item:eq(${i})`).outerWidth() + 14;
        widthToScroll = prevItemsWidth - 14 - (Math.abs(scrollX) + wrapW);
        let newScrollX = scrollX - widthToScroll;
        $("#horizontal-nav .menu-wrap .menu").css({ left: newScrollX });
        console.log(Math.round(Math.abs(newScrollX + wrapW)), menuW);
        $(this).toggleClass(
        "hidden",
        Math.round(Math.abs(newScrollX) + wrapW) >= Math.round(menuW)
        );
        $(".btn-prev").removeClass("hidden");
    }
}
$(() => {
    $("#horizontal-nav .list-item").each(function () {
        if ($(this).find(".sub-menu").length) $(this).addClass("has-submenu");
    });
    $("#horizontal-nav").on("click", ".btn-prev, .btn-next", resizeMenu);

    $(document).on("resize", resizeMenu);
});
  

Update_Termin();
Update_Links();
Update_Menu();

$( document ).ready(function() {
    console.log( "ready!" );
    Update_News();

    
});
//News_AutoSlide();