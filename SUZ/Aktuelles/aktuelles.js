var siteUrl = '/portale/P5619770115/';
var clientContext
var id = 1;

function update_SiteUrl() {
	var linkdata = document.URL.split("/");
	siteUrl = "/" + linkdata[3] + "/" + linkdata[4];
}

function date_to_String(d_val){
    var d_str = "";
    var day = d_val.getDate();
	var month = d_val.getMonth()+1;
	var year = d_val.getYear()+1900;
	
	var day_str = day.toString();
	var month_str = month.toString();
	var year_str = year.toString();

	if(day_str.length == 1) day_str = "0" + day_str;
	if(month_str.length == 1) month_str = "0" + month_str;

    d_str = day_str + "." + month_str + "." + year_str

    return d_str; 
}

var z_soll = 0;
var z_ist = 0;
var z_elements = 0;
function Update_Page_Title(){
    document.title = "Aktuelles: "+ z_elements + " - (" + z_ist + "/" + z_soll +")";
}

function Update_News() {
    this.clientContext = new SP.ClientContext(siteUrl);
    var oList = clientContext.get_web().get_lists().getById(list_guid);//.getByTitle('Ankndigungen');
    var query = "";    
    var camlQuery = new SP.CamlQuery();


    query = query +'<View><Query><Where>';
    query = query +'<Eq><FieldRef Name="ID" /><Value Type="Number">'+ id +'</Value></Eq>';
    query = query +'</Where>';
    query = query +'<OrderBy><FieldRef Name="Modified" Ascending="True" /></OrderBy>';
    query = query +'</Query><RowLimit>1</RowLimit></View>'; 

    camlQuery.set_viewXml(query);
    this.collListItem15 = oList.getItems(camlQuery);
        
    clientContext.load(collListItem15); //, 'Include(Author)'
    clientContext.executeQueryAsync(Function.createDelegate(this, this.onQuerySucceeded_News), Function.createDelegate(this, this.onQueryFailed_News));         
}

function onQuerySucceeded_News(sender, args) {
    var html_News = '';
    var listItemEnumerator = collListItem15.getEnumerator();
    var title = ""
    var title_img = "";
    var Body = "";
    var images = "";
    var files = "";
    var creator = "";
    var creator_date;

    var c = 0;
    while (listItemEnumerator.moveNext()) {
	    c++;
        var oListItem = listItemEnumerator.get_current();
        title = oListItem.get_item('Title');
        title_img = oListItem.get_item('Bild').get_url();
        Body = oListItem.get_item('Body');

        creator = oListItem.get_item('Author');
        creator_date = oListItem.get_item('Modified');

        var attachmentFiles = oListItem.get_attachmentFiles();
        clientContext.load(attachmentFiles);
        clientContext.executeQueryAsync(function () {
            //debugger;
            if(attachmentFiles.get_count() > 0){
                var attachmentsItemsEnumerator = attachmentFiles.getEnumerator();
                while (attachmentsItemsEnumerator.moveNext()) {
                    var attachitem = attachmentsItemsEnumerator.get_current(); 
                    var fileName = attachitem.get_fileName();
                    var filepath = attachitem.get_path();
                    var serverPath = attachitem.get_serverRelativePath();
                    var serverUrl = attachitem.get_serverRelativeUrl();
                    var objetctData = attachitem.get_objectData();
                    var typedObj = attachitem.get_typedObject();
                    var extension = fileName.split('.').pop();

                    //debugger;

                    var is_img = false;
                    var img = "https://ustgber.ecm.bundeswehr.org/portale/P5619770115/Libs/img/ext/b-document-80.png"; 
                    switch (extension) {
                        case "docm":
                        case "dotx":
                        case "docx":{
                            img = "https://ustgber.ecm.bundeswehr.org/portale/P5619770115/Libs/img/ext/b-doc-80.png"; 
                            break;
                        }
                        case "pdf":{
                            img = "https://ustgber.ecm.bundeswehr.org/portale/P5619770115/Libs/img/ext/b-pdf-80.png"; 
                            break;
                        }
                        case "pptx":
                        case "ppt":{
                            img = "https://ustgber.ecm.bundeswehr.org/portale/P5619770115/Libs/img/ext/b-ppt-80.png"; 
                            break;
                        }
                        case "xlsx":
                        case "xlsm":{
                            img = "https://ustgber.ecm.bundeswehr.org/portale/P5619770115/Libs/img/ext/b-xls-80.png"; 
                            break;
                        }
                        case "rar":
                        case "zip":
                        case "7zip":{
                            img = "https://ustgber.ecm.bundeswehr.org/portale/P5619770115/Libs/img/ext/b-archiv-80.png"; 
                            break;
                        }
                        case "jpg":
                        case "gif": 
                        case "jpeg":
                        case "png":   
                        case "svg": {
                            is_img = true;
                            break;
                        }
                    }

                    if(is_img){
                        //in die bilderanzeige
                        images = images + '<img src="'+serverUrl+'" width="100%" height="600px" style="object-fit: cover; object-position: center;">';
                    }else{
                        //in die dateianzeige
                        files = files + '<div class="file_item" onclick="window.open(\''+serverUrl+'\', \'_blank\');">';
                        files = files + '<img class="file_item_pic" src="'+img+'">';
                        files = files + '<div class="file_item_txt">'+fileName+'</div>';
                        files = files + '</div>';
                    }


                    //debugger;
                    
                }
            }

            var style = '';
            if(files === "") style = 'style="width: 100%;"'

            html_News = html_News + '<img height="400px" src="'+title_img+'" style="margin: -40px -45px -40px -32px; object-fit:cover; max-width:1284px; width:1284px;">';

            html_News = html_News + '<div class="box_3">';

            html_News = html_News + '<div class="box_inhalt" '+style+'>';
            html_News = html_News + '<div class="box_Header">'+title+'</div>';
            html_News = html_News + '<div class="box_text">'+Body+'</div>';
            html_News = html_News + '<div class="box_creator"></div>';
            html_News = html_News + '</div>';

            if(files !== ""){
                html_News = html_News + '<div class="box_events">';
                html_News = html_News + '<div class="box_Header">FILES</div>';
                html_News = html_News + '<div class="box_Files">'+files+'</div>';
                html_News = html_News + '</div>';
            }

            if(images !== ""){
                html_News = html_News + '<div class="box_inhalt" style="width: 100%;">';
                html_News = html_News + '<div class="box_Header">Bilder</div>';
                html_News = html_News + '<div id="slideshow">'+images+'</div>';
                html_News = html_News + '</div>';
            }
            
            //debugger;
            html_News = html_News + '<div style="height: 50px; width: 100%;"></div>';
            html_News = html_News + '</div>';
            
            $("#inhalt").html(html_News);
            if(images !== ""){

                //debugger;
                $('#slideshow').cycle({
                    fx: 'cover', direction: 'right', timeout: 3000, speed: 300, next: '#next', prev: '#prev', pager: '#news_dots'
                });

            }
        });
    }
}

function onQueryFailed_News(sender, args) {
    alert('News failed. ' + args.get_message() + '\n' + args.get_stackTrace());
}

var listData = [];
var sel_TE = [];
var s_item = "date";
var s_dir = "desc";
var search = "";

function Update_NewsList(site, listID, TE, limit) {
    this.clientContext = new SP.ClientContext(site);
    var oList = clientContext.get_web().get_lists().getById(listID) ;//.getByTitle('Ankndigungen');
    var query = "";    
    var camlQuery = new SP.CamlQuery();

    sel_TE.push({TE: TE, enable: true});
    Display_NewsContext();

    query = query +'<View><Query><Where>';
    //query = query +'<Eq><FieldRef Name="ID" /><Value Type="Number">'+ id +'</Value></Eq>';
    query = query +'</Where>';
    query = query +'<OrderBy><FieldRef Name="Modified" Ascending="True" /></OrderBy>';
    query = query +'</Query><RowLimit>'+limit+'</RowLimit></View>'; 

    camlQuery.set_viewXml(query);
    var collListItem = oList.getItems(camlQuery);
        
    clientContext.load(collListItem); //, 'Include(Author)'
    clientContext.executeQueryAsync(function () { 
        //daten abrufbar
        var listItemEnumerator = collListItem.getEnumerator();
        var c = 0;

        while (listItemEnumerator.moveNext()) {
            var oListItem = listItemEnumerator.get_current();
            c++;

            var listitem = {
                id: oListItem.get_id(),
                te: TE,
                site: site,
                list: listID,
                title: oListItem.get_item('Title'),
                creator: oListItem.get_item('Editor').$5c_1,
                date: oListItem.get_item('Modified'),
                body:  oListItem.get_item('Body')
            };

            
            z_elements++;
            listData.push(listitem);
            //debugger;
        }

        if(c > 0){
            //update list
            Display_NewsList();
        }

        z_ist++;
        Update_Page_Title();
    }, function (sender, args) {
        alert('News failed. ' + args.get_message() + '\n' + args.get_stackTrace());
        //vorgang fehlgeschalgen
    });         
}

function Display_NewsContext(){
    html = "";
    if(sel_TE.length > 1){
        sel_TE.forEach(async (item) => {
            var subclass = "";
            if(item.enable) subclass = " c_active";
    
            html = html + '<div onclick="Context_Toggle(\''+item.TE+'\');" class="context_news_item'+subclass+'">'+item.TE+'</div>';
        });
    
        $("#context_news").show();
        $("#context_news").html(html);
    }else{
        $("#context_news").hide();
    }
    
    //debugger;
}

function Context_Toggle(str_te){
    sel_TE.forEach(function(item, index) {
        if(item.TE === str_te){
            if(item.enable){
                sel_TE[index].enable = false;
            }else{
                sel_TE[index].enable = true;
            }
        };
    });

    
    Display_NewsContext();
    Display_NewsList();
}

function Sort_News(item){
    if(this.s_item !== item){
        this.s_dir = "asc";
    }else{
        if(s_dir === "asc") this.s_dir = "desc"; else this.s_dir = "asc";
    }
    this.s_item = item;

    Display_NewsList();
}

function News_Search(){
    var text = $('#context_news_tb_input').val();
    if(text.length > 3){
        search = text.toLowerCase();
        Display_NewsList();
    }else{
        if(search !== "") {
            search = "";
            Display_NewsList();
        }else{
            search = "";
        }
    }
}

function Display_NewsList(){
    //sortiern der news liste
    switch(true){
        case s_item == "date" && s_dir == "asc":
            listData.sort((a,b)=> (a.date - b.date));
            break;
        case s_item == "date" && s_dir == "desc": 
            listData.sort((a,b)=> (b.date - a.date));
            break;
        case s_item == "title" && s_dir == "asc":
            listData.sort((a,b)=> (a.title.localeCompare(b.title)));
            break;
        case s_item == "title" && s_dir == "desc": 
            listData.sort((a,b)=> (b.title.localeCompare(a.title)));
            break;
        case s_item == "creator" && s_dir == "asc": 
            listData.sort((a,b)=> (a.creator.localeCompare(b.creator)));
            break;
        case s_item == "creator" && s_dir == "desc": 
            listData.sort((a,b)=> (b.creator.localeCompare(a.creator)));
            break;
        case s_item == "te" && s_dir == "asc": 
            listData.sort((a,b)=> (a.te.localeCompare(b.te)));
            break;
        case s_item == "te" && s_dir == "desc": 
            listData.sort((a,b)=> (b.te.localeCompare(a.te)));
            break;
    }

    var html = "";
    html = html + '<tr class="list_news_header">';

    var p_up = "&uarr;";
    var p_down = "&darr;";

    arrow = p_up;
    if(this.s_dir === "asc") arrow = p_down;

    up_title = "";
    if(this.s_item == "title"){
        up_title = ' <font class="news_pfeil">' + arrow + '</font>';
    }

    up_te = "";
    if(this.s_item == "te"){
        up_te = ' <font class="news_pfeil">' + arrow + '</font>';
    }

    up_date = "";
    if(this.s_item == "date"){
        up_date = ' <font class="news_pfeil">' + arrow + '</font>';
    }

    up_creator = "";
    if(this.s_item == "creator"){
        up_creator = ' <font class="news_pfeil">' + arrow + '</font>';
    }

    html = html + '<th class="list_news_header_item" onclick="Sort_News(\'title\');">Titel'+up_title+'</th>';
    html = html + '<th class="list_news_header_item" onclick="Sort_News(\'te\');">Bereich'+up_te+'</th>';
    html = html + '<th class="list_news_header_item" onclick="Sort_News(\'date\');">Geändert'+up_date+'</th>';
    html = html + '<th class="list_news_header_item" onclick="Sort_News(\'creator\');">Geändert von'+up_creator+'</th>';
    html = html + '</tr>';

    var gestern = new Date();
    gestern.setDate(gestern.getDate() -1);

    listData.forEach(function(item) {
        var i_te = sel_TE.findIndex(sa => sa.TE === item.te);

        if(i_te === -1  || !sel_TE[i_te].enable) return;
        if(search !== ""){
            if(!item.title.toLowerCase().includes(search) && !item.body.toLowerCase().includes(search)){
                //nicht gefunden
                return;
            }
        }

        str_date = date_to_String(item.date);
        if(str_date === date_to_String(new Date())) str_date = "heute";
        if(str_date === date_to_String(gestern)) str_date = "gestern";

        var url = "https://ustgber.ecm.bundeswehr.org" + item.site + "/SitePages/Aktuelles.aspx?news=" + item.id;
        //item.date

        html = html + '<tr onclick="location.href = \''+url+'\';" class="list_news_item">';
        html = html + '<td class="list_news_title">'+item.title+'</td>';
        html = html + '<td>'+item.te+'</td>';
        html = html + '<td>'+str_date+'</td>';
        html = html + '<td>'+item.creator+'</td>';
        html = html + '</tr>';
    });

    $("#list_news").html(html);
    //debugger;
}

const queryString = window.location.search;
const urlParams = new URLSearchParams(queryString);

$( document ).ready(function() {
    console.log( "ready!" );

    z_soll = list_updateData.length;
    Update_Page_Title();
    update_SiteUrl();
    //test();

    if(urlParams.has('news')){
        id = urlParams.get('news');
        Update_News();
    }else{
        var html = "";
        $("#img").show();

        html = html + '<div>';
        html = html + '<div class="context_news" id="context_news"></div>';
        html = html + '<div class="context_news_tb" id="context_news_tb"><input id="context_news_tb_input" oninput="News_Search();" type="text"></div>';
        html = html + '</div>';
        html = html + '<div class="search_news" id="serach_news"></div>';
        html = html + '<table class="list_news" id="list_news"></table>';
        $("#inhalt").html(html);

        list_updateData.forEach(async (item) => {
            Update_NewsList(item.SITE, item.GUID, item.TE, item.LIMIT);
        });
    }
});




async function Abo_List(site, listid, abo_name) {
    var spContext = new SP.ClientContext(site);
    var user = spContext.get_web().get_currentUser();
    var alerts = user.get_alerts();
    var list = spContext.get_web().get_lists().getById(listid);//.getByTitle("Announcements");
    spContext.load(user);
    spContext.load(alerts);
    spContext.load(list);

    return await new Promise((resolve, reject) => {
        spContext.executeQueryAsync(async () => {
            //Create new alert Object and Set its Properties
            let notify = new SP.AlertCreationInformation;
            notify.set_title(abo_name);
            notify.set_alertFrequency(SP.AlertFrequency.immediate);
            notify.set_alertType(SP.AlertType.list);
            notify.set_list(list);
            notify.set_deliveryChannels(SP.AlertDeliveryChannel.email);
            notify.set_alwaysNotify(true);
            notify.set_status(SP.AlertStatus.on);
            notify.set_user(user);
            notify.set_eventType(SP.AlertEventType.all);

            notify.Filter = '0';

            debugger;
            alerts.add(notify);
            user.update();
            var result = await new Promise((resolve, reject) => {
                spContext.executeQueryAsync(function() { // onSuccess
                        console.log("Alert added success fully")
                        resolve(1);
                    },
                    function(sender, args) { // onError
                        reject(args.get_message());

                    },
                );
            });
            resolve(result);
        }, (error) => {
            console.error(`Error while fetching alerts for user. ${error}`);
            reject(new Error(error));
        });
    });
}