var siteUrl = '/portale/P5619770115/';
var clientContext
var id = 1;

function update_SiteUrl() {
	//var linkdata = document.URL.split("/");
	siteUrl = list_updateData[0].SITE;
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
    var clientContext = new SP.ClientContext(siteUrl);
    var oList = clientContext.get_web().get_lists().getById(list_updateData[0].GUID) ;//.getByTitle('Ankndigungen');
    var now = new Date().toISOString(); 
    var query = "";    
    var camlQuery = new SP.CamlQuery();

    //var test = SP.ListOperation.Selection.getSelectedList();
    //alert(test.pageListId);
    


    query = query +'<View><Query><Where>';
    //query = query +'<Or>';
    //query = query +'<Geq><FieldRef Name="EventDate" /><Value Type="DateTime">'+ now +'</Value></Geq>';
    //query = query +'<And>';
    //query = query +'<Geq><FieldRef Name="EndDate" /><Value Type="DateTime">'+ now +'</Value></Geq>';
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
        url = siteUrl+"/SitePages/Aktuelles.aspx?news="+oListItem.get_id();
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
    alert('Links failed. ' + args.get_message() + '\n' + args.get_stackTrace());
}

$( document ).ready(function() {
    console.log( "ready!" );
    update_SiteUrl();
    Update_News();
});

