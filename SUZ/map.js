var is_open = false;

function map_hide_all(){
    if(!is_open) return;
    $("#SVZStadum").hide();
    $("#SVZKiel").hide();

    is_open = false;
}

function map_show(name){
    map_hide_all();
    $("#"+name).show(300);
    is_open = true;
}

function map_hide(name){
    $("#"+name).fadeOut(300);
    is_open = false;
}
