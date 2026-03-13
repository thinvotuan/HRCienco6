
//convert date(dd/MM/yyyy) to date(MM/dd/yyyy)
function convertDate(value)
{
    var arr = value.split('/');
    return new Date(parseInt(arr[1],10)+'/'+parseInt(arr[0],10)+'/'+parseInt(arr[2],10));
}

//Get string from Date, tho.ta
function getStringDate(value)
{    
    return appenString(value.getDate()) + "/" + appenString((value.getMonth() + 1)) +"/" + value.getFullYear();
}

function getStringDateVDA(value) {
    return appenString(value.getDate()) + "/" + appenString((value.getMonth())) + "/" + value.getFullYear();
}

//Khi thang = 1 so thi + 0 vao
function appenString(str){  
    return str < 10 ? "0" + str : str;
}

function parseDateTimeByJson(dt){
    if(dt.length == 0 || dt == null || dt == ""){
        return "";
    }
    else{
        var d = eval("new " + dt.slice(1, -1));
        return appenString(d.getDate()) + "/" + appenString((d.getMonth() + 1)) + "/" + d.getFullYear()
    }
}
