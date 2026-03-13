/**
*
*  Javascript validate field.
*  @Author: viet.nguyen@worldsoft-vn.net
*  @Create date: 07/05/2010  
*/
var colorInvalid = "#FF9933";
var colorValid = "#fff"; 

//set dispaly invalid in jquery
function setInvalid(obj, title)
{
    obj.css("backgroundColor", colorInvalid)
    obj.attr("title", title);
    obj.focus(function() 
    {
        obj.css("backgroundColor", colorValid);
        obj.attr("title", null);
    });
}

function setValid(obj)
{
    obj.css("backgroundColor", colorValid);
    obj.attr("title", null);
}
//kiểm tra ngày với format "dd/MM/yyyy"
function isDate(obj, empty)
{
    var title = "Ngày không hợp lệ, nhập theo định dạng dd/MM/yyyy";
    if(obj.val() == "")
        if(empty)
            return true
        else
        {
            title = "Bạn phải chọn ngày";
        }
    var d = convertDate(obj.val());
    if(d == "Invalid Date" || d == "NaN")
    {
        setInvalid(obj, title);
        return false;
    }
    else
        return true;
}

function checkDate(obj, empty)
{
    var title = "Ngày không hợp lệ, nhập theo định dạng dd/MM/yyyy";    
    var d = convertDate(obj.val());
    if(d == "Invalid Date" || d == "NaN")
    {
        setInvalid(obj, title);
        return false;
    }
    else
        return true;
}


//kiểm tra số thập phân cho thẻ input
function isInt(obj, empty)
{
    var title = "Không hợp lệ, giá trị phải là số nguyên";
    value = obj.val().replace(/,/gi, "");
    if(value == "" && empty)
        return true
    if(value == "")
        title = "Bạn phải nhập vào giá trị";
    if(!isInteger(value))
    {
        setInvalid(obj, title);
        return false;
    }
    else
        return true;
}

//kiểm tra số điện thoại, cmnd
function isPhone(obj, title)
{    
    value = obj.val().replace(/,/gi, "");        
    
    if(value =="")
        return true;
    if(!isInteger(value))
    {
        setInvalid(obj, title);
        return false;
    }
    else
        return true;
}

function isSoLuong(obj)
{           
    if(obj.val() == '')
    {
        title = "Bạn phải nhập giá trị!";
        setInvalid(obj, title);
        return false;
    }
    else if(obj.val() == 0)
    {                
        title = "Giá trị phải lớn hơn 0!";
        setInvalid(obj, title);
        return false;
    }
    else
        return true;
}

//kiểm tra số thập phân cho thẻ input
function isEmpty(obj,title)
{
    if(title == "")
        title = "Bạn phải nhập dữ liệu vào";
    if(obj.val() == '')
    {
        setInvalid(obj, title);
        return true;
    }
    else
        return false;
}

//kiểm tra combobox voi value = 0 duoc chon
function isZeroInCombo(obj, title) {
    if (title == "")
        title = "Bạn phải nhập dữ liệu vào";
    if (parseInt(obj.val()) == 0) {
        setInvalid(obj, title);
        return true;
    }
    else
        return false;
}

//kiểm tra số thập phân cho thẻ input
function isFloat(obj, empty)
{
    var title = "Không hợp lệ, giá trị phải là số thập phân";
    value = obj.val().replace(/,/gi, "");
    if(value == "" && empty)
        return true
    if(value == "")
        title = "Bạn phải nhập vào giá trị";
    if(!isNumber(value))
    {
        setInvalid(obj, title);
        return false;
    }
    else
        return true;
}

function isEmail(obj,title) 
{   
    var btloc=/^([\w-]+(?:\.[\w-]+)*)@((?:[\w-]+\.)*\w[\w-]{0,66})\.([a-z]{2,6}(?:\.[a-z]{2})?)$/i;
    if(btloc.test(obj.val())) 
    {        
        return true ;
    }               
    setInvalid(obj, title);
    return false;
}

//kiểm tra số thập phân dương cho thẻ input
function isFloatPositive(obj, empty)
{
    var title = "Không hợp lệ, giá trị phải là số thập phân";
    value = obj.val().replace(/,/gi, "");
    if(value == "" && empty)
        return true
    if(value == "")
        title = "Bạn phải nhập vào giá trị >= 0";
    if(!isNumberPositive(value))
    {
        setInvalid(obj, title);
        return false;
    }
    else
        return true;
}

//kiểm tra số phần trăm cho thẻ input
function isPercent(obj, empty)
{
    var title = "Không hợp lệ, giá trị phải là số >=0 && <= 100";
    value = obj.val().replace(/,/gi, "");
    if(value == "" && empty)
        return true
    if(value == "")
        title = "Bạn phải nhập vào giá trị";
    if(!isNumber(value))
    {
        setInvalid(obj, title);
        return false;
    }
    else
    {
        var percent = getNumber(value);
        if(percent >=0 && percent <= 100)
            return true;
        else
        {
            setInvalid(obj, title);
            return false;
        }
    }
}

//kiểm tra số thập phân
function isNumber(value)
{
    value = value.replace(/,/gi, '');
    if(value == '' || isNaN(value))
        return false;
    return true;    
}

//kiểm tra số thập phân
function isNumberPositive(value)
{
    value = value.replace(/,/gi, '');
    if(value == '' || isNaN(value))
        return false;
    return getNumber(value) >= 0;
}

//kiểm tra số nguyên
function isInteger(value)
{
    value = value.replace(/,/gi, '');
    if(isNumber(value) && value.split('.').length == 1)
        return true;
    return false; 
}

//lấy số thập, nếu có lỗi trả về 0
function getNumber(value)
{
    if(isNumber(value))
        return parseFloat(value.replace(/,/gi, ''));
    else 
        return 0;
}

//lấy số nguyên, nếu có lỗi trả về 0
function getInt(value)
{
    if(isInteger(value))
        return parseInt(value.replace(/,/gi, ''),10);
    else 
        return 0;
}

//convert XML Dom
function convertXmlDom(text)
{
    if(text != '')
    {
        if (window.DOMParser)
        {
            parser = new DOMParser();
            return parser.parseFromString(text,"text/xml");
        }
        else // Internet Explorer
        {
            xmlDoc = new ActiveXObject("Microsoft.XMLDOM");
            xmlDoc.async = "false";
            xmlDoc.loadXML(text);
            return xmlDoc;
        }
    }
    return null;
}

//tạo node xml	
function createNodeXml(tagName,value)
{
    var str = "<," + tagName + ",>";
    str += value;
    str += "<,/" + tagName + ",>";
    return str;
}

//tạo node xml với string encode	
function setXmlEncode(tagName,value)
{
    var str = "&lt;" + tagName + "&gt;";
    str += value;
    str += "&lt;&frasl;" + tagName + "&gt;";
    return str;
}

// đọc node xml
function getXmlNodeValue(obj, index, tagName)
{
    var node = obj[index].getElementsByTagName(tagName)[0];
    if(node.childNodes.length > 0)
        return node.childNodes[0].nodeValue;
    return '';
}

//Json    
$.fn.clearSelect = function() {
return this.each(function() {
    if (this.tagName == 'SELECT')
        this.options.length = 0;
    });
 } 

$.fn.fillSelectDropDownList = function(data) {
    return this.clearSelect().each(function() {
        if (this.tagName == 'SELECT') {
            var dropdownList = this;
            $.each(data, function(key, value) {                   
                var option = new Option(value, key);
                //alert(value.Value);
                if ($.browser.msie) {
                    dropdownList.add(option);
                }
                else {
                    dropdownList.add(option, null);
                }
            });
        }
    });
 } 
 
 // delete all row in tbody    
 function deleteAllRow(id_tbody)
 {
    var table = document.getElementById(id_tbody);
    for(var i = table.rows.length - 1; i > -1; i--)
        table.deleteRow(i);
 }
 
 // delete one row in tbody    
 function deleteRow(id_tbody, index, sTT)
 {
    var table = document.getElementById(id_tbody);
    table.deleteRow(index);
    for(var i = 0; i < table.rows.length; i++)
    {
        if(sTT == null)
            $('#'+ id_tbody +' tr:eq('+i+')').find("td:last").html("<a href='javascript:void(0);' onclick='deleteRow(\"" + id_tbody + "\"," + i + ")'>Xóa</a>");
        else
        {
            $('#'+ id_tbody +' tr:eq('+i+')').find("td:first").html(i+1);
            $('#'+ id_tbody +' tr:eq('+i+')').find("td:last").html("<a href='javascript:void(0);' onclick='deleteRow(\"" + id_tbody + "\"," + i + ", true)'>Xóa</a>");
        }
    }
 }
 
// show Popup    
function showPopup(url, width, height)
{
    var window_height = height;
    var window_width  = width;
    var height = window.screen.availHeight;
    var width = window.screen.availWidth;
    var left_point = parseInt(width/2) - parseInt(window_width/2);
    var top_point = parseInt(height/2) - parseInt(window_height/2);
    window.open(url,'','width=' + window_width +',height=' + window_height + ',left='+left_point+',top='+top_point+', toolbar=0, location=0, scrollbars=1, status=0, resizable=1');
}

/**
*
*  Javascript check all checkbox if cehcked is true or otherwise.
*  @Author: tho.ta@worldsoft-vn.net
*  @Create date: 03/12/2010  
*/

function checkAllTbl(tblName, chkRoot){	
	var x = document.getElementById(tblName);
	var rowAll = x.rows.length - 1;	
	var check = document.getElementById(chkRoot);
	if(check.checked)
	{
		$("[name='chk']").each(function() 
		{
			if($(this).attr('disabled') != true)
				$(this).attr('checked',true);
		});
	}
	else
	    $("[name='chk']").removeAttr('checked');
}

/**
*
*  Javascript handle when check one checkbox in table, checkbox all checked=false .
*  @Author: tho.ta@worldsoft-vn.net
*  @Create date: 03/12/2010  
*/

function checkOneTable(checkOne, tblName, chkRoot){
    if(checkOne.checked == false){
       document.getElementById(chkRoot).checked = false; 
    }
    else if(getCheckInTbl(tblName)){        
        document.getElementById(chkRoot).checked = true;          
    }
}

//Kiem tra xem cac checkbox trong table check het hay chua.
// return true if all checkbox checked else return false
function getCheckInTbl(tblName){
	for(var i = 0; i < $("[name='chk']").length; i++)
		if(!document.getElementsByName('chk').item(i).checked)
			return false;
	return true;
}

/*
*
* Kiem tra xem user co check vao mot dong nao do trong table hay ko
*/
function getCheck(tblName) {  
    for(var i = 0; i < $("[name='chk']").length; i++)
		if(document.getElementsByName('chk').item(i).checked)
			return true;
	return false;
}

/*
*
* Kiem tra xem user co check vao mot dong nao do trong table hay ko
* Su dung cho truong hop update va chi chon 1 checkbox
* Return value of checkbox
*/

function getCheckUpdate(tblName) {   
    var obj;
	for(var i = 0; i < $("[name='chk']").length; i++)
	{
	    obj = document.getElementsByName('chk').item(i);
		if(obj.checked)
			return obj.value;
	}
}

// Delete row not exist database
function deleteRowNotExistDb(id_tbody, index, sTT)
    {
        var table = document.getElementById(id_tbody);
        table.deleteRow(index);
        for(var i = 0; i < table.rows.length; i++)
        {
            if(sTT == null)
                $('#'+ id_tbody +' tr:eq('+i+')').find("td:last").html("<a href='javascript:void(0);' onclick='deleteRowNotExistDb(\"" + id_tbody + "\"," + i + ")'>Xóa</a>");
            else
            {
                $('#'+ id_tbody +' tr:eq('+i+')').find("td:first").html(i+1);                    
                if($('#'+ id_tbody +' tr:eq('+i+')').find("a").attr("alt") == "xoa")
                {
                    $('#'+ id_tbody +' tr:eq('+i+')').find("td:last").html("<a alt = 'xoa' href='javascript:void(0);' onclick='deleteRowNotExistDb(\"" + id_tbody + "\"," + i + ", true)'>Xóa</a>");                                            
                }                    
            }
        }
    }


    //kiem tra ngay dung dinh dang
    function dateValidation(value) {

        var obj = value.val();
        var valueDate = obj.split('/');

        var day = valueDate[0];

        var month = valueDate[1];
        var year = valueDate[2];
        if ((day < 1 || day > 31) || (month < 1 && month > 12) ||  (year.length != 4)) {
            return false;
        }

    }
// Get key enter/ ie/ firefox
    function getKeyPressEnter(e) {
        var ENTER_KEY = 13;
        var code = "";
        var flag = false;
        if (window.event) // IE
        {
            code = e.keyCode;
        }
        else if (e.which) // Netscape/Firefox/Opera
        {
            code = e.which;
        }
        if (code == ENTER_KEY) {
            flag = true;
        }
        return flag;
    }