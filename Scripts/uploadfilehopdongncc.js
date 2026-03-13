var names = [];
var pfileSize = [];
function SelectFiles(obj) {
    $('#btnUpLoad').attr('disabled', false);
    if ($('#txtFileName').val() != '') {
        names = $('#txtFileName').val().split(',');
    }
    if ($('#txtfileSize').val() != '') {
        pfileSize = $('#txtfileSize').val().split(',');
    }

    for (var i = 0; i < $(obj).get(0).files.length; ++i) {

        // build gia tri vao
        var fileSize = Math.round($(obj).get(0).files[i].size * 100 / (1024 * 1024)) / 100;
        if (fileSize > 100) {
            //alert("Your file is not uploaded! File size exceeds the maximum  size  (100MB)");
            alert("Kích thước File quá lớn. \n Vui lòng chọn file có kích thước <=100MB.");
            return;
        }
        if ($(obj).get(0).files[i].size > 1024 * 1024)
            fileSize = (Math.round($(obj).get(0).files[i].size * 100 / (1024 * 1024)) / 100).toString() + 'MB';
        else
            fileSize = (Math.round($(obj).get(0).files[i].size * 100 / 1024) / 100).toString() + 'KB';

        var tenFile = $(obj).get(0).files[i].name; //.replace(" ","");
        if (tenFile.split(',').length > 1 || tenFile.split(';').length > 1 || tenFile.split('.').length > 2 || tenFile.split('\\').length > 1 || tenFile.split('/').length > 1 || tenFile.split('&').length > 1 || tenFile.split('+').length > 1) {
            //alert("File Name not validate! File Name can not include  context char: (,;./\&+) ");
            alert("File không được chứa các ký tự ( <b>, ; . / \\ & +</b> ) \n Vui lòng đổi tên file hoặc chọn file khác.");
            return;
        }
        names.push(tenFile);
        pfileSize.push(fileSize);
        var fileName = "<a style ='color:Blue;' title ='Download File Uploaded' onclick = \"DownloadFile('" + tenFile + "','" + $('#txtDocumentNo').val() + "')\">" + tenFile + "</a>";
        var filenameParty = "<a style='color: Blue;' title='Download File Uploaded' onclick=\"OpenViewFilePart('" + tenFile + "','" + $('#txtDocumentNo').val() + "')\">" + tenFile + "</a>";
        var pimg = $(obj).get(0).files[i].name.split('.');
        //----view file online---
        if (pimg[1].toLowerCase() == "doc" || pimg[1].toLowerCase() == "docx" || pimg[1].toLowerCase() == "xls" || pimg[1].toLowerCase() == "xlsx" || pimg[1].toLowerCase() == "pdf" || pimg[1].toLowerCase() == "ppt" || pimg[1].toLowerCase() == "pptx") {
            fileName = filenameParty;
        }
        var pimg = $(obj).get(0).files[i].name.split('.');
        //                var stimg = '<img width="24" height="24" src=\'<%= Url.Content("../Content/IconFiles/_imgtemplate.png") %>\' />';
        var stimg = '<img width="24" height="24" src="../../Content/images/_imgtemplate.png" />';
        var imgfile = stimg.replace('_imgtemplate', pimg[1]);
        var trnew = "<tr><td>" + imgfile + fileName + "</td><td>" + fileSize + "</td><td align='center'>" + $("#UserDangNhap").val() + "</td><td></td><td align='center'>" +
               "<a style='color: Blue; font-style: oblique;' href='javascript:void(0)' onclick ='fn_DeleteFile(this)'>Delete</a>" + "</td></tr>";
        $('#tblFiles > tbody:last').append(trnew);


    }
    uploadFile(obj);
    $('#txtFileName').val(names);
    $('#txtfileSize').val(pfileSize);
}


function fn_DeleteFile(obj) {
    if (confirm("Bạn muốn xóa file này?")) {
        var oRow = obj.parentElement.parentElement;
        var row = oRow.rowIndex - 1;
        names = $('#txtFileName').val().split(',');
        pfileSize = $('#txtfileSize').val().split(',');
        //alert($('#txtFileName').val());
        var pfile = names[row];
        var _fileNameDelete = pfile.split('.')[0];
        //$.post('<%=Url.Action("DeleteFile","Document")%>', { fileName: pfile, docNumber: $('#txtDocumentNo').val() });
        var stt = oRow.rowIndex;
        document.getElementById("tblFiles").deleteRow(stt);
        if (document.getElementById("fileInfo").getElementsByTagName("tr").length > 1) {
            document.getElementById("fileInfo").deleteRow(stt);
        }
        names.splice(row, 1); pfileSize.splice(row, 1);
        //                if ($('#txtDescription').val() == _fileNameDelete) {
        //                    var docNo10 = $('#txtDocumentNo').val();
        //                    var _tenMoi = '04';
        //                    if (names.length >= 1) { _tenMoi = names[0].split('.')[0]; }
        //                    $('#txtDocumentNo').val(docNo10.replace("-" + _fileNameDelete + "-", "-" + _tenMoi + "-"));
        //                    $('#txtDescription').val(_tenMoi);
        //                }
        $('#txtFileName').val(names);
       // DeLeTeFile(_fileNameDelete);
    }
}
function uploadFile(obj) {

    var fd = new FormData();

    if (!window.FileReader) {
        //alert("The file API isn't supported on this browser yet.");
        alert("File không được hỗ trợ trên trình duyệt này.");
        return;
    }
    for (var i = 0; i < $(obj).get(0).files.length; ++i) {
        fd.append("fileToUpload", $(obj).get(0).files[i]);

    }
    var xhr = new XMLHttpRequest();
    xhr.upload.addEventListener("progress", uploadProgress, false);
    xhr.addEventListener("load", uploadComplete, false);
    xhr.addEventListener("error", uploadFailed, false);
    xhr.addEventListener("abort", uploadCanceled, false);
    //alert(fd);
    // var formData = new FormData(document.getElementById("frmFile"))
    xhr.open("POST", "../../HopDong/UploadFileMutil", true);
    xhr.send(fd);
}

function uploadProgress(evt) {
    //alert(evt)
    if (evt.lengthComputable) {
        var percentComplete = Math.ceil(evt.loaded * 100 / evt.total);
        document.getElementById('progressNumber').innerHTML = percentComplete.toString() + '%';
    }
    else {
        document.getElementById('progressNumber').innerHTML = 'không thể tính toán';
    }
}

function uploadComplete(evt) {
    /* This event is raised when the server send back a response */
    alert("Upload file thành công!");
    //$('#fileInfo').html($('#divListFile').html());
}

function uploadFailed(evt) {
    //$('#tblFiles > tbody tr:last').remove();
    //alert("Upload file bị lỗi.");
    //location.reload();
    alert("Upload file bị lỗi. \n Vui lòng kiểm tra lại.", "Thông báo", function() {
        location.reload();
    });
}

function uploadCanceled(evt) {
    //alert("The upload has been canceled by the user or the browser dropped the connection.");
    alert("Upload file đã bị hủy bởi người sử dụng hoặc trình duyệt đã mất kết nối.");
}
function SaveFile() {   
    var rows = $("#tblFiles > tbody").html();
    $('#fileInfo > tbody:last').html(rows);
    $('#closePopup').click();
}

