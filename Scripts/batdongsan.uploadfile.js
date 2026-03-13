function handleFileSelectNext(evt) {
    var nguoiUpLoad = $("#nguoiUpLoad").val(),
        ngayUpload = moment().format("DD/MM/YYYY hh:mm:ss A"),
        j = 1;

    //Lấy chính input
    var that = $(evt.target),              
        files = evt.target.files;  // FileList object

    //Tạo lại một input file mới sau mỗi sự kiện change
    var parent = that.parent(),
        newInput = that.clone();
    that.removeAttr("id").hide();
    parent.append(newInput.val(""));

    //Lấy arr những tên file đã upload
    var nameaccepts = [];
    $("#tbl-uploading").find("[name=nameaccept]").each(function () {
        nameaccepts.push(this.value);
    })

    var thumbnail = "",
        not_acceptable = "";
    var compressFiles = [".rar", ".zip"];
    // Loop through the FileList and render image files as thumbnails.
    for (var i = 0, f; f = files[i]; i++) {

        // Only process image files.                
        var reader = new FileReader();

        // Closure to capture the file information.
        reader.onload = (function (theFile) {

            return function (e) {
                if (theFile.type.includes('application/msword') || theFile.type.includes('application/vnd.openxmlformats-officedocument.wordprocessingml.document')) {
                    thumbnail = "/Images/thumbnail_word.png";
                }
                else if (theFile.type.includes('application/vnd.ms-excel') || theFile.type.includes('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')) {
                    thumbnail = "/Images/thumbnail_excel.png";

                }
                else if (theFile.type.includes('application/vnd.ms-powerpoint') || theFile.type.includes('application/vnd.openxmlformats-officedocument.presentationml.presentation')) {
                    thumbnail = "/Images/thumbnail_powerpoint.png";

                }
                else if (theFile.type.includes('application/pdf')) {
                    thumbnail = "/Images/thumbnail_pdf.png";

                }
                else if ($.inArray(theFile.name.substring(theFile.name.length - 4), compressFiles) != -1) {
                    thumbnail = "/Images/thumbnail_rar.png";
                }
                else if (theFile.name.substring(theFile.name.length - 4) == ".txt") {
                    thumbnail = "/Images/thumbnail_text.png";
                }
                else if (theFile.type.includes('image')) {
                    thumbnail = URL.createObjectURL(theFile);
                }
                else if (theFile.type == "application/x-msdownload") {
                    Alert.error("Định dạng tập tin " + theFile.name + " không được hỗ trợ", 'Lỗi định dạng', { displayDuration: 2000 });
                    return false;
                }

                if ($.inArray(theFile.name, nameaccepts) != -1)
                {
                    return;
                }

                // Render thumbnail.
                var tr = document.createElement('tr');
                tr.innerHTML = ['<td style="text-align: left"><img style="width:32px; height: 32px" src="', thumbnail, '" /></td>',
                                   '<td style="text-align: left"><b>', theFile.name, '</b></td>',
                                   '<td style="text-align: left">', nguoiUpLoad, '</td>',
                                   '<td style="text-align: left">', ngayUpload, '</td>',
                                   '<td style="text-align: right"><a style="color:#519ee5" href="javascript:void(0)" class="delRow"><i>Xóa</i></a></td>',
                                   '<td style="display:none"><input type="hidden" name="nameaccept" value="', theFile.name, '"/><input type="hidden" name="thumbnail" value="', theFile.name, '-SplitPoint-', thumbnail, '"/></td>',
                ].join('');
                $("#tbl-uploading >tbody").append(tr);
                j++;
            };
        })(f);
        // Read in the image file as a data URL.        
        reader.readAsDataURL(f);
    }
}

$(document).ready(function () {
    $(document).on("change", "#input-upload", function (e) {        
        handleFileSelectNext(e);        
    })

    $(document).on("click", ".delRow", function () {
        $(this).closest("tr").remove();
    })
});