function resetUnicode(str) {
    str = str.toLowerCase();
    str = str.replace(/à|á|ạ|ả|ã|â|ầ|ấ|ậ|ẩ|ẫ|ă|ằ|ắ|ặ|ẳ|ẵ/g, "a");
    str = str.replace(/è|é|ẹ|ẻ|ẽ|ê|ề|ế|ệ|ể|ễ/g, "e");
    str = str.replace(/ì|í|ị|ỉ|ĩ/g, "i");
    str = str.replace(/ò|ó|ọ|ỏ|õ|ô|ồ|ố|ộ|ổ|ỗ|ơ|ờ|ớ|ợ|ở|ỡ/g, "o");
    str = str.replace(/ù|ú|ụ|ủ|ũ|ư|ừ|ứ|ự|ử|ữ/g, "u");
    str = str.replace(/ỳ|ý|ỵ|ỷ|ỹ/g, "y");
    str = str.replace(/đ/g, "d");
    str = str.replace(/!|@|%|\^|\*|\(|\)|\+|\=|\<|\>|\?|\/|,|\.|\:|\;|\'| |\"|\&|\#|\[|\]|~|$|_/g, "-");

    str = str.replace(/-+-/g, "-"); //thay thế 2- thành 1- 
    str = str.replace(/^\-+|\-+$/g, "");
    return str;
}


(function ($) {
    $.fn.justText = function () {
        return $(this).clone()
                .children()
                .remove()
                .end()
                .text();

    };
    $.fn.filterBySearch = function (str) {
        var that = this, $allListElements = that.find('a');
        var $matchingListElements = $allListElements.filter(function (i, a) {
            var listItemText = $(a).text().toUpperCase();
            var searchText = str.toUpperCase();
            listItemText = resetUnicode(listItemText);
            searchText = resetUnicode(searchText);
            return ~listItemText.indexOf(searchText);
        });
        $allListElements.hide();
        return $matchingListElements.show();
    }

    $.fn.filterLink = function (str) {
        var that = this, $allListElements = that.find('li');
        var $matchingListElements = $allListElements.filter(function (i, li) {
            var listItemText = $(li).text().toUpperCase();
            var searchText = str.toUpperCase();
            listItemText = resetUnicode(listItemText);
            searchText = resetUnicode(searchText);
            return ~listItemText.indexOf(searchText);
        });
        $allListElements.hide();
        return $matchingListElements.show();
    }

    $.fn.imageUploader = function (params) {
        options.url = params.url;
        options.button = params.button;
        var boundnary = this,
            child = "<div class='form-upload'>"
                                       + "<div class='queue-inlist' id='form-upload-queue-inlist'>"
                                             + "<div class='place-container' id='form-upload-place-container'>"
                                             + "<div class='uploader-container'>"
                                                    + "<div class='uploader-picker'>Chọn tập tin</div>"
                                                    + "<div style='position: absolute; top: 0px; left: 510px; width: 160px; height: 44px; overflow: hidden; bottom: auto; right: auto;'>"
                                                      + "<input type='file' name='files' class='uploader-control' id='uploader-control-1' multiple='multiple' />"
                                                      + "<label for='uploader-control-1' style='opacity: 0; width: 100%; height: 100%; display: block; cursor: pointer; background: rgb(255, 255, 255);'></label></div>"
                                                    + "</div>"
                                            + "</div>"
                                            + "<ul class='file-inlist' id='form-upload-file-inlist'></ul>"
                                        + "</div>"
                                        + "<div class='uploader-status' id='form-upload-uploader-status'>"
                                        + "</div>"
                       + "</div>";        
        return $(this).append(child);
    }
}(jQuery));

var url, data, button = true, options = {
    url: "",
    data: "",
    button: button,
};

function addEventHandler(obj, evt, handler) {
    try {
        if (obj.addEventListener) {
            // W3C method
            obj.addEventListener(evt, handler, false);
        } else if (obj.attachEvent) {
            // IE method.
            obj.attachEvent('on' + evt, handler);
        } else {
            // Old school method.
            obj['on' + evt] = handler;
        }
        return true;
    } catch (err) {
        return false;
    }
}

$(document).ready(function () {
    $(document).on("change", "[name=files]", function (e) {
        if (e.target.id == "uploader-control-1")
            handleFileSelect(e);
        else
            handleFileSelectNext(e);
    })

    if (window.FileReader) {
        addEventHandler(window, 'load', function () {
            var drop = document.getElementById('form-upload-place-container');

            function cancel(e) {
                if (e.preventDefault) { e.preventDefault(); }
                return false;
            }

            //// Tells the browser that we *can* drop on this target
            addEventHandler(drop, 'dragover', cancel);
            addEventHandler(drop, 'dragenter', cancel);

            addEventHandler(drop, 'drop', function (e) {
                e = e || window.event; // get window.event if e argument missing (in IE)   
                if (e.preventDefault) { e.preventDefault(); } // stops the browser from redirecting off to the image.

                var dt = e.dataTransfer;
                var files = dt.files;
                sizeUploaded = 0;
                var container = document.getElementById("form-upload-uploader-status");
                var Status_Details = "<div class='uploader-progress' id='uploader-progress' style='display: none;'>"
                                        + "<span class='text'>0%</span>"
                                        + "<span class='percentage' style='width: 0%;'></span>"
                                   + "</div>"
                                   + "<div class='uploader-info'></div>"
                                   + "<div class='btn-uploader' id='form-upload-btn-uploader'>"
                                   + "<div class='uploader-container-ready'>"
                                        + "<div class='uploader-picker'>Tiếp tục bổ sung</div>"
                                        + "<div style='position: absolute; top: 0px; left: 9px; width: 143px; height: 42px; overflow: hidden; bottom: auto; right: auto;'>"
                                            + "<input type='file' name='files' class='uploader-control'id='uploader-control-2' multiple='multiple' />"
                                            + "<label style='opacity: 0; width: 100%; height: 100%; display: block; cursor: pointer; background: rgb(255, 255, 255); position: absolute;' for='uploader-control-2'></label>"
                                        + "</div>"
                                   + "</div>"
                                   + "</div>";
                if (options.button == true) {
                    var upButton = document.createElement("div");
                    upButton.setAttribute("class", "btn-ready");
                    upButton.id = "btnUpload";
                    upButton.innerText = "Bắt đầu tải lên";
                    var btnContainer = document.getElementById("form-upload-btn-uploader");
                    btnContainer.appendChild(upButton);
                }
                var queue = document.getElementById("form-upload-queue-inlist");
                document.getElementById("form-upload-place-container").style.display = "none";
                queue.setAttribute("class", "queue-inlist filled");
                container.innerHTML = Status_Details;
                var file_inlist = document.getElementById("form-upload-file-inlist");
                loopForRender(files);

                var successMssg = "<p class='isSuccess'>Đã tải lên thành công</p>",
                            errorMssg = "<p class='isError'>Tải lên không thành công</p>";
                var formData = new FormData(document.querySelector("#form-submit"));
                data = formData;
                options.data = data;
                for (var i = 0; i < files.length; i++) {
                    formData.append("filenames", files[i].name);
                    formData.append("files", files[i]);
                }
                var progress = document.getElementsByClassName("uploader-progress")[0];
                $.ajax({
                    url: options.url,
                    type: "POST",
                    data: options.data,
                    processData: false,  // tell jQuery not to process the data
                    contentType: false,
                    xhr: function () {
                        var xhr = new window.XMLHttpRequest();
                        //Upload progress
                        xhr.upload.addEventListener("progress", function (evt) {
                            if (evt.lengthComputable) {
                                var percentComplete = evt.loaded / evt.total;
                                var ratio = parseFloat(percentComplete * 100).toFixed(2);
                                //Do something with upload progress   
                                $(progress).css("display", "inline-block");
                                $(progress).find(".text").html(ratio + "%");
                                $(progress).find(".percentage").css("width", ratio + "%");
                            }
                        }, false);
                        //Download progress
                        xhr.upload.onloadstart = function (e) {
                            $(progress).find(".text").html("0%");
                            $(progress).find(".percentage").css("width", "0%");
                        }
                        xhr.upload.onloadend = function (e) {
                            var percentComplete = e.loaded / e.total;
                            var ratio = parseFloat(percentComplete * 100).toFixed(0);
                            $(progress).find(".text").html(ratio + "%");
                            $(progress).find(".percentage").css("width", ratio + "%");
                        }
                        return xhr;
                    },
                    success: function () {
                        $(".form-upload").append(successMssg);
                        setTimeout(function () {
                            $("p.isSuccess").remove();
                        }, 2000);
                    },
                    error: function () {
                        $(".form-upload").append(errorMssg);
                        setTimeout(function () {
                            $("p.isError").remove();
                        }, 2000);
                    },
                    complete: function () {
                        document.forms[0].reset();
                    }// tell jQuery not to set contentType           ,
                });
                return false;
            });
        });
    } else {
        alert('Your browser does not support the HTML5 FileReader.');
    }

    $(document).on("click", ".img-delete", function (e) {
        getUploadedInfo();
        var size = parseFloat($(this).data("size"));
        numberUploaded--;
        sizeUploaded -= size;
        $(".uploader-info").html("Đã chọn " + numberUploaded + " hình ảnh, dung lương " + parseFloat(sizeUploaded / 1048576).toFixed(2) + " MB");
        $(this).closest("li").remove();
    })


    $(document).on("click", "#btnUpload", function () {
        var successMssg = "<p class='isSuccess'>Đã tải lên thành công</p>",
            errorMssg = "<p class='isError'>Tải lên không thành công</p>";
        var formData = new FormData(document.querySelector("#form-submit"));
        data = formData;
        options.data = data;
        var progress = document.getElementsByClassName("uploader-progress")[0];
        $.ajax({
            url: options.url,
            type: "POST",
            data: options.data,
            processData: false,  // tell jQuery not to process the data
            contentType: false,
            xhr: function () {
                var xhr = new window.XMLHttpRequest();
                //Upload progress
                xhr.upload.addEventListener("progress", function (evt) {
                    if (evt.lengthComputable) {
                        var percentComplete = evt.loaded / evt.total;
                        var ratio = parseFloat(percentComplete * 100).toFixed(2);
                        //Do something with upload progress   
                        $(progress).css("display", "inline-block");
                        $(progress).find(".text").html(ratio + "%");
                        $(progress).find(".percentage").css("width", ratio + "%");
                    }
                }, false);
                //Download progress
                xhr.upload.onloadstart = function (e) {
                    $(progress).find(".text").html("0%");
                    $(progress).find(".percentage").css("width", "0%");
                }
                xhr.upload.onloadend = function (e) {
                    var percentComplete = e.loaded / e.total;
                    var ratio = parseFloat(percentComplete * 100).toFixed(0);
                    $(progress).find(".text").html(ratio + "%");
                    $(progress).find(".percentage").css("width", ratio + "%");
                }
                return xhr;
            },
            success: function () {
                $(".form-upload").append(successMssg);
                setTimeout(function () {
                    $("p.isSuccess").remove();
                }, 2000);
            },
            error: function (XMLHttpRequest, textStatus, errorThrown) {
                $(".form-upload").append(errorMssg);
                setTimeout(function () {
                    $("p.isError").remove();
                }, 2000);
            },
            complete: function () {
                document.forms[0].reset();
            }// tell jQuery not to set contentType           ,
        });
    })
})

var sizeUploaded = 0, numberUploaded = 0, nameUploaded = [];

function getUploadedInfo() {
    sizeUploaded = 0; numberUploaded = 0;
    $(".file-inlist a").each(function () {
        var size = $(this).data("size");
        sizeUploaded += parseFloat(size);
        numberUploaded++;
    })
}

function handleFileSelect(evt) {
    var container = document.getElementById("form-upload-uploader-status");
    var Status_Details = "<div class='uploader-progress' id='uploader-progress' style='display: none;'>"
                            + "<span class='text'>0%</span>"
                            + "<span class='percentage' style='width: 0%;'></span>"
                       + "</div>"
                       + "<div class='uploader-info'></div>"
                       + "<div class='btn-uploader' id='form-upload-btn-uploader'>"
                       + "<div class='uploader-container-ready'>"
                            + "<div class='uploader-picker'>Tiếp tục bổ sung</div>"
                            + "<div style='position: absolute; top: 0px; left: 9px; width: 143px; height: 42px; overflow: hidden; bottom: auto; right: auto;'>"
                                + "<input type='file' name='files' class='uploader-control'id='uploader-control-2' multiple='multiple' />"
                                + "<label style='opacity: 0; width: 100%; height: 100%; display: block; cursor: pointer; background: rgb(255, 255, 255); position: absolute;' for='uploader-control-2'></label>"
                            + "</div>"
                       + "</div>"
                       + "</div>";
    var queue = document.getElementById("form-upload-queue-inlist");
    document.getElementById("form-upload-place-container").style.display = "none";
    queue.setAttribute("class", "queue-inlist filled");
    container.innerHTML = Status_Details;
    if (options.button == true) {
        var upButton = document.createElement("div");
        upButton.setAttribute("class", "btn-ready");
        upButton.id = "btnUpload";
        upButton.innerText = "Bắt đầu tải lên";
        var btnContainer = document.getElementById("form-upload-btn-uploader");
        btnContainer.appendChild(upButton);
    }
    var files = evt.target.files; // FileList object    
    var file_inlist = document.getElementById("form-upload-file-inlist");
    loopForRender(files);
}

function loopForRender(files) {
    // Loop through the FileList and render image files as thumbnails.
    for (var i = 0, f; f = files[i]; i++) {

        // Only process image files.
        if (!f.type.match('image.*')) {
            continue;
        }
        var reader = new FileReader();

        // Closure to capture the file information.
        reader.onload = (function (theFile) {
            return function (e) {
                // Render thumbnail.
                var figure = document.createElement('figure');
                figure.style.marginLeft = "5px";
                figure.innerHTML = ['<img class="thumb" src="', URL.createObjectURL(theFile),
                                  '" title="', escape(theFile.name), '" style="width:100%" />',
                                 '<figcaption>', theFile.name, '</figcaption>',
                                 '<input type="hidden" name="filenames" value="', theFile.name, '"/>',
                                 '<a href="#" class="img-delete" data-size="', theFile.size, '">×</a>'].join('');
                var li = document.createElement('li');
                li.innerHTML = figure.outerHTML;
                document.getElementById('form-upload-file-inlist').insertBefore(li, null);
                nameUploaded.push(theFile.name);
                numberUploaded++;
                sizeUploaded += theFile.size;
                $(".uploader-info").html("Đã chọn " + numberUploaded + " hình ảnh, dung lương " + parseFloat(sizeUploaded / 1048576).toFixed(2) + " MB");
            };
        })(f);

        // Read in the image file as a data URL.
        reader.readAsDataURL(f);
    }
}

function handleFileSelectNext(evt) {
    getUploadedInfo()
    var addedFiles = [];
    $("figcaption").each(function () {
        addedFiles.push($(this).html().trim());
    })
    var that = $(evt.target),
        div = that.parent(),
        numberInput = $(".form-upload input[type=file]").length,
        files = evt.target.files, // FileList object   
        file_inlist = document.getElementById("form-upload-file-inlist");
    //Create new instant of input
    var id = "uploader-control-" + (numberInput + 1),
        input = that.clone().attr("id", id);
    div.append(input);
    div.find("label").attr("for", id);

    // Loop through the FileList and render image files as thumbnails.
    for (var i = 0, f; f = files[i]; i++) {

        // Only process image files.
        if (!f.type.match('image.*')) {
            continue;
        }
        var reader = new FileReader();

        // Closure to capture the file information.
        reader.onload = (function (theFile) {
            if ($.inArray(theFile.name, addedFiles) == -1) {
                return function (e) {
                    // Render thumbnail.
                    var figure = document.createElement('figure');
                    figure.style.marginLeft = "5px";
                    figure.innerHTML = ['<img class="thumb" src="', URL.createObjectURL(theFile),
                                      '" title="', escape(theFile.name), '" style="width:100%" />',
                                     '<figcaption>', theFile.name, '</figcaption>',
                                     '<input type="hidden" name="filenames" value="', theFile.name, '"/>',
                                     '<a href="#" class="img-delete" data-size="', theFile.size, '">×</a>'].join('');
                    var li = document.createElement('li');
                    li.innerHTML = figure.outerHTML;
                    document.getElementById('form-upload-file-inlist').insertBefore(li, null);
                    numberUploaded++;
                    sizeUploaded += theFile.size;
                    $(".uploader-info").html("Đã chọn " + numberUploaded + " hình ảnh, dung lương " + parseFloat(sizeUploaded / 1048576).toFixed(2) + " MB");
                };
            }
        })(f);

        // Read in the image file as a data URL.
        reader.readAsDataURL(f);
    }
}