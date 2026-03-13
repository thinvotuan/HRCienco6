var Alert = undefined;

(function (Alert) {    
    var alert, error, info, success, warning, _container;
    info = function (message, title, options) {
        return alert("info", message, title, "icon-info-sign", options);
    };
    warning = function (message, title, options) {
        return alert("warning", message, title, "glyphicon glyphicon-warning-sign", options);
    };
    error = function (message, title, options) {
        return alert("error", message, title, "glyphicon glyphicon-ban-circle", options);
    };
    success = function (message, title, options) {
        return alert("successful", message, title, "glyphicon glyphicon-ok", options);
    };
    alert = function (type, message, title, icon, options) {
        //var alerts = document.getElementById("alerts");
        //if (alerts != null) {
        //    $(alerts).fadeOut("fast");
        //    alerts.parentElement.removeChild(alerts);
        //}
        var boundnary = document.createElement("div"), ul = "<ul id='alerts'></ul>";
        boundnary.innerHTML = ul;
        document.body.appendChild(boundnary);
        var alertElem, messageElem, titleElem, iconElem, innerElem, _container;
        if (typeof options === "undefined") {
            options = {};
        }
        options = $.extend({}, Alert.defaults, options);
        if (!_container) {
            _container = $("#alerts");
            if (_container.length === 0) {
                _container = $("<ul>").attr("id", "alerts").appendTo($("body"));
            }
        }
        if (options.width) {
            _container.css({
                width: options.width
            });
        }
        alertElem = $("<li>").addClass("alerts").addClass("alerts-" + type);
        setTimeout(function () {
            alertElem.addClass('open-alert');
        }, 1);
        if (icon) {
            iconElem = $("<i>").addClass(icon);
            alertElem.append(iconElem);
        }
        innerElem = $("<div>").addClass("alerts-block");
        alertElem.append(innerElem);
        if (title) {
            titleElem = $("<div>").addClass("alerts-title").append(title);
            innerElem.append(titleElem);
        }
        if (message) {
            messageElem = $("<div>").addClass("alerts-message").append(message);
            innerElem.append(messageElem);
        }
        if (options.displayDuration > 0) {
            setTimeout((function () {
                leave();
            }), options.displayDuration);
        } else {
            innerElem.append("<em>Thoát</em>");
        }
        alertElem.on("click", function () {
            leave();
        });
        function leave() {
            alertElem.removeClass('open-alert');
            alertElem.one('webkitTransitionEnd otransitionend oTransitionEnd msTransitionEnd transitionend', function () { return alertElem.remove(); });
        }
        return _container.prepend(alertElem);
    };
    Alert.defaults = {
        width: "",
        icon: "",
        displayDuration: 3000,
        pos: ""
    };
    Alert.info = info;
    Alert.warning = warning;
    Alert.error = error;
    Alert.success = success;
    return _container = void 0;


})(Alert || (Alert = {}));

this.Alert = Alert;

$(document).ready(function () {
    $('#test').on('click', function () {
        Alert.info('Message');
    });

    $("form").bind("keypress", function (e) {
        if (e.keyCode == 13) return false;
    });
})

var bdsConfirm = undefined;
(function (bdsConfirm) {
    var confirm, alert;
    confirm = function (message, title) {
        var that = document.getElementById("modal-confirm")
        document.body.removeChild(that);
        var elementChilds = "<div class='confirm-header'>"
                            + "<a href='#' data-dismiss='modal' class='confirm-close' onclick='Cancel();'>×</a>"
                            + "<h5>" + title + "</h3>"
                            + "</div>"
                            + "<div class='confirm-body'>"
                                + "<p>" + message + ".</p>"
                            + "</div>"
                            + "<div class='confirm-footer'>"
                                + "<a href='#' id='btnConfirmYes' onclick='Accept();'>Đồng ý</a>"
                                + "<a href='#' id='btnConfirmNo' onclick='Cancel();'>Hủy</a>"
                         + "</div>";
        var div = document.createElement("div");
        div.id = "modal-confirm";
        document.body.appendChild(div)
        var parent = document.getElementById("modal-confirm").innerHTML = elementChilds;
        var boundnary = document.createElement("div");
        boundnary.id = "confirm-backdrop";
        document.body.appendChild(boundnary);
    };

    alert = function (message, title) {
        var modal_confirm = document.getElementById("modal-confirm");
        var confirm_backdrop = document.getElementById("confirm_backdrop");
        if (modal_confirm != null) {
            modal_confirm.parentElement.removeChild(that);
            confirm_backdrop.parentElement.removeChild(confirm_backdrop);
        }
        var elementChilds = "<div class='confirm-header'>"
                            + "<a href='#' class='confirm-close' onclick='Cancel();'>×</a>"
                            + "<h5>" + title + "</h3>"
                            + "</div>"
                            + "<div class='confirm-body'>"
                                + "<p>" + message + ".</p>"
                            + "</div>"
                            + "<div class='confirm-footer'>"
                                + "<a href='#' id='btnConfirmNo' onclick='Cancel();'>Đóng</a>"
                         + "</div>";
        var div = document.createElement("div");
        div.id = "modal-confirm";
        document.body.appendChild(div)
        var parent = document.getElementById("modal-confirm").innerHTML = elementChilds;
        var boundnary = document.createElement("div");
        boundnary.id = "confirm-backdrop";
        document.body.appendChild(boundnary)
    };
    bdsConfirm.confirm = confirm;
    bdsConfirm.alert = alert;
})(bdsConfirm || (bdsConfirm = {}));

this.bdsConfirm = bdsConfirm;

function Accept() {
    var child1 = document.getElementById("modal-confirm");
    var child2 = document.getElementById("confirm-backdrop");
    document.body.removeChild(child1);
    document.body.removeChild(child2);
    confirmTimer = true;
}
function Cancel(event) {
    var child1 = document.getElementById("modal-confirm");
    var child2 = document.getElementById("confirm-backdrop");
    document.body.removeChild(child1);
    document.body.removeChild(child2);
    confirmTimer = false
}
function detectEmail(email) {
    return email.split('@')[0];
}