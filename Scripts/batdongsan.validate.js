jQuery.extend(jQuery.validator.messages, {
    required: "Vui lòng nhập thông tin vào đây"
});

jQuery.validator.addMethod("greaterThanZero", function (value, element) {
    return this.optional(element) || (parseFloat(value) > 0);
}, "Value phải lớn hơn 0");

jQuery.validator.addMethod("nonWhiteSpace", function (value, element) {
    return this.optional(element) || (parseFloat(value) > 0) || value != "" || value != null;
}, "Value không được rỗng");

jQuery.validator.addMethod("equalOneHundred", function (value, element) {
    return this.optional(element) || parseFloat(value) == 100;
}, "Value phải bằng 100");

jQuery.extend($.validator.prototype, {
    checkForm: function () {
        this.prepareForm();
        for (var i = 0, elements = (this.currentElements = this.elements()) ; elements[i]; i++) {
            if (this.findByName(elements[i].name).length != undefined && this.findByName(elements[i].name).length > 1) {
                for (var cnt = 0; cnt < this.findByName(elements[i].name).length; cnt++) {
                    this.check(this.findByName(elements[i].name)[cnt]);
                }
            } else {
                this.check(elements[i]);
            }
        }
        return this.valid();
    }
});

jQuery(document).ready(function () {
    $("#btnSubmit").click(function (e) {
        $("[required=required],[required=true]").each(function () {
            if ($(this).val() != null) {
                $(this).val($(this).val().trim());
                $(this).parent().removeClass("error");
            }
        });

        $("form").bind("keypress", function (e) {
            if (e.keyCode == 13) return false;
        });

        $("#form-submit").validate({
            ignore: ".ignore",
            rules: {
                required: true,
                idTinhThanh: {
                    greaterThanZero: true
                },
                idNgheNghiep: {
                    greaterThanZero: true
                },
                nonWhiteSpace: {
                    nonWhiteSpace: true
                },
                sumTyLe: {
                    equalOneHundred: true
                },
                idPhuongThucTT: {
                    greaterThanZero: true
                },
                greaterThanZero: {
                    greaterThanZero: true
                },
                idThoiHanHopDong: {
                    greaterThanZero: true
                },
                notZero: {
                    equalOneHundred: true
                }
            },
            errorPlacement: function (error, element) {
                $(error).remove();
                if ($(element).is("input[type=radio], input[type=checkbox]")) {
                    $(element).parent().addClass("error");
                }
            },

            invalidHandler: function (form, validator) {

                if (!validator.numberOfInvalids())
                    return;

                var element = $(validator.errorList[0].element);
                var parent = element.parents(".tab-pane")[0];
                if (parent != undefined) {
                    var id = parent.id;
                    $("a[href=#" + id + "]").tab('show');
                }

                //$('html, body').animate({
                //    scrollTop: $(validator.errorList[0].element).offset().top
                //}, 1000);
            },

            submitHandler: function (form) {
                $(".numberinput").each(function () {
                    $(this).val($(this).autoNumeric('get'));
                });

                $(".numberinputInt").each(function () {
                    $(this).val($(this).autoNumeric('get'));
                });
                if (form.method == "post") {
                    $(".pageLoading").show();

                    $(".btnSubmit, #btnSubmit").replaceWith("Đang xử lý...");
                    form.submit();
                } else {
                    return false;
                }
                
            }
        });
    });
});

