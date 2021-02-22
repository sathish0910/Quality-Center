QCNewForm = { vars: { QCNewForm: [] }, fn: {} };
QCNewForm.vars.LM = [];
QCNewForm.vars.webUrl = _spPageContextInfo.webAbsoluteUrl;
QCNewForm.vars.currentUserID = _spPageContextInfo.userId;
QCNewForm.vars.currentUsers = [];
QCNewForm.vars.waitDialog = null;
QCNewForm.vars.btLogo = "";
QCNewForm.vars.spLogo = "";

// Summary : On load function
$(document).ready(function () {
    QCNewForm.fn.convertImgBase64Formate();
    QCNewForm.fn.loadForm(); //hilly billy functionality
    QCNewForm.vars.requestType = GetUrlKeyValue('requestType');
    QCNewForm.fn.common(); //common functionalities
    QCNewForm.fn.setFormType(QCNewForm.vars.requestType);
    QCNewForm.fn.initilizeFunctions(); //Click or Change functions
});

QCNewForm.fn.loadForm = function () {
    $('span.hillbillyForm').each(function () {
        //get the display name from the custom layout
        QCNewForm.vars.displayName = $(this).attr("data-displayName");
        QCNewForm.vars.displayName = QCNewForm.vars.displayName.replace(/&(?!amp;)/g, '&amp;');
        QCNewForm.vars.elem = $(this);
        //find the corresponding field from the default form and move it 
        //into the custom layout
        $("table.ms-formtable td").each(function () {
            if (this.innerHTML.indexOf('FieldInternalName="' + QCNewForm.vars.displayName + '"') != -1) {
                $(this).contents().appendTo(QCNewForm.vars.elem);
            }
        });
    });
    QCNewForm.fn.getUserDetails(QCNewForm.vars.webUrl, QCNewForm.vars.currentUserID).done(function (data) {
        QCNewForm.vars.currentUsers = data;
        QCNewForm.vars.splLogNam = QCNewForm.vars.currentUsers.Email.split("@")
        if(QCNewForm.vars.requestType === "BTUser") {
            $("input[id^='LoginName']").val(QCNewForm.vars.splLogNam[0])
        }
    });
}

//Switch case to set fields
QCNewForm.fn.setFormType = function (type) {
    switch (type) {
        case "BTUser":
            QCNewForm.fn.btUserForm();
            break;
        case "Additional":
            QCNewForm.fn.additionalAccessForm();
            break;
        case "Reactivation":
            QCNewForm.fn.reactivationForm();
            break;
        case "Unlocking":
            QCNewForm.fn.unlockingForm();
            break;
    }
}
// hide & show the required fields in form
/* <-- start --> */
QCNewForm.fn.btUserForm = function () {
    $('.qc .header').html('New BT User');
    $("input[id^='RequestType'][value='New BT User']").prop("checked", true);
    $("input[id^='LoginName']").attr('readonly', true);
    $("input[id^='DefaultAccess']").parent().parent().parent().parent().show();
    $("select[id^='AccessRequest']").parent().parent().parent().parent().show();
    $("select[id^='ProjectRole']").parent().parent().parent().parent().show();
}

QCNewForm.fn.additionalAccessForm = function () {
    $('.qc .header').html('Additional QC Access');
    $("input[id^='RequestType'][value='Additional QC Access']").prop("checked", true);
    $("select[id^='AccessRequest']").parent().parent().parent().parent().show();
    $("select[id^='ProjectRole']").parent().parent().parent().parent().show();
}

QCNewForm.fn.reactivationForm = function () {
    $('.qc .header').html('Account Re-activation');
    $("input[id^='RequestType'][value='Account Re-activation']").prop("checked", true);
    $("input[id^='FirstName']").parent().parent().parent().parent().hide();
    $("input[id^='SecondName']").parent().parent().parent().parent().hide();
}

QCNewForm.fn.unlockingForm = function () {
    $('.qc .header').html('Unlocking Your QC account');
    $("input[id^='RequestType'][value='Unlocking Your QC account']").prop("checked", true);
    $("input[id^='FirstName']").parent().parent().parent().parent().hide();
    $("input[id^='SecondName']").parent().parent().parent().parent().hide();
}

QCNewForm.fn.common = function () {
    $('table[id^="RequestType"]').parent().parent().parent().parent().hide();
    $("input[id^='LoginName']").attr('placeholder', 'Enter your login name');   
    $("input[id^='FirstName']").attr('readonly', true);
    $("input[id^='SecondName']").attr('readonly', true);
    $("input[id^='Title']").attr('readonly', true);
    $('select[id^="ProjectRole"][id$="LookupField"]').attr("disabled", true);
    $("input[id^='DefaultAccess']").parent().parent().parent().parent().hide();
    $("select[id^='AccessRequest']").parent().parent().parent().parent().hide();
    $("select[id^='ProjectRole']").parent().parent().parent().parent().hide();
}
/// hide & show the required fields in form
/* <-- end --> */

QCNewForm.fn.initilizeFunctions = function () {
    var EAPP = SPClientPeoplePicker.SPClientPeoplePickerDict["EmailAddress_ac60445f-b9d5-4c59-a7bb-5786a69259fc_$ClientPeoplePicker"];
    EAPP.OnValueChangedClientScript = function (peoplePickerId, selectedUsersInfo) {
        setTimeout(function () {
            if (selectedUsersInfo.length > 0) {
                if (selectedUsersInfo[0].EntityData != undefined) {
                    QCNewForm.vars.LM = selectedUsersInfo;
                    if (selectedUsersInfo[0].DisplayText) {
                        var selectedUser = selectedUsersInfo[0].DisplayText;
                    }
                    if (selectedUsersInfo[0].AutoFillDisplayText) {
                        var selectedUser = selectedUsersInfo[0].DisplayText;
                    }
                    if (selectedUser) {
                        if (selectedUser[selectedUser.length - 1].indexOf("R") > -1) {
                            // user details get function goes here
                            QCNewForm.fn.getAndBindUserDetails(QCNewForm.vars.LM[0].Key);
                        } else {
                            alert('User is a sub con! Please enter BT/EE.');
                            var usersobject = EAPP.GetAllUserInfo();
                            usersobject.forEach(function (index) {
                                EAPP.DeleteProcessedUser(EAPP[index]);
                            });
                            QCNewForm.vars.LM = [];
                        }
                    }
                }
            }
        }, 1000);
        $("a[id^='EmailAddress']").click(function () {
            $("input[id^='FirstName']").val('');
            $("input[id^='SecondName']").val('');
            $("input[id^='Title']").val('');
        })
    };
    $('select[id^="AccessRequest"][id$="LookupField"]').change(function () {
        $('select[id^="ProjectRole"][id$="LookupField"]').val($(this).val());
    });
}

/**************************************************************************/
/// Function name: get SP.UserProfiles.PeopleManager
/// Summary:  To get user id in sharepoint site by user accountname
/**************************************************************************/
QCNewForm.fn.getAndBindUserDetails = function (accountName) {
    QCNewForm.fn.getUserIdByAccountName(QCNewForm.vars.webUrl, accountName).done(function (data) {
        console.log(data);
    });
}
QCNewForm.fn.getUserIdByAccountName = function (url, accountName) {
    var dfd = $.Deferred();
    $.ajax({
        url: url + "/_api/SP.UserProfiles.PeopleManager/GetPropertiesFor(accountName=@v)?@v='" + encodeURIComponent(accountName) + "'",
        method: "GET",
        headers: { "Accept": "application/json; odata=verbose" },
        success: function (data) {
            dfd.resolve(data.d.Id);
            for (var i = 0; i < data.d.UserProfileProperties.results.length; i++) {
                var item = data.d.UserProfileProperties.results[i];
                if (item.Key === "PreferredName") {
                    $('input[id^="Title"]').val(item.Value);
                    QCNewForm.vars.requestorTitle = item.Value;
                }
                if (item.Key === "FirstName")
                    $('input[id^="FirstName"]').val(item.Value);
                if (item.Key === "LastName")
                    $('input[id^="SecondName"]').val(item.Value);
                QCNewForm.vars.requestorEmailID = data.d.Email;
            }
        },
        error: function (data) {
            dfd.reject(JSON.stringify(data));
        }
    });
    return dfd.promise();
};

/**************************************************************************/
/// Function name: getuserbyID
/// Summary:  To get login details in sharepoint site by user ID
/// Current user deatils email ID, Login Name
/**************************************************************************/
QCNewForm.fn.getUserDetails = function (url, userid) {
    var dfd = $.Deferred();
    $.ajax({
        url: url + "/_api/web/getuserbyid(" + userid + ")",
        method: "GET",
        headers: { "Accept": "application/json; odata=verbose" },
        success: function (data) {
            dfd.resolve(data.d);
        },
        error: function (data) {
            dfd.reject(JSON.stringify(data));
        }
    });
    return dfd.promise();
};

// Summary: Valid mandatory fields
QCNewForm.fn.checkValidation = function () {
    QCNewForm.vars.valid = false;
    var mandatoryTextField = [];
    var mandatorySelectField = [];
    if (QCNewForm.vars.requestType === "BTUser") {
        mandatoryTextField = ["First Name", "Second Name", "Full Name", "Default Access", "Login Name"];
        mandatorySelectField = ["Access Request", "Project Role"];
    }
    else if (QCNewForm.vars.requestType === "Additional") {
        mandatoryTextField = ["First Name", "Second Name", "Full Name", "Project Role", "Login Name"];
        mandatorySelectField = ["Access Request", "Project Role"];
    }
    else if (QCNewForm.vars.requestType === "Reactivation" || QCNewForm.vars.requestType === "Unlocking") {
        mandatoryTextField = ["Full Name", "Login Name"];
    }

    $.each($('input[type="text"]'), function (index, val) {
        var returnInputVal = jQuery.grep(mandatoryTextField, function (a) {
            return a === val.title;
        });
        if (returnInputVal.length > 0) {
            if ($(val).val() === "") {
                $(val).css('border', '1px solid #ff0000');
                QCNewForm.vars.valid = true;
            } else {
                $(val).css('border', '1px solid #ababab');
            }
        }
    });
    $.each($('select'), function (index, val) {
        var returnVal = jQuery.grep(mandatorySelectField, function (a) {
            return a === val.title;
        });
        if (returnVal.length > 0) {
            if ($(val).val() === "" || $(val).val() === null) {
                $(val).css('border', '1px solid #ff0000');
                QCNewForm.vars.valid = true;
            } else if (($("select[ID^='AccessRequest'] :selected").text() === "(None)") && (QCNewForm.vars.requestType === "BTUser")) {
                $(val).css('border', '1px solid #ff0000');
                QCNewForm.vars.valid = true;
            } else if (($("select[ID^='ProjectRole'] :selected").text() === "(None)") && (QCNewForm.vars.requestType === "BTUser")) {
                $(val).css('border', '1px solid #ff0000');
                QCNewForm.vars.valid = true;
            } else if ($("select[ID^='AccessRequest'] :selected").text() === "(None)" && QCNewForm.vars.requestType === "Additional") {
                $(val).css('border', '1px solid #ff0000');
                QCNewForm.vars.valid = true;
            } else if ($("select[ID^='ProjectRole'] :selected").text() === "(None)" && QCNewForm.vars.requestType === "Additional") {
                $(val).css('border', '1px solid #ff0000');
                QCNewForm.vars.valid = true;
            } else {
                $(val).css('border', '1px solid #ababab');
            }
        }
    });
    return QCNewForm.vars.valid
}
// Summary:  To show OOB SP wait dialog
QCNewForm.fn.ShowWaitDialog = function (title, message) {
    if (QCNewForm.vars.waitDialog === null) {
        QCNewForm.vars.waitDialog = (typeof (title) !== "undefined" && title !== '') ? SP.UI.ModalDialog.showWaitScreenWithNoClose(title, message) :
            SP.UI.ModalDialog.showWaitScreenWithNoClose(SP.Res.dialogLoading15);
    }
};
// Summary:  To hide OOB SP wait dialog
QCNewForm.fn.HideWaitDialog = function () {
    if (QCNewForm.vars.waitDialog !== null) {
        QCNewForm.vars.waitDialog.close(SP.UI.DialogResult.OK);
        QCNewForm.vars.waitDialog = null;
    }
};

//submit new form
QCNewForm.fn.createRequest = function () {
    var checkSubmit = QCNewForm.fn.checkValidation();
    QCNewForm.fn.setInputVal();
    if (!checkSubmit) {
        QCNewForm.fn.ShowWaitDialog("Please Wait!", "We are submitting your data into SharePoint QC database.");
        var emailTemplate = QCNewForm.fn.emailTemplate("Requestor", QCNewForm.vars.requestType, $("select[id^='Status'] :selected").text(), QCNewForm.vars.currentUsers.LoginName);
        InjectScript.fn.sendAnEmail([QCNewForm.vars.currentUsers.Email], [], emailTemplate, "Quality Center '" + $('input[name^="RequestType"]:checked').val() + "' Request form has been created by - " + QCNewForm.vars.currentUsers.Title + "").done(function (data) {
            QCNewForm.fn.HideWaitDialog();
            $($('[id$="SaveItem"]')[1]).trigger('click');
            console.log("Data Submitted");
        });
    }
    else {
        alert("Please fill the all mandatory fields.");
    }
}

//Cancel new form
QCNewForm.fn.cancelForm = function () {
    location.href = QCNewForm.vars.webUrl;
}

//Fetch Input form values
QCNewForm.fn.setInputVal = function () {
    QCNewForm.vars.firstName = $('input[id^="FirstName"]').val();
    QCNewForm.vars.secondName = $('input[id^="SecondName"]').val();
    QCNewForm.vars.fullName = $('input[id^="Title"]').val();
    QCNewForm.vars.defaultAccess = $('input[id^="DefaultAccess"]').val();
    QCNewForm.vars.accessRequest = $('select[id^="AccessRequest"] :selected').text();
}

QCNewForm.fn.convertImgBase64Formate = function () {
    InjectScript.fn.toDataURL("/sites/QCPortal/Resources/Common/Images/BTLogo.png", function (base64Img) {
        QCNewForm.vars.btLogo = base64Img;
    });
    InjectScript.fn.toDataURL("/sites/QCPortal/Resources/Common/Images/SPLogo.png", function (base64Img) {
        QCNewForm.vars.spLogo = base64Img;
    });
}

QCNewForm.fn.emailTemplate = function (key, type, status, userLoginName) {
    if (key === "Requestor") {
        return reqEmailFormat = '<!DOCTYPE html>' +
            '<html>' +
            '  <body>' +
            '      <table style="width: 100%;background: #5514b4;border: 1px solid #5514b4;">' +
            '          <tr>' +
            '              <td style="background: #5514b4;"><img width="40" height="40" src="' + QCNewForm.vars.btLogo + '">' +
            '              <td style="text-align: center;color: white;background: #5514b4;"><span style="text-align: center;">Quality Central Request -  ' + $('input[name^="RequestType"]:checked').val() + '</span></td>' +
            '              <td style="text-align: right;background: #5514b4;"><img width="80" height="40" src="' + QCNewForm.vars.spLogo + '"></td>' +
            '          </tr>' +
            '      </table>' +
            '      <table style="width: 100%;border: 1px solid #5514b4;">' +
            '          <tr>' +
            '              <td colspan="3" style="padding: 10px;">' +
            '                  <p>Hi ' + QCNewForm.vars.currentUsers.Title + ',</p>' +
            '                  <p>This is an auto generated SharePoint email.</p>' +
            '                  <p>Your request has been successfully created in Quality Central Portal. SME team will shortly responses to your request.</p>' +
            '              </td>' +
            '          </tr>' +
            '          <tr style="margin-top: 10px;margin-bottom: 30px;">' +
            '          	<td colspan="3" style="padding: 10px;"><span><b>Thanks & Regards,</b></span><br /><span>Quality Center Team</span></td>' +
            '      </tr>' +
            '      </table>' +
            '    <table style="width: 100%;background: #5514b4;border: 1px solid #5514b4;">' +
            '        <tr>' +
            '            <td style="text-align: center;color: white;"><span></span></td>' +
            '        </tr>' +
            '    </table>' +
            '  </body>' +
            '</html>';
    }
}
}