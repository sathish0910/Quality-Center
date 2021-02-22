// Summary      : SME dashboard, Get, Updating items from/to SharePoint List
// Ceated by    : E Sathish
// Modified by  : E Sathish
// Modified     : 27-April-2020
QCUpdateForm = { vars: { QCUpdateForm: [] }, fn: {} };
QCUpdateForm.vars.webUrl = _spPageContextInfo.webAbsoluteUrl;
QCUpdateForm.vars.restBody = [];

$(document).ready(function () {
    QCUpdateForm.fn.loadForm(); //hilly billy functionality
    QCUpdateForm.fn.common(); //common functionalities
    QCUpdateForm.fn.getRequests(); //get all requests from SP list
    QCUpdateForm.fn.initilizeFunctions(); //Click or Change functions

});

QCUpdateForm.fn.loadForm = function () {
    $('Title').html("Quality Center Requests");
    $('#pending-tab').css({ "background-color": "transparent", "color": "#337ab7" });
    $('#rejected-tab').css({ "background-color": "#ff0000", "color": "#000" });
    $('#approved-tab').css({ "background-color": "#008000", "color": "#000" });
    $('#completed-tab').css({ "background-color": "#00ffff", "color": "#000" });
}

QCUpdateForm.fn.common = function () {
    $("#pending-tab").click(function () {
        $('#pending-tab').css({ "background-color": "transparent", "color": "#337ab7" });
        $('#rejected-tab').css({ "background-color": "#ff0000", "color": "#000" });
        $('#approved-tab').css({ "background-color": "#008000", "color": "#000" });
        $('#completed-tab').css({ "background-color": "#00ffff", "color": "#000" });
    });
    $("#rejected-tab").click(function () {
        $('#pending-tab').css({ "background-color": "#ffff00", "color": "#000" });
        $('#rejected-tab').css({ "background-color": "transparent", "color": "#337ab7" });
        $('#approved-tab').css({ "background-color": "#008000", "color": "#000" });
        $('#completed-tab').css({ "background-color": "#00ffff", "color": "#000" });
    });
    $("#approved-tab").click(function () {
        $('#pending-tab').css({ "background-color": "#ffff00", "color": "#000" });
        $('#rejected-tab').css({ "background-color": "#ff0000", "color": "#000" });
        $('#approved-tab').css({ "background-color": "transparent", "color": "#337ab7" });
        $('#completed-tab').css({ "background-color": "#00ffff", "color": "#000" });
    });
    $("#completed-tab").click(function () {
        $('#pending-tab').css({ "background-color": "#ffff00", "color": "#000" });
        $('#rejected-tab').css({ "background-color": "#ff0000", "color": "#000" });
        $('#approved-tab').css({ "background-color": "#008000", "color": "#000" });
        $('#completed-tab').css({ "background-color": "transparent", "color": "#337ab7" });
    });
}

QCUpdateForm.fn.getRequests = function () {
    var query = "?$select=ID,RequestType,Title,Created,DefaultAccess,Status,BotStatus,Comments,AccessRequest/Title,ProjectRole/ProjectRoles,Author/Title&$expand=AccessRequest,ProjectRole,Author";
    InjectScript.getListItems(QCUpdateForm.vars.webUrl, "Quality Center Requests", query, true).done(function (data) {
        QCUpdateForm.vars.allRequests = data;
    });
    QCUpdateForm.fn.manageRequestType(QCUpdateForm.vars.allRequests);
    QCUpdateForm.fn.manageBOTViews(QCUpdateForm.vars.allRequests);
    QCUpdateForm.fn.appendListItems(QCUpdateForm.vars.allRequests);
}

QCUpdateForm.fn.initilizeFunctions = function () {
    $(".approveBtn").click(function () {
        var ItemId = $(this).attr('itemID');
        QCUpdateForm.vars.restBody.push({
            "Status": "Approved",
            "BotStatus": "SME Approved"
        });
        QCUpdateForm.fn.bindFormDetails(ItemId, QCUpdateForm.vars.restBody[0]);
        alert(QCUpdateForm.vars.restBody[0].Status + ", Thank you for approving the request!!!");
        location.reload();
    });
    $(".rejectBtn").click(function () {
        var ItemId = $(this).attr('itemID');
        $('#rejCmtsBtn').click(function () {
            if ($('#rejectComments').val() === "" || $('#rejectComments').val() === null) {
                alert("Comments are mandatory, when you the request is 'Rejected'");
            } else {
                QCUpdateForm.vars.restBody.push({
                    "Status": "Rejected",
                    "BotStatus": "SME Rejected",
                    "Comments": $('#rejectComments').val()
                });
                QCUpdateForm.fn.bindFormDetails(ItemId, QCUpdateForm.vars.restBody[0]);
                alert(QCUpdateForm.vars.restBody[0].Status + ", We are updating the request in rejected with Comments!!!");
                location.reload();
            }
        })
    });
}

QCUpdateForm.fn.manageRequestType = function (data) {
    $.each(data, function (index, val) {
        switch (val.Status) {
            case "Inprogress":
                $('#pending-btuser-tab').html('New BT User (' + QCUpdateForm.fn.returnFilterData("New BT User", val.Status).length + ')');
                $('#pending-additional-tab').html('Additional QC Access (' + QCUpdateForm.fn.returnFilterData("Additional Access", val.Status).length + ')');
                $('#pending-reactivation-tab').html('Account Re-activation (' + QCUpdateForm.fn.returnFilterData("Reactivation", val.Status).length + ')');
                $('#pending-unlock-account-tab').html('Unlocking Your QC account (' + QCUpdateForm.fn.returnFilterData("Account Unlock", val.Status).length + ')');
                break;
            case "Rejected":
                $('#rejected-btuser-tab').html('New BT User (' + QCUpdateForm.fn.returnFilterData("New BT User", val.Status).length + ')');
                $('#rejected-additional-tab').html('Additional QC Access (' + QCUpdateForm.fn.returnFilterData("Additional Access", val.Status).length + ')');
                $('#rejected-reactivation-tab').html('Account Re-activation (' + QCUpdateForm.fn.returnFilterData("Reactivation", val.Status).length + ')');
                $('#rejected-unlock-account-tab').html('Unlocking Your QC account (' + QCUpdateForm.fn.returnFilterData("Account Unlock", val.Status).length + ')');
                break;
            case "Approved":
                $('#approved-btuser-tab').html('New BT User (' + QCUpdateForm.fn.returnFilterData("New BT User", val.Status).length + ')');
                $('#approved-additional-tab').html('Additional QC Access (' + QCUpdateForm.fn.returnFilterData("Additional Access", val.Status).length + ')');
                $('#approved-reactivation-tab').html('Account Re-activation (' + QCUpdateForm.fn.returnFilterData("Reactivation", val.Status).length + ')');
                $('#approved-unlock-account-tab').html('Unlocking Your QC account (' + QCUpdateForm.fn.returnFilterData("Account Unlock", val.Status).length + ')');
        }
    });
}
QCUpdateForm.fn.manageBOTViews = function (data) {
    $.each(data, function (index, key) {
        switch (key.BotStatus) {
            case "Completed":
                $('#completed-btuser-tab').html('New BT User (' + QCUpdateForm.fn.returnFilterData("New BT User", key.BotStatus).length + ')');
                $('#completed-additional-tab').html('Additional QC Access (' + QCUpdateForm.fn.returnFilterData("Additional Access", key.BotStatus).length + ')');
                $('#completed-reactivation-tab').html('Account Re-activation (' + QCUpdateForm.fn.returnFilterData("Reactivation", key.BotStatus).length + ')');
                $('#completed-unlock-account-tab').html('Unlocking Your QC account (' + QCUpdateForm.fn.returnFilterData("Account Unlock", key.BotStatus).length + ')');
        }
    });
}

QCUpdateForm.fn.returnFilterData = function (reqType, status) {
    switch (status) {
        case "Completed":
            return inProgressRequests = _.filter(QCUpdateForm.vars.allRequests, function (temp) {
                return temp.RequestType === reqType;
            });
        default:
            return inProgressRequests = _.filter(QCUpdateForm.vars.allRequests, function (temp) {
                return temp.Status === status && temp.RequestType === reqType;
            });
    }
}

QCUpdateForm.fn.appendListItems = function () {
    //Pending Items
    var pendingBTUser = QCUpdateForm.fn.returnFilterData("New BT User", "Inprogress");
    var pendingAdditional = QCUpdateForm.fn.returnFilterData("Additional Access", "Inprogress");
    var pendingReactivation = QCUpdateForm.fn.returnFilterData("Reactivation", "Inprogress");
    var pendingUnlockAccount = QCUpdateForm.fn.returnFilterData("Account Unlock", "Inprogress");
    $('#pending-btuser tbody').append(_.map(pendingBTUser, function (val, index) {
        return "'<tr>" +
            "<td>" + val.ID + "</td>" +
            "<td>" + val.Author.Title + "</td>" +
            "<td>" + val.Created + "</td>" +
            "<td>" + val.DefaultAccess + "</td>" +
            "<td>" + val.AccessRequest.Title + "</td>" +
            "<td>" + val.ProjectRole.ProjectRoles + "</td>" +
            "<td>" + val.Status + "</td>" +
            "<td><input type='button' itemID='" + val.ID + "' class='approveBtn' value='Approve' /><input type='button' itemID='" + val.ID + "' class='rejectBtn' value='Reject' data-toggle='modal' data-target='#commentsModal' /></td>" +
            "</tr>";
    }).join(""));
    $('#pending-additional tbody').append(_.map(pendingAdditional, function (val, index) {
        return "'<tr>" +
            "<td>" + val.ID + "</td>" +
            "<td>" + val.Author.Title + "</td>" +
            "<td>" + val.Created + "</td>" +
            "<td>" + val.DefaultAccess + "</td>" +
            "<td>" + val.AccessRequest.Title + "</td>" +
            "<td>" + val.ProjectRole.ProjectRoles + "</td>" +
            "<td>" + val.Status + "</td>" +
            "<td><input type='button' itemID='" + val.ID + "' class='approveBtn' value='Approve' /><input type='button' itemID='" + val.ID + "' class='rejectBtn' value='Reject' data-toggle='modal' data-target='#commentsModal' /></td>" +
            "</tr>";
    }).join(""));
    $('#pending-reactivation tbody').append(_.map(pendingReactivation, function (val, index) {
        return "'<tr>" +
            "<td>" + val.ID + "</td>" +
            "<td>" + val.Author.Title + "</td>" +
            "<td>" + val.Created + "</td>" +
            "<td>" + val.Status + "</td>" +
            "</tr>";
    }).join(""));
    $('#pending-unlock-account tbody').append(_.map(pendingUnlockAccount, function (val, index) {
        return "'<tr>" +
            "<td>" + val.ID + "</td>" +
            "<td>" + val.Author.Title + "</td>" +
            "<td>" + val.Created + "</td>" +
            "<td>" + val.Status + "</td>" +
            "</tr>";
    }).join(""));

    //Rejected Items
    var rejectedBTUser = QCUpdateForm.fn.returnFilterData("New BT User", "Rejected");
    var rejectedAdditional = QCUpdateForm.fn.returnFilterData("Additional Access", "Rejected");
    var rejectedReactivation = QCUpdateForm.fn.returnFilterData("Reactivation", "Rejected");
    var rejectedUnlockAccount = QCUpdateForm.fn.returnFilterData("Account Unlock", "Rejected");
    $('#rejected-btuser tbody').append(_.map(rejectedBTUser, function (val, index) {
        return "'<tr>" +
            "<td>" + val.ID + "</td>" +
            "<td>" + val.Author.Title + "</td>" +
            "<td>" + val.Created + "</td>" +
            "<td>" + val.DefaultAccess + "</td>" +
            "<td>" + val.AccessRequest.Title + "</td>" +
            "<td>" + val.ProjectRole.ProjectRoles + "</td>" +
            "<td>" + val.Status + "</td>" +
            "<td>" + val.Comments + "</td>" +
            "</tr>";
    }).join(""));
    $('#rejected-additional tbody').append(_.map(rejectedAdditional, function (val, index) {
        return "'<tr>" +
            "<td>" + val.ID + "</td>" +
            "<td>" + val.Author.Title + "</td>" +
            "<td>" + val.Created + "</td>" +
            "<td>" + val.DefaultAccess + "</td>" +
            "<td>" + val.AccessRequest.Title + "</td>" +
            "<td>" + val.ProjectRole.ProjectRoles + "</td>" +
            "<td>" + val.Status + "</td>" +
            "<td>" + val.Comments + "</td>" +
            "</tr>";
    }).join(""));
    $('#rejected-reactivation tbody').append(_.map(rejectedReactivation, function (val, index) {
        return "'<tr>" +
            "<td>" + val.ID + "</td>" +
            "<td>" + val.Author.Title + "</td>" +
            "<td>" + val.Created + "</td>" +
            "<td>" + val.Status + "</td>" +
            "<td>" + val.Comments + "</td>" +
            "</tr>";
    }).join(""));
    $('#rejected-unlock-account tbody').append(_.map(rejectedUnlockAccount, function (val, index) {
        return "'<tr>" +
            "<td>" + val.ID + "</td>" +
            "<td>" + val.Author.Title + "</td>" +
            "<td>" + val.Created + "</td>" +
            "<td>" + val.Status + "</td>" +
            "<td>" + val.Comments + "</td>" +
            "</tr>";
    }).join(""));

    //Approved Items
    var approvedBTUser = QCUpdateForm.fn.returnFilterData("New BT User", "Approved");
    var approvedAdditional = QCUpdateForm.fn.returnFilterData("Additional Access", "Approved");
    var approvedReactivation = QCUpdateForm.fn.returnFilterData("Reactivation", "Approved");
    var approvedUnlockAccount = QCUpdateForm.fn.returnFilterData("Account Unlock", "Approved");
    $('#approved-btuser tbody').append(_.map(approvedBTUser, function (val, index) {
        return "'<tr>" +
            "<td>" + val.ID + "</td>" +
            "<td>" + val.Author.Title + "</td>" +
            "<td>" + val.Created + "</td>" +
            "<td>" + val.DefaultAccess + "</td>" +
            "<td>" + val.AccessRequest.Title + "</td>" +
            "<td>" + val.ProjectRole.ProjectRoles + "</td>" +
            "<td>" + val.Status + "</td>" +
            "</tr>";
    }).join(""));
    $('#approved-additional tbody').append(_.map(approvedAdditional, function (val, index) {
        return "'<tr>" +
            "<td>" + val.ID + "</td>" +
            "<td>" + val.Author.Title + "</td>" +
            "<td>" + val.Created + "</td>" +
            "<td>" + val.DefaultAccess + "</td>" +
            "<td>" + val.AccessRequest.Title + "</td>" +
            "<td>" + val.ProjectRole.ProjectRoles + "</td>" +
            "<td>" + val.Status + "</td>" +
            "</tr>";
    }).join(""));
    $('#approved-reactivation tbody').append(_.map(approvedReactivation, function (val, index) {
        return "'<tr>" +
            "<td>" + val.ID + "</td>" +
            "<td>" + val.Author.Title + "</td>" +
            "<td>" + val.Created + "</td>" +
            "<td>" + val.Status + "</td>" +
            "</tr>";
    }).join(""));
    $('#approved-unlock-account tbody').append(_.map(approvedUnlockAccount, function (val, index) {
        return "'<tr>" +
            "<td>" + val.ID + "</td>" +
            "<td>" + val.Author.Title + "</td>" +
            "<td>" + val.Created + "</td>" +
            "<td>" + val.Status + "</td>" +
            "</tr>";
    }).join(""));

    //Completed Bot Items
    var completedBTUser = QCUpdateForm.fn.returnFilterData("New BT User", "Completed");
    var completedAdditional = QCUpdateForm.fn.returnFilterData("Additional Access", "Completed");
    var completedReactivation = QCUpdateForm.fn.returnFilterData("Reactivation", "Completed");
    var completedUnlockAccount = QCUpdateForm.fn.returnFilterData("Account Unlock", "Completed");
    $('#completed-btuser tbody').append(_.map(completedBTUser, function (val, index) {
        return "'<tr>" +
            "<td>" + val.ID + "</td>" +
            "<td>" + val.Author.Title + "</td>" +
            "<td>" + val.Created + "</td>" +
            "<td>" + val.DefaultAccess + "</td>" +
            "<td>" + val.AccessRequest.Title + "</td>" +
            "<td>" + val.ProjectRole.ProjectRoles + "</td>" +
            "<td>" + val.Status + "</td>" +
            "<td>" + val.BotStatus + "</td>" +
            "</tr>";
    }).join(""));
    $('#completed-additional tbody').append(_.map(completedAdditional, function (val, index) {
        return "'<tr>" +
            "<td>" + val.ID + "</td>" +
            "<td>" + val.Author.Title + "</td>" +
            "<td>" + val.Created + "</td>" +
            "<td>" + val.DefaultAccess + "</td>" +
            "<td>" + val.AccessRequest.Title + "</td>" +
            "<td>" + val.ProjectRole.ProjectRoles + "</td>" +
            "<td>" + val.Status + "</td>" +
            "<td>" + val.BotStatus + "</td>" +
            "</tr>";
    }).join(""));
    $('#completed-reactivation tbody').append(_.map(completedReactivation, function (val, index) {
        return "'<tr>" +
            "<td>" + val.ID + "</td>" +
            "<td>" + val.Author.Title + "</td>" +
            "<td>" + val.Created + "</td>" +
            "<td>" + val.Status + "</td>" +
            "<td>" + val.BotStatus + "</td>" +
            "</tr>";
    }).join(""));
    $('#completed-unlock-account tbody').append(_.map(completedUnlockAccount, function (val, index) {
        return "'<tr>" +
            "<td>" + val.ID + "</td>" +
            "<td>" + val.Author.Title + "</td>" +
            "<td>" + val.Created + "</td>" +
            "<td>" + val.Status + "</td>" +
            "<td>" + val.BotStatus + "</td>" +
            "</tr>";
    }).join(""));
}

QCUpdateForm.fn.bindFormDetails = function (itemId, restBody) {
    //bind form values
    InjectScript.updateListItem(QCUpdateForm.vars.webUrl, "Quality Center Requests", itemId, restBody, "SP.Data.QCRequestsListItem", true).done(function (data) {
        console.log(data)
    });
}