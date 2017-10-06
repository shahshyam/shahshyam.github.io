var item;
Office.initialize = function () {
    item = Office.context.mailbox.item;
}

// Helper function to add a status message to the info bar.
function statusUpdate(icon, text) {
  Office.context.mailbox.item.notificationMessages.replaceAsync("status", {
    type: "informationalMessage",
    icon: icon,
    message: text,
    persistent: false
  });
}

function defaultStatus(event) {
  statusUpdate("icon16" , "Hello World!");
}

function doSomethingAndShowDialog(event) {
            clickEvent = event;
            //writeToDoc("Ribbon button clicked.");
            openDialogAsIframe();
            event.completed();
        }

function addTextToBody(text, icon, event) {
    Office.context.mailbox.item.body.setSelectedDataAsync(text, { coercionType: Office.CoercionType.Text },
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                statusUpdate(icon, "\"" + text + "\" inserted successfully.");
            }
            else {
                Office.context.mailbox.item.notificationMessages.addAsync("addTextError", {
                    type: "errorMessage",
                    message: "Failed to insert \"" + text + "\": " + asyncResult.error.message
                });
            }
            event.completed();
        });
}



function UpdateSubject(event)
{
    var body;
    var Recipients;
    Office.context.mailbox.item.cc.getAsync(callback);   
    item.body.getAsync(
        "html",
        { asyncContext: "This is passed to the callback" },
        function callback(result) {
            body = result.value;

            //Setting update body
            Office.context.mailbox.item.body.setAsync(
                "<b>" + body + "</b>",
                { coercionType: "html", asyncContext: "This is passed to the callback" },
                function callback(result) {
                    if (result.status == Office.AsyncResultStatus.Succeeded)
                    {
                       
                    }
                });


        });  
    event.completed();
}

function PostDataToAPI(event)
{
    var item = Office.context.mailbox.item;
    var uri = 'http://localhost:1300/api/cities';
    var imapSetting = {
        Subject: item.subject,
        Body: 'helllos'       
    }
    $.ajax({
        url: uri,
        type: 'POST',
        data: JSON.stringify(imapSetting),
        contentType: "application/json",
        success: function (d) {
            //  alert("Saved Successfully");
            //document.getElementById("postform").reset();
            if (d = true) {
                writemessage(d);
            }
        },
        error: function () {
            writemessage("errrpr");
        }
    });

    event.completed();
}

function writemessage(messages)
{
   
}

function DisaplyBody(asyncResult) {

    var currentbody = asyncResult.value;
    addTextToBody(currentbody + " update Text", "red-icon-16", event);
}
function addMsg1ToBody(event) {
    addTextToBody("Hello World!", "red-icon-16", event);
}