
Office.initialize = function () {
}

function trackUrlAndLink(event) {
    var item = Office.context.mailbox.item;   

    var body;  

    item.body.getAsync(Office.CoercionType.Html, { asyncContext: "This is passed to the callback" }, function callback(asyncResult) {
        if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
            body = asyncResult.value;
            Office.context.mailbox.item.to.getAsync(function callback(result)
            {
                if (result.status == Office.AsyncResultStatus.Succeeded) {
                    var arrayOfToRecipients = result.value;
                    replaceBodyText(body, arrayOfToRecipients, event);
                }
                else {
                    writeLog('error in recipients')
                }
            });
        }
        else {
            writeLog('error in body')
        }
    });  
  
}

function replaceBodyText(body,listRecipient,event)
{
    var emailList = "?id_op=";
    var isRecipients = false;
    if (listRecipient != null)
    {
        var count = listRecipient.length;
        for (var i = 0; i < count; i++)
        {
            var item = listRecipient[i];
            if (count - 1 == i)
            {
                emailList += item.emailAddress;
            } else {
                emailList += item.emailAddress+"&";
            }
           
            isRecipients = true;
        }

    }
    if (isRecipients && body != null)
    {
      
        var href_regex = /<a.*?href=.*?>/g;
       
        var matches = body.match(href_regex);
        //Iterate loop to find href and relace it
        for (var i = 0; i < matches.length; i++)
        {
            //var singleMatch = matches[i];
            var url = matches[i];
            if (url != null)
            {
                var updateUrl = url.replace("\">", emailList + "\">");
                body = body.replace(url, updateUrl);
            }
        }
        //update body
        Office.context.mailbox.item.body.setAsync(
            body,
            { coercionType: "html", asyncContext: "This is passed to the callback" },
            function callback(resultUp) {
                if (resultUp.status == Office.AsyncResultStatus.Succeeded) {
                    statusUpdate("icon16", "Tracking infor is added");
                    event.completed();
                }
            });
    }
    
}


function writeLog(message)
{

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