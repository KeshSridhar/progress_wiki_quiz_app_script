//This function sets up the trigger on Google Form quiz submission

function setUpTrigger() {
    ScriptApp.newTrigger('outcomeWikiGroup')
        .forSpreadsheet('[GOOGLE SPREADSHEET ID]') // EDIT HERE: Add Google Spreadsheet ID here
        .onFormSubmit()
        .create();
}

//This function grabs the newest submission, and if they pass, moves them to the next level. 
// An email is sent regardless of outcome.
function outcomeWikiGroup(e) {
    var quiz_num = e.range.getSheet().getName(); // Name of tab

    try {
        var email = e.range.getCell(1, 2).getValue();
        var finalScore = parseInt(e.range.getCell(1, 3).getValue());
        var addGroup = []
    } catch (err) {
        var email = e.range.offset(0, -1).getValue();
        var finalScore = parseInt(e.range.getValue());
        var addGroup = []
    } finally { // EDIT HERE: Add Wiki UserGroup ID Numbers here, based on Tab Names on the responses spreadsheet. 
        if (parseInt(quiz_num) === [SPREADSHEET TAB NAME]) {
            addGroup = [USER GROUP ID]
        } else if (parseInt(quiz_num) === [SPREADSHEET TAB NAME]) {
            addGroup = [USER GROUP ID]
        } else if (parseInt(quiz_num) === [SPREADSHEET TAB NAME]) {
            addGroup = [USER GROUP ID]
        } else if (parseInt(quiz_num) === [SPREADSHEET TAB NAME]) {
            addGroup = [USER GROUP ID]
        } else if (parseInt(quiz_num) === [SPREADSHEET TAB NAME]) {
            addGroup = [USER GROUP ID]
        } else if (parseInt(quiz_num) === [SPREADSHEET TAB NAME]) {
            addGroup = [USER GROUP ID]
        } else if (parseInt(quiz_num) === [SPREADSHEET TAB NAME]) {
            addGroup = [USER GROUP ID]
        } else if (parseInt(quiz_num) === [SPREADSHEET TAB NAME]) {
            addGroup = [USER GROUP ID]
        } else if (parseInt(quiz_num) === [SPREADSHEET TAB NAME]) {
            addGroup = [USER GROUP ID]
        } else if (parseInt(quiz_num) === [SPREADSHEET TAB NAME]) {
            addGroup = [USER GROUP ID]
        } else if (parseInt(quiz_num) === [SPREADSHEET TAB NAME]) {
            addGroup = [USER GROUP ID]
        } else if (parseInt(quiz_num) === [SPREADSHEET TAB NAME]) {
            addGroup = [USER GROUP ID]
        } else if (parseInt(quiz_num) === [SPREADSHEET TAB NAME]) {
            addGroup = [USER GROUP ID]
        } else if (parseInt(quiz_num) === [SPREADSHEET TAB NAME]) {
            addGroup = [USER GROUP ID]
        }
        // Passed Quiz block below
        if (finalScore >= 75) { // EDIT HERE: Here, pass is definined as >= 75. You can change as needed 
            var data = {
                'email': email,
                'addGroups': addGroup,
                'removeGroups': [] // You can add even more logic above and remove groups, if needed 
            };
            var options = {
                'method': 'post',
                'contentType': 'application/json',
                'payload': JSON.stringify(data),
                'muteHttpExceptions': true
            };
            var response = UrlFetchApp.fetch('[WEBHOOK URL]', options); //EDIT HERE: Webhook is pinged, sending the user group additions/changes as needed  

            var code = parseInt(response.getResponseCode());

            // The emails below can also be done with html templates if desired, instead of text in the app script. 

            if (code === 200) { //EDIT HERE: If they passed the quiz, and no error, then htmlBody is sent to the emails. Edit emails and message here as necessary. 
                var opt = {
                    'bcc': '[ADMIN 1 EMAIL]' + ',' + '[ADMIN 2 EMAIL]' + ',' + '[ADMIN 3 EMAIL]',
                    'name': '[FROM EMAIL]',
                    'replyTo': '[REPLY-TO EMAIL]',
                    'htmlBody': `With a score of ${finalScore} you passed the quiz! Please log out of and log back into your Progress Wiki account, and you should see the next part of the VAN coursework!<br> If you have any questions, please contact us at votebuilder@kydemocrats.org. Thank you for your effort!`
                }

                MailApp.sendEmail(email, `VAN ${quiz_num} Quiz Results`, opt.htmlBody, opt);

            } else if (code === 404) { // EDIT HERE: If they passed the quiz, and a 404 error, then htmlBody is sent to the emails, but user groups are not changed. 
                var opt = {
                    'bcc': '[ADMIN 1 EMAIL]' + ',' + '[ADMIN 2 EMAIL]' + ',' + '[ADMIN 3 EMAIL]',
                    'name': '[FROM EMAIL]',
                    'replyTo': '[REPLY-TO EMAIL]',
                    'htmlBody': `With a score of ${finalScore} you passed the quiz. However, the email you used for the quiz did not match the email that you signed up with for your VAN coursework. Please resubmit the quiz using the correct email. If you have any questions, please contact us at votebuilder@kydemocrats.org.<br> Thank you for your effort!`
                }

                MailApp.sendEmail(email, `VAN ${quiz_num} Quiz Results`, opt.htmlBody, opt);

            } else {
                var opt = { //EDIT HERE: If some other error, then you need to troubleshoot!
                    'cc': '[ADMIN OPTIONAL EMAIL]',
                    'name': '[FROM EMAIL]',
                    'replyTo': '[REPLY-TO EMAIL]',
                    'htmlBody': `With a score of ${finalScore}, ${email} passed the quiz. Something went really wrong though, so please troubleshoot.`
                }
                MailApp.sendEmail('[ADMIN 1 EMAIL]' + ',' + '[ADMIN 2 EMAIL]', `VAN ${quiz_num} Quiz Error`, opt.htmlBody, opt);

            }

            // Did not pass quiz block below. Same as above block, but no groups added. 
        } else {
            var data = {
                'email': email,
                'addGroups': [],
                'removeGroups': []
            };
            var options = {
                'method': 'post',
                'contentType': 'application/json',
                'payload': JSON.stringify(data),
                'muteHttpExceptions': true
            };
            var response = UrlFetchApp.fetch('[WEBHOOK URL]', options); // EDIT HERE

            var code = parseInt(response.getResponseCode());

            if (code === 200) { // EDIT HERE
                var opt = {
                    'bcc': '[ADMIN 1 EMAIL]' + ',' + '[ADMIN 2 EMAIL]' + ',' + '[ADMIN 3 EMAIL]',
                    'name': '[FROM EMAIL]',
                    'replyTo': '[REPLY-TO EMAIL]',
                    'htmlBody': `With a score of ${finalScore} you did not pass the quiz. You need a minimum of 75 points to     pass. If you are waiting for free response answers to be graded,<br> then your score may increase; please wait for those grades to be submitted. Otherwise, please review the training, and take the quiz again. If you have <br> any questions, please contact us at votebuilder@kydemocrats.org. Thank you for your effort!`
                }

                MailApp.sendEmail(email, `VAN ${quiz_num} Quiz Results`, opt.htmlBody, opt);

            } else if (code === 404) { // EDIT HERE
                var opt = {
                    'bcc': '[ADMIN 1 EMAIL]' + ',' + '[ADMIN 2 EMAIL]' + ',' + '[ADMIN 3 EMAIL]',
                    'name': '[FROM EMAIL]',
                    'replyTo': '[REPLY-TO EMAIL]',
                    'htmlBody': `With a score of ${finalScore} you did not pass the quiz. You need a minimum of 75 points to pass. Please note that the email you used did not match the email that you signed up with for your VAN coursework.<br> Please resubmit the quiz using the correct email. If you have any questions, please contact us at votebuilder@kydemocrats.org.<br> Thank you for your effort!`
                }

                MailApp.sendEmail(email, `VAN ${quiz_num} Quiz Results`, opt.htmlBody, opt);

            } else {
                var opt = { // EDIT HERE
                    'cc': '[ADMIN OPTIONAL EMAIL]',
                    'name': '[FROM EMAIL]',
                    'replyTo': '[REPLY-TO EMAIL]',
                    'htmlBody': `With a score of ${finalScore}, ${email} did not pass the quiz. Something went really wrong though, so please troubleshoot.`
                }
                MailApp.sendEmail('[ADMIN 1 EMAIL]' + ',' + '[ADMIN 2 EMAIL]', `VAN ${quiz_num} Quiz Error`, opt.htmlBody, opt);

            }

        }


    }



}
