//This function sets up the trigger on Google Form quiz submission

function setUpTrigger() {
    ScriptApp.newTrigger('outcomeWikiGroup')
        .forSpreadsheet('[GOOGLE SPREADSHEET ID]') // Add Google Spreadsheet ID here
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
    } finally { // Add Group ID Numbers here, based on tab names on the responses spreadsheet
        if (parseInt(quiz_num) === 101) {
            addGroup = [53]
        } else if (parseInt(quiz_num) === 102) {
            addGroup = [54]
        } else if (parseInt(quiz_num) === 103) {
            addGroup = [55]
        } else if (parseInt(quiz_num) === 104) {
            addGroup = [56]
        } else if (parseInt(quiz_num) === 105) {
            addGroup = [57]
        } else if (parseInt(quiz_num) === 106) {
            addGroup = [58]
        } else if (parseInt(quiz_num) === 107) {
            addGroup = [59]
        } else if (parseInt(quiz_num) === 108) {
            addGroup = [60]
        } else if (parseInt(quiz_num) === 109) {
            addGroup = [61]
        } else if (parseInt(quiz_num) === 201) {
            addGroup = [62]
        } else if (parseInt(quiz_num) === 202) {
            addGroup = [63]
        } else if (parseInt(quiz_num) === 203) {
            addGroup = [64]
        } else if (parseInt(quiz_num) === 204) {
            addGroup = [65]
        } else if (parseInt(quiz_num) === 205) {
            addGroup = [66]
        }
        // Passed Quiz block below
        if (finalScore >= 75) { // Here, pass is definined as >= 75. You can change as needed 
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
            var response = UrlFetchApp.fetch('[WEBHOOK URL]', options); //Webhook is pinged, sending the user group additions/changes as needed  

            var code = parseInt(response.getResponseCode());

            // The emails below can also be done with html templates if desired, instead of text in the app script. 

            if (code === 200) { // If they passed the quiz, and no error, then htmlBody is sent to the emails. Edit emails and message here as necessary. 
                var opt = {
                    'bcc': 'keshavan@staclabs.io' + ',' + 'josh@kydemocrats.org' + ',' + 'brandon@kydemocrats.org',
                    'name': 'votebuilder@kydemocrats.org',
                    'replyTo': 'votebuilder@kydemocrats.org',
                    'htmlBody': `With a score of ${finalScore} you passed the quiz! Please log out of and log back into your Progress Wiki account, and you should see the next part of the VAN coursework!<br> If you have any questions, please contact us at votebuilder@kydemocrats.org. Thank you for your effort!`
                }

                MailApp.sendEmail(email, `VAN ${quiz_num} Quiz Results`, opt.htmlBody, opt);

            } else if (code === 404) { // If they passed the quiz, and a 404 error, then htmlBody is sent to the emails, but user groups are not changed. 
                var opt = {
                    'bcc': 'keshavan@staclabs.io' + ',' + 'josh@kydemocrats.org' + ',' + 'brandon@kydemocrats.org',
                    'name': 'votebuilder@kydemocrats.org',
                    'replyTo': 'votebuilder@kydemocrats.org',
                    'htmlBody': `With a score of ${finalScore} you passed the quiz. However, the email you used for the quiz did not match the email that you signed up with for your VAN coursework. Please resubmit the quiz using the correct email. If you have any questions, please contact us at votebuilder@kydemocrats.org.<br> Thank you for your effort!`
                }

                MailApp.sendEmail(email, `VAN ${quiz_num} Quiz Results`, opt.htmlBody, opt);

            } else {
                var opt = { // If some other error, then you need to troubleshoot!
                    'cc': 'keshavan@staclabs.io',
                    'name': 'votebuilder@kydemocrats.org',
                    'replyTo': 'votebuilder@kydemocrats.org',
                    'htmlBody': `With a score of ${finalScore}, ${email} passed the quiz. Something went really wrong though, so please troubleshoot.`
                }
                MailApp.sendEmail('josh@kydemocrats.org' + ',' + 'brandon@kydemocrats.org', `VAN ${quiz_num} Quiz Error`, opt.htmlBody, opt);

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
            var response = UrlFetchApp.fetch('[WEBHOOK URL]', options);

            var code = parseInt(response.getResponseCode());

            if (code === 200) {
                var opt = {
                    'bcc': 'keshavan@staclabs.io' + ',' + 'josh@kydemocrats.org' + ',' + 'brandon@kydemocrats.org',
                    'name': 'votebuilder@kydemocrats.org',
                    'replyTo': 'votebuilder@kydemocrats.org',
                    'htmlBody': `With a score of ${finalScore} you did not pass the quiz. You need a minimum of 75 points to     pass. If you are waiting for free response answers to be graded,<br> then your score may increase; please wait for those grades to be submitted. Otherwise, please review the training, and take the quiz again. If you have <br> any questions, please contact us at votebuilder@kydemocrats.org. Thank you for your effort!`
                }

                MailApp.sendEmail(email, `VAN ${quiz_num} Quiz Results`, opt.htmlBody, opt);

            } else if (code === 404) {
                var opt = {
                    //  'cc' : 'josh@kydemocrats.org',
                    'bcc': 'keshavan@staclabs.io' + ',' + 'josh@kydemocrats.org' + ',' + 'brandon@kydemocrats.org',
                    'name': 'votebuilder@kydemocrats.org',
                    'replyTo': 'votebuilder@kydemocrats.org',
                    'htmlBody': `With a score of ${finalScore} you did not pass the quiz. You need a minimum of 75 points to pass. Please note that the email you used did not match the email that you signed up with for your VAN coursework.<br> Please resubmit the quiz using the correct email. If you have any questions, please contact us at votebuilder@kydemocrats.org.<br> Thank you for your effort!`
                }

                MailApp.sendEmail(email, `VAN ${quiz_num} Quiz Results`, opt.htmlBody, opt);

            } else {
                var opt = {
                    'cc': 'keshavan@staclabs.io',
                    'name': 'votebuilder@kydemocrats.org',
                    'replyTo': 'votebuilder@kydemocrats.org',
                    'htmlBody': `With a score of ${finalScore}, ${email} did not pass the quiz. Something went really wrong though, so please troubleshoot.`
                }
                MailApp.sendEmail('josh@kydemocrats.org' + ',' + 'brandon@kydemocrats.org', `VAN ${quiz_num} Quiz Error`, opt.htmlBody, opt);

            }

        }


    }



}
