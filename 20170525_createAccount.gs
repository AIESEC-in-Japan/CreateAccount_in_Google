var MCIS = 'mc_is_log@aiesec.jp';
var SHEETNAME = "フォームの回答 2";

function main() {
    var sheet       = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETNAME);

//    var lock        = LockService.getScriptLock();
//    lock.waitLock(3000); // Lock this process

    var lastRow     = sheet.getLastRow();
    var lastColumn  = sheet.getLastColumn();
    var rowRange    = sheet.getRange(lastRow, 1, 1, lastColumn);
    var rowData     = rowRange.getValues();

    try {
        insertUser(rowData[0]);
    } catch (e) {
        var subject = '[GAS: Error] createAccount: ' + e;
        var body = JSON.stringify(rowData, null, 4);
        MailApp.sendEmail(MCIS, subject, body);
    }
//    lock.releaseLock();
}

function insertUser(row) {

    var givenName   = row[3];
    var familyName  = row[4];
    var middleName  = row[5];
    var extEmail    = row[6];
    var chkEmail    = row[7];
    var birthYear   = row[8];
    var lc          = row[9];

    var primaryEmail = givenName +'.'+ familyName +'@aiesec.jp';
    givenName       = givenName.capitalize();
    familyName      = familyName.capitalize();
    middleName      = middleName.capitalize();
    var fullName    = (middleName) ? givenName +" "+ middleName +" "+ familyName
                                   : givenName +" "+ familyName;
    var randomStr = Math.random().toString(36).slice(-8);
    var userObj = {
        "orgUnitPath": "/LC/" + lc,
        "isMailboxSetup": true,
        "primaryEmail": primaryEmail,
        "kind": "admin#directory#user",
        "isAdmin": false,
        "suspended": false,
        "isDelegatedAdmin": false,
        "name": {
            "givenName": givenName,
            "familyName": familyName,
            "fullName": fullName
        },
        "ipWhitelisted": false,
        "emails": [{
            "address": primaryEmail,
            "primary": true
        },{
            "address": extEmail,
            "primary": false
        }],
            "changePasswordAtNextLogin": true,
            "agreedToTerms": true,
            "includeInGlobalAddressList": true,
            "organizations": [{
                "description": String(birthYear),
                "title": fullName,
                "department": lc,
                "primary": true
            }],
            "password": randomStr
    };
    var lcml    = lc.toLowerCase() + '_all@aiesec.jp';
    var lcAdmin = lc.toLowerCase() + '.admin@aiesec.jp';

    AdminDirectory.Users.insert(userObj);
    addGroup(lcml, userObj);
    mailInfo(fullName, primaryEmail, randomStr, extEmail, chkEmail, lc, lcml, lcAdmin);
}

function addGroup(lcml, userObj) {
    var memberObj = {
        "email": userObj.primaryEmail,
        "role": "MEMBER"
    };
    AdminDirectory.Members.insert(memberObj, lcml);
}

function mailInfo(name, account, pass, extEmail, chkEmail, lc, lcml, lcAdmin) {
    var body1 = (
        '<h3>' +
        name +
        ' さんのアカウントが作成されました．</h3>' +
        '<p>アカウント: ' +
        account +
        '</p><p>パスワード: ' +
        pass +
        '</p><a href="https://accounts.google.com/login">ログイン</a>'
    );

    MailApp.sendEmail({
        to:         extEmail,
        cc:         lcAdmin + ', ' + chkEmail,
        bcc:        MCIS,
        subject:    '【重要】新規 aiesec.jp アカウント作成の完了',
        htmlBody:   body1
    });

    var body2 = ('<p><b>'+ name + '</b>さんがグループ ' + lcml + ' に参加しました!</p>');
    MailApp.sendEmail({
        to:         lcml,
        subject:    '【自動通知】新規グループメンバー',
        htmlBody:   body2
    });
}

String.prototype.capitalize = function() {
    return this.charAt(0).toUpperCase() + this.toLowerCase().slice(1);
};
