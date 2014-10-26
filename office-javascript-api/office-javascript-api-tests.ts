/// <reference path="office-javascript-api.d.ts" />

function testItemProperties() {
    var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
    var dateTimeModified = Office.context.mailbox.item.dateTimeModified;
}