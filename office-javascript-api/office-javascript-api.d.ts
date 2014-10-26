// Type definitions for Office Javascript APIs v1.1
// Project: http://dev.office.com/
// Definitions by: Alex Park <https://github.com/MrPiao/>, Adam Sheldon <https://github.com/flacnut>
// Definitions: https://github.com/borisyankov/DefinitelyTyped

declare module OfficeJavascriptApis
{
    interface IOffice {
        context: IContext;
    }

    interface IContext {
        mailbox: IMailbox;
    }

    interface IMailbox {
        item: IItem;
    }

    interface IItem {
        dateTimeCreated: Date;
        dateTimeModified: Date;
    }
}

declare var Office: OfficeJavascriptApis.IOffice;