import { Web, ItemAddResult } from "sp-pnp-js";
import $ from 'jquery';

export interface IDiscussionService {
    addMessage(folderUrl, messageInfo);
}
export default class DiscussionService implements IDiscussionService {
    private web: Web;
    private webUrl: string;
    private listTitle: string = "testDiscussion";
    constructor(webUrl: string) {
        this.web = new Web(webUrl);
        this.webUrl = webUrl;
    }
    public addMessage(folderUrl: string, messageInfo: string) {
        this.web.lists.getByTitle(this.listTitle).items.add({
            ParentItemID: 1,
            ParentItemEditorId: 9,
            ContentTypeId: "0x0107",
            Body: messageInfo,
        }).then((result: ItemAddResult) => {
            result.item.update({
                FileRef: folderUrl + "/" + result.data.ID + "_.000"
            })
        })


        // let url = this.webUrl + "/_api/web/lists/getByTitle('" + this.listTitle + "')/AddValidateUpdateItemUsingPath";
        // $.ajax({

        // })

        // let postData = {
        //     "listItemCreateInfo": {
        //         "FolderPath": {
        //             "DecodedUrl":
        //                 folderUrl
        //         },
        //         "UnderlyingObjectType": 0
        //     },
        //     "formValues": [
        //         {
        //             "FieldName": "Body",
        //             "FieldValue": "Reply"
        //         },
        //         {
        //             "FieldName": "ParentItemID",
        //             "FieldValue": 1
        //         },
        //         {
        //             "FieldName": "ParentItemEditorId",
        //             "FieldValue": 9
        //         },
        //         {
        //             "FieldName": "contentTypeId",
        //             "FieldValue": "0x0107"
        //         }
        //     ],
        //     "bNewDocumentUpdate": false
        // };

    }
}