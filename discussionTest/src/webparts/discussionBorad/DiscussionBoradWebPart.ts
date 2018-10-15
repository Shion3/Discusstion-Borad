import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'DiscussionBoradWebPartStrings';
import DiscussionBorad from './components/DiscussionBorad';
import { IDiscussionBoradProps } from './components/IDiscussionBoradProps';
import { Web } from '../../../node_modules/sp-pnp-js/lib/pnp';


export interface IDiscussionBoradWebPartProps {
  description: string;
}

export default class DiscussionBoradWebPart extends BaseClientSideWebPart<IDiscussionBoradWebPartProps> {
  private listTitle = "testDiscussion";
  public render(): void {

    this.loadSpecialData();
    // let datas = [];
    // this.loadAllDiscussions(datas);
  }
  loadSpecialData(): any {
    let web = new Web(this.context.pageContext.web.absoluteUrl);
    const expandFileds = ["FieldValuesAsText", "Folder"];
    let filterStr = `ContentType eq 'Discussion'`;
    web.lists.getByTitle(this.listTitle).items.getById(1).expand(...expandFileds).get().then((Discussion) => {
      this.RetriveMessages(web, Discussion);
    })
  }
  loadAllDiscussions(datas: any[]): any {
    let web = new Web(this.context.pageContext.web.absoluteUrl);
    const expandFileds = ["FieldValuesAsText", "Folder"];
    let filterStr = `ContentType eq 'Discussion'`;
    web.lists.getByTitle(this.listTitle).items.filter(filterStr).expand(...expandFileds).get().then((Discussions) => {
      Discussions.forEach(discussion => {
        datas.push({ DiscussionTitle: discussion.Title, DiscussionBody: discussion.Body, DiscussionAuthor: discussion.AuthorId, DiscussionLike: discussion.LikesCount, DiscussionLikeStringId: discussion.LikedByStringId, MessagesCount: discussion.Folder.ItemCount })
      })

    })
  }
  RetriveMessages(web: Web, Discussion: any): any {
    const selectField = ["Body", "AuthorId", "FieldValuesAsText/FileRef", "LikedByStringId", "LikesCount", "ParentItemID", "ID"];
    const expandFileds = ["FieldValuesAsText"];
    let filterStr = `(startswith(FileRef,'${Discussion.Folder.ServerRelativeUrl}')) and (ContentType eq 'Message')`;

    web.lists.getByTitle(this.listTitle).items.select(...selectField).filter(filterStr).expand(...expandFileds).get().then((Messages) => {
      this.CombinDatas(Messages, Discussion);
    })

  }
  CombinDatas(Messages: any, Discussion: any): any {
    let messages = [];
    Messages.forEach(message => {
      messages.push({ MessageBody: message.Body, MessageAuthor: message.AuthorId, MessageLikesCount: message.LikesCount, MessageParentID: message.ParentItemID, MessageID: message.ID, MessageLikedByStringId: message.LikedByStringId });
    })
    let IDiscussion = { DiscussionId: Discussion.ID, DiscussionTitle: Discussion.Title, DiscussionBody: Discussion.Body, DiscussionAuthor: Discussion.AuthorId, DiscussionLike: Discussion.LikesCount, DiscussionLikeStringId: Discussion.LikedByStringId, DiscussionFolder: Discussion.FieldValuesAsText.FileRef, MessagesCount: Discussion.Folder.ItemCount, Messages: messages };
    const element: React.ReactElement<IDiscussionBoradProps> = React.createElement(
      DiscussionBorad,
      {
        description: this.properties.description,
        discussion: IDiscussion,
        webUrl: this.context.pageContext.web.absoluteUrl,
      }
    );
    ReactDom.render(element, this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
