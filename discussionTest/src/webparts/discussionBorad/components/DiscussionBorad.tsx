import * as React from 'react';
import styles from './DiscussionBorad.module.scss';
import { IDiscussionBoradProps } from './IDiscussionBoradProps';
import { escape } from '@microsoft/sp-lodash-subset';
import DiscussionService, { IDiscussionService } from './DiscussionService';

export default class DiscussionBorad extends React.Component<IDiscussionBoradProps, {}> {

  private servcice: IDiscussionService;

  constructor(props: IDiscussionBoradProps) {
    super(props);
    this.servcice = new DiscussionService(this.props.webUrl);
  }
  protected createMessage() {
    let messageBlock = null;
    messageBlock = this.props.discussion.Messages.map((message, index) => {
      var html = { __html: message.MessageBody };
      return <div key={"message" + index} className={styles.row}>
        <span className={styles.title}>AuthorID: {message.MessageAuthor}</span>
        <p dangerouslySetInnerHTML={html} className={styles.subTitle}></p>
        <p className={styles.description}>{message.MessageLikesCount} Likes</p>
      </div>
    })
    return messageBlock;
  }
  protected createDiscussion() {
    var html = { __html: this.props.discussion.DiscussionBody };
    return <div className={styles.row}>
      <span className={styles.title}>AuthorID: {this.props.discussion.DiscussionAuthor}</span>
      <p className={styles.subTitle}>{this.props.discussion.MessagesCount} replies.</p>
      <p dangerouslySetInnerHTML={html} className={styles.description}></p>
      <p className={styles.description}>{this.props.discussion.DiscussionLike} Likes</p>
    </div>
  }


  public render(): React.ReactElement<IDiscussionBoradProps> {
    let messageBlock = this.createMessage();
    let discussionBlock = this.createDiscussion();
    return (
      <div className={styles.discussionBorad}>
        <div className={styles.container}>
          <div className={styles.row}>
            <div className={styles.column}>
              {/* <span className={ styles.title }>Welcome to SharePoint!</span>
              <p className={ styles.subTitle }>Customize SharePoint experiences using Web Parts.</p>
              <p className={ styles.description }>{this.props.discussion.Messages[0].MessageParentID}   {this.props.discussion.DiscussionLike}{escape(this.props.description)}</p>
              <a href="https://aka.ms/spfx" className={ styles.button }>
                <span className={ styles.label }>Learn more</span>
              </a> */}
            </div>
            {discussionBlock}
            {messageBlock}
          </div>
          <div style={{ width: "50px", height: "50px", backgroundColor: "green" }} onClick={() => this.servcice.addMessage(this.props.discussion.DiscussionFolder, "aaa")}>Commit</div>
        </div>
      </div>
    );
  }
}
