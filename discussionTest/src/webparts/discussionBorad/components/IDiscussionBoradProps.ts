import { Web } from "sp-pnp-js";

export interface IDiscussionBoradProps {
  description: string;
  discussion: IDiscussion;
  webUrl: string;
}
export interface IDiscussion {
  DiscussionId: number;
  DiscussionTitle: string;
  DiscussionBody: string;
  DiscussionAuthor: number;
  DiscussionLike: number;
  DiscussionLikeStringId: string[];
  DiscussionFolder: string;
  MessagesCount: number;
  Messages: IMessage[];
}
export interface IMessage {
  MessageAuthor: number;
  MessageBody: string;
  MessageID: number;
  MessageLikedByStringId: string[];
  MessageLikesCount: number;
  MessageParentID: number;
}
