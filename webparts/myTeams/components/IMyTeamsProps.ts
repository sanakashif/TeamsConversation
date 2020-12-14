import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IMyTeamsProps {

context: WebPartContext;
teamDisplayName: string;
channelDisplayName: string;
readOnly:boolean;
numberOfPosts:number;
description: string;
}
