import { SPUser } from '@microsoft/sp-page-context';
import { HttpClient } from '@microsoft/sp-http';

export interface IChatWindowProps {
    user: SPUser;
    httpClient: HttpClient;
    logSource: string;
    knowledgeBaseKey: string;
    endpointKey: string;
    hostUrl: string;
    filters:any[];
    filterOperator:string;
}

export interface IChatWindowState {
    collapsed?: boolean;
    conversation?: IConversation[];
    question?: string;
    isLoading?: boolean;
}

export interface IConversation {
    source: ISource;
    content: string;
    score: number;
    qnaID: number;
}

export enum ISource {
    BOT,
    PERSON,
    PROMPT
}