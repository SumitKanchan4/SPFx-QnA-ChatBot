import * as React from 'react';
import { IChatWindowProps, IChatWindowState, ISource, IConversation } from './IChatWindowProps';
import styles from './chatWindow.module.scss';
import { Icon, TextField, Link } from 'office-ui-fabric-react';
import { HttpClient, HttpClientResponse, IHttpClientOptions } from '@microsoft/sp-http';
import { Log } from '@microsoft/sp-core-library';
import { animateScroll } from "react-scroll";
import { ChatBotImage } from '../extensions/chatBot/chatBotImage';
import Markdown from 'markdown-to-jsx';

export default class ChatWindow extends React.Component<IChatWindowProps, IChatWindowState>{

    constructor(props: IChatWindowProps, state: IChatWindowState) {
        super(props);

        this.state = {
            collapsed: true,
            conversation: [{ source: ISource.BOT, content: `Hello ${this.props.user.displayName.split(" ")[0]}`, score: 100, qnaID: -1 }],
            isLoading: false,
            question: undefined
        };

        // Listen to the key for "ENTER" to get the answer on the press
        window.addEventListener("keydown", (e) => {
            if (e.which == 13) {
                this.getAnswer();
                return false;
            }
        });
    }

    /**
     * Returns the formatted date string to be shown against each chat message
     */
    private get getDateString(): string {
        let date: Date = new Date();
        return `${date.getHours()}:${date.getMinutes()}`;
    }

    public render(): React.ReactElement<IChatWindowProps> {
        return (

            <div className={styles.chatWindow}>
                {
                    this.state.collapsed ?
                        <div>
                            {!this.isMobile ?
                                <section className={`${styles.avenueMessenger} ${styles.bottom}`} onClick={() => this.minimizeChatPop(false)}>
                                    <div>
                                        <Icon className={`ms-fontSize-18`} iconName={`ChatBot`} ></Icon>
                                        Ask Indie
                                    </div>
                                </section>
                                :
                                <div>
                                </div>
                            }
                        </div>
                        :
                        <div>
                            <section className={styles.avenueMessenger}>
                                <div className={styles.menu}>
                                    <Icon iconName={`ChromeMinimize`} className={styles.button} onClick={() => this.minimizeChatPop(true)} ></Icon>
                                </div>
                                <div className={styles.agentFace}>
                                    <div className={styles.half}>
                                        <img className={styles.circle} src={ChatBotImage.base64} alt="SK BOT" /></div>
                                </div>
                                <div className={styles.chat}>
                                    <div className={styles.chatTitle}>
                                        <h1>Ask Indie</h1>
                                        <h2>Your SharePoint assistant</h2>
                                    </div>
                                    <div className={styles.messages} id={`chatMessages`}>
                                        <div className={styles["messages-content"]}>
                                            {
                                                this.state.conversation.map(item => {
                                                    return item.source == ISource.PERSON ? this.renderPersonalChat(item) : this.renderBotChat(item);
                                                })
                                            }
                                        </div>
                                    </div>
                                    <div className={styles.messageBox}>
                                        <TextField rows={2} multiline={this.state.question && this.state.question.length > 15 ? true : false} borderless={true} value={this.state.question || ''} className={styles.messageInput} placeholder="Type message..." onChange={(e, txt) => this.onTextChanged(e, txt)}></TextField>
                                        <button type="submit" className={styles.messageSubmit} onClick={() => this.getAnswer()} disabled={this.state.question == undefined}>Send</button>
                                    </div>
                                </div>
                            </section>
                        </div>
                }
            </div>
        );
    }


    /**
     * Renders the bot chat section
     */
    private renderBotChat(item: IConversation): JSX.Element {
        return (
            <div className={`${styles.message} ${styles.new}`}>
                <figure className={styles.avatar}>
                    <img src={ChatBotImage.base64} title="INDIE says..." />
                </figure>
                {
                    item.source == ISource.BOT ?
                        <Markdown>{item.content}</Markdown>
                        :
                        <Link onClick={() => this.onPromptClick(item.content)}>{item.content}</Link>
                }


                <div className={styles.timestamp}>{this.getDateString}</div>
            </div>
        );
    }

    /**
     * Click event for the propmpt question
     * @param question 
     */
    private onPromptClick(question: string): void {
        this.setState({ question: question }, () => {
            this.getAnswer();
        });
    }

    /**
     * Capture the questiona as user types in
     * @param e 
     * @param txt 
     */
    private onTextChanged(e: any, txt: string): void {
        this.setState({ question: txt });
    }

    /**
     * Renders the personal chat section
     */
    private renderPersonalChat(item: IConversation): JSX.Element {
        return (
            <div className={`${styles.message} ${styles.messagePersonal}`}>
                {item.content}
                <div className={styles.timestamp}>{this.getDateString}</div>
            </div>
        );
    }

    /**
     * Minimizes or maximizes the chat pop
     * @param isCollapsed bool to make the chat pop collapse and open
     */
    private minimizeChatPop(isCollapsed: boolean): void {
        this.setState({ collapsed: isCollapsed });
    }


    /**
     * Calls for the QnA API to fetch the answer for thas asked question
     */
    private async getAnswer(): Promise<void> {

        if (this.state.question) {
            //Get the question form the state
            let questionAsked: string = this.state.question;
            // Show the loading while the answer is reqreived from the qna
            this.setState({ isLoading: true, question: undefined });
            // Adds the conversation as the user
            this.addConversation([{ content: questionAsked, score: 100, source: ISource.PERSON, qnaID: -1 }]);

            // Create url
            let qnaUrl: string = decodeURIComponent(`${this.props.hostUrl}/knowledgebases/${this.props.knowledgeBaseKey}/generateAnswer`);

            // format the query and headers
            let httpClientOptions: IHttpClientOptions = this.createHttpClientOptions(questionAsked);

            try {
                let queryResponse: HttpClientResponse = await this.props.httpClient.post(qnaUrl, HttpClient.configurations.v1, httpClientOptions);
                let json: any = await queryResponse.json();

                if (json) {
                    let answerObject: any = json.answers[0];
                    let answer: IConversation[] = [{
                        content: answerObject.answer,
                        qnaID: answerObject.id,
                        score: answerObject.score,
                        source: ISource.BOT
                    }];

                    this.addConversation(answer);
                    this.checkForPrompts(answerObject);
                    Log.verbose(this.props.logSource, JSON.stringify(answer));
                }
            }
            catch (error) {
                Log.error(this.props.logSource, error);
                this.addConversation([{ content: "Some error occured while retreiving the answer :(", score: 0, source: ISource.BOT, qnaID: -1 }]);
            }
        }
    }

    /**
     * Checks if there are any prompts available for the question
     * @param answerObject 
     */
    private checkForPrompts(answerObject: any): void {

        if (answerObject["context"] && answerObject.context["prompts"] && answerObject.context.prompts.length > 0) {
            let prompts: IConversation[] = [];
            answerObject.context.prompts.forEach((prompt) => {
                try {
                    prompts.push({
                        content: prompt.displayText,
                        score: 0,
                        source: ISource.PROMPT,
                        qnaID: prompt.qnaId
                    });
                }
                catch (ex) {
                    Log.error(this.props.logSource, ex);
                }
            });

            prompts.length > 0 ? this.addConversation(prompts) : undefined;
        }
    }

    /**
     * Creates the HttpClientOptions object which sets the header and bory for the query
     * @param questionAsked 
     */
    private createHttpClientOptions(questionAsked: string): IHttpClientOptions {
        let requestHeaders: Headers = new Headers();
        requestHeaders.append('Content-type', 'application/json');
        requestHeaders.append('Authorization', `EndpointKey ${this.props.endpointKey}`);

        let body: any = { question: questionAsked };

        // add the strict filters if filter is required
        if (this.props.filters.length > 0) {
            body["strictFilters"] = this.props.filters;

            // add the filter operater OR , AND if the filters are more than 1
            if (this.props.filters.length > 1) {
                body["strictFiltersCompoundOperationType"] = this.props.filterOperator;
            }
        }

        let httpClientOptions: IHttpClientOptions = {
            body: JSON.stringify(body),
            headers: requestHeaders
        };

        return httpClientOptions;
    }

    /**
     * Adds teh conversation answer/question to the state for rendering
     * @param content 
     * @param score 
     * @param source 
     */
    private addConversation(conversation: IConversation[]): void {
        this.setState({ conversation: [...this.state.conversation, ...conversation] });
        this.scrollToBottom();
    }

    /**
     * Scrolls the conversation area to the bottom to see the latest reply/question
     */
    private scrollToBottom() {
        animateScroll.scrollToBottom({
            containerId: "chatMessages"
        });
    }

    /**
     * Checks if the device is mobile
     */
    private get isMobile(): boolean {

        return (navigator.userAgent.indexOf("Android") > -1
            || navigator.userAgent.indexOf("webOS") > -1
            || navigator.userAgent.indexOf("iPhone") > -1
            || navigator.userAgent.indexOf("iPad") > -1
            || navigator.userAgent.indexOf("iPod") > -1
            || navigator.userAgent.indexOf("BlackBerry") > -1
            || navigator.userAgent.indexOf("Windows Phone") > -1);
    }
}