import * as ReactDom from 'react-dom';
import * as React from 'react';
import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer, PlaceholderContent, PlaceholderName
} from '@microsoft/sp-application-base';
import { Dialog } from '@microsoft/sp-dialog';

import * as strings from 'ChatBotApplicationCustomizerStrings';
import ChatWindow from '../../components/chatWindow';

const LOG_SOURCE: string = 'SPFx.QnABot';


export interface IChatBotApplicationCustomizerProperties {
  KbKey: string;
  EndPoint: string;
  HostUrl: string;
  Filters: any[];
  FiltersOperator: string;
}

/** A Custom Action which can be run during execution of a Client Side Application */
export default class ChatBotApplicationCustomizer
  extends BaseApplicationCustomizer<IChatBotApplicationCustomizerProperties> {

  private _bottomPlaceholder: PlaceholderContent | undefined;

  @override
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);

    this.context.placeholderProvider.changedEvent.add(this, this._renderPlaceHolder);

    return Promise.resolve();
  }

  private _renderPlaceHolder(): void {

    Log.info(LOG_SOURCE, this.context.placeholderProvider.placeholderNames.map(name => PlaceholderContent[name]).join(", "));

    if (!this._bottomPlaceholder) {
      this._bottomPlaceholder = this.context.placeholderProvider.tryCreateContent(PlaceholderName.Bottom);
    }

    if (this._bottomPlaceholder) {
      const element: React.ReactElement<any> = React.createElement(ChatWindow, {
        knowledgeBaseKey: escape(this.properties.KbKey),
        endpointKey: escape(this.properties.EndPoint),
        user: this.context.pageContext.user,
        httpClient: this.context.httpClient,
        logSource: LOG_SOURCE,
        hostUrl: escape(this.properties.HostUrl),
        filters: this.properties.Filters,
        filterOperator: this.properties.FiltersOperator
      });

      ReactDom.render(element, this._bottomPlaceholder.domElement);
    }
  }

}
