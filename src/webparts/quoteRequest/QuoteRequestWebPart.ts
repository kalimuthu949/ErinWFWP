import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart, WebPartContext } from '@microsoft/sp-webpart-base';
import * as strings from 'QuoteRequestWebPartStrings';
import QuoteRequest from './components/QuoteRequest';
import { IQuoteRequestProps } from './components/IQuoteRequestProps';

export interface IQuoteRequestWebPartProps {
  description: string;
  context:WebPartContext;
  siteUrl:string;
}

export default class QuoteRequestWebPart extends BaseClientSideWebPart<IQuoteRequestWebPartProps> {
  public render(): void {
    const element: React.ReactElement<IQuoteRequestProps> = React.createElement(
      QuoteRequest,
      {
        description: this.properties.description,
        context:this.context,
        spcontext:""
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
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
