import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'HelloWorldWebPartStrings';
import HelloWorld from './components/HelloWorld';
import { IHelloWorldProps } from './components/IHelloWorldProps';
import { IUserProps } from './components/IUserProps';
import MockHttpClient from './MockHttpClient';
import { ISPList } from './ISPList';
import { Environment, EnvironmentType } from '@microsoft/sp-core-library';
import { SPHttpClient, MSGraphClientFactory } from '@microsoft/sp-http';
import { MSGraphClient } from "@microsoft/sp-client-preview";

export interface IHelloWorldWebPartProps {
  description: string;
}

export default class HelloWorldWebPart extends BaseClientSideWebPart<IHelloWorldWebPartProps> {

  private _getSharePointListData(): Promise<ISPList[]> {
    const url: string = this.context.pageContext.web.absoluteUrl + `/_api/web/siteusers?`;

    return this.context.spHttpClient.get(url, SPHttpClient.configurations.v1)
      .then(response => {
        console.log("json response: ", response);
        return response.json();
      })
      .then(json => {
        return json.value;
      }) as Promise<ISPList[]>;
   }
   
   private _getListData(): Promise<ISPList[]> {
      if(Environment.type === EnvironmentType.Local) {
         return this._getMockListData();
      }
      else {
        return this._getSharePointListData();
      }
   }
  
  private _getMockListData(): Promise<ISPList[]> {
    return MockHttpClient.get(this.context.pageContext.web.absoluteUrl)
        .then((data: ISPList[]) => {
             return data;
         });
  }

  public render(): void {
    const element: React.ReactElement<IUserProps> = React.createElement(HelloWorld, 
      {
        graphClient: this.context.serviceScope.consume(MSGraphClient.serviceKey)
     }
    );
    ReactDom.render(element, this.domElement);
    /*
    this.context.msGraphClientFactory
    .getClient()
    .then((client: MSGraphClient): void => {
      // get information about the current user from the Microsoft Graph
      client
        .api('/users')
        .get((error, res: any, rawResponse?: any) => {
          console.log(res);
      }).then(res => {
        console.log("res2", res);
      });
    });
    */
    /*
    this.context.msGraphClientFactory
    .getClient()
    .then((client: MSGraphClient): void => {
      client
      .api('/me')
      .get((err, res) => {
        const element: React.ReactElement<IHelloWorldProps> = React.createElement(HelloWorld, {
          description: this.properties.description,
          lists: res
        });
        console.log("Results:", res); // prints info about authenticated user
        console.log("Error:", err); // prints info about authenticated user
     });
    });
    */
    /*
    this._getListData().then(lists => {
      const element: React.ReactElement<IHelloWorldProps> = React.createElement(HelloWorld, {
        description: this.properties.description,
        lists: lists
      });

      ReactDom.render(element, this.domElement);
    });
    */
    
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
