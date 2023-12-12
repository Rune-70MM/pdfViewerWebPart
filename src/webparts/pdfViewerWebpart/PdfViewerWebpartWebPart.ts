import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'PdfViewerWebpartWebPartStrings';
import PdfViewerWebpart from './components/PdfViewerWebpart';
import { IPdfViewerWebpartProps } from './components/IPdfViewerWebpartProps';
import { DataAccessService, IDataAccessService } from '../../sp/services/data-access-service';

export interface IPdfViewerWebpartWebPartProps {
  description: string;
}

export default class PdfViewerWebpartWebPart extends BaseClientSideWebPart<IPdfViewerWebpartWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';


  private _dataAccessService: IDataAccessService = null;
 

  public render(): void {
    const element: React.ReactElement<IPdfViewerWebpartProps> = React.createElement(
      PdfViewerWebpart,
      {
        description: this.properties.description,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        dataService: this._dataAccessService,
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> 
  {
    this._environmentMessage = this._getEnvironmentMessage();

    this._dataAccessService = this.context.serviceScope.consume(DataAccessService.serviceKey);

    
    return super.onInit()
  }



  private _getEnvironmentMessage(): string 
  {
    if (!!this.context.sdks.microsoftTeams)
    { //running in Teams
      return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
    }

    return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment;
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }

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
