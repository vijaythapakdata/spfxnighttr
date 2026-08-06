import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
// import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'SpfxformWebPartStrings';
import Spfxform from './components/Spfxform';
import { ISpfxformProps } from './components/ISpfxformProps';
import GetChoiceApi from '../../Service/ChoiceServiceApi';
export interface ISpfxformWebPartProps {
  description: string;
}

export default class SpfxformWebPart extends BaseClientSideWebPart<ISpfxformWebPartProps> {
private choiceServiceClass!:GetChoiceApi;
  protected async onInit(): Promise<void> {
    this.choiceServiceClass=new GetChoiceApi(this.context);
    return super.onInit();
  }
  public async render(): Promise<void> {
    const element: React.ReactElement<ISpfxformProps> = React.createElement(
      Spfxform,
      {
       context:this.context,
       siteurl:this.context.pageContext.web.absoluteUrl,
       departmentOptions:await this.choiceServiceClass.getChoiceValues(this.context.pageContext.web.absoluteUrl,"Department"),
        genderOptions:await this.choiceServiceClass.getChoiceValues(this.context.pageContext.web.absoluteUrl,"Gender"),
       skillsOptions:await this.choiceServiceClass.getChoiceValues(this.context.pageContext.web.absoluteUrl,"Skills"),

       cityOptions:await this.choiceServiceClass.getLookupValues()
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
