import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import {
  IPropertyPaneConfiguration,
  IPropertyPaneDropdownOption,
  PropertyPaneDropdown,
} from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import { IReadonlyTheme } from "@microsoft/sp-component-base";
import { SPHttpClient } from "@microsoft/sp-http";

import * as strings from "JciTkbBrowseMenuWebPartStrings";
import { IJciTkbBrowseMenuProps } from "./components/IJciTkbBrowseMenuProps";
import JciTkbBrowseMenuList from "./components/JciTkbBrowseMenu";
import { PropertyFieldMultiSelect } from "@pnp/spfx-property-controls/lib/PropertyFieldMultiSelect";

export interface IJciTkbBrowseMenuWebPartProps {
  description: string;
  selectedJciTkbBrowseMenus: string[];
  selectedGroupId: string;
}

export default class JciTkbBrowseMenuWebPart extends BaseClientSideWebPart<IJciTkbBrowseMenuWebPartProps> {
  private termSetOptions: IPropertyPaneDropdownOption[] = [];
  private termStoreGroupOptions: IPropertyPaneDropdownOption[] = [];

  public async render(): Promise<void> {
    const element: React.ReactElement<IJciTkbBrowseMenuProps> =
      React.createElement(JciTkbBrowseMenuList, {
        context: this.context,
        groupId: this.properties.selectedGroupId,
        setNames: this.properties.selectedJciTkbBrowseMenus,
      });

    ReactDom.render(element, this.domElement);
  }

  private async getJciTkbBrowseMenus(groupId: string): Promise<void> {
    const url = `${this.context.pageContext.web.absoluteUrl}/_api/v2.1/termstore/groups/${groupId}/sets`;
    const response = await this.context.spHttpClient.get(
      url,
      SPHttpClient.configurations.v1
    );
    const data = await response.json();

    // Populate the dropdown options
    this.termSetOptions = [
      {
        key: "",
        text: "Select Set",
      },
      ...data.value.map((termSet: any) => {
        return {
          key: termSet.localizedNames[0].name,
          text: termSet.localizedNames[0].name,
        };
      }),
    ];
  }
  protected onInit(): Promise<void> {
    return this.getTermStoreGroups();
  }

  private async getTermStoreGroups(): Promise<void> {
    const url = `${this.context.pageContext.web.absoluteUrl}/_api/v2.1/termstore/groups`;
    const response = await this.context.spHttpClient.get(
      url,
      SPHttpClient.configurations.v1
    );
    const data = await response.json();

    this.termStoreGroupOptions = data.value.map((group: any) => {
      return {
        key: group.id,
        text: group.name,
      };
    });
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    const { semanticColors } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty(
        "--bodyText",
        semanticColors.bodyText || null
      );
      this.domElement.style.setProperty("--link", semanticColors.link || null);
      this.domElement.style.setProperty(
        "--linkHovered",
        semanticColors.linkHovered || null
      );
    }
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse("1.0");
  }

  protected async onPropertyPaneConfigurationStart(): Promise<void> {
    // Load the term store groups if not already loaded
    if (this.termStoreGroupOptions.length === 0) {
      await this.getTermStoreGroups();
    }

    if (this.termSetOptions.length === 0 && this.properties.selectedGroupId) {
      await this.getJciTkbBrowseMenus(this.properties.selectedGroupId);
    }

    this.context.propertyPane.refresh();
  }

  protected async onPropertyPaneFieldChanged(
    propertyPath: string,
    oldValue: any,
    newValue: any
  ): Promise<void> {
    if (propertyPath === "selectedGroupId" && newValue) {
      await this.getJciTkbBrowseMenus(newValue);
      this.context.propertyPane.refresh();
    }
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: "Configure your side navigation",
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneDropdown("selectedGroupId", {
                  label: "Select Term Store Group",
                  options: this.termStoreGroupOptions,
                  selectedKey: this.properties.selectedGroupId,
                }),
                PropertyFieldMultiSelect("selectedJciTkbBrowseMenus", {
                  key: "selectedJciTkbBrowseMenus",
                  label: "Select Term Sets",
                  options: this.termSetOptions,
                  selectedKeys: this.properties.selectedJciTkbBrowseMenus,
                }),
              ],
            },
          ],
        },
      ],
    };
  }
}
