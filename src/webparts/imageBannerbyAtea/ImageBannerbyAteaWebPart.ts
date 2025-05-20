import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import {
  type IPropertyPaneConfiguration,
  PropertyPaneDropdown,
  PropertyPaneTextField,
} from "@microsoft/sp-property-pane";
import {
  BaseClientSideWebPart,
  WebPartContext,
} from "@microsoft/sp-webpart-base";
import { spfi, SPFx, SPFI } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/site-groups";
import "@pnp/sp/site-groups/web";
import ImageBannerbyAtea from "./components/ImageBannerbyAtea";
import {
  PropertyFieldFilePicker,
  IFilePickerResult,
} from "@pnp/spfx-property-controls/lib/PropertyFieldFilePicker";

interface IBannerFileUrl {
  fileAbsoluteUrl: string;
  fileName?: string;
  fileNameWithoutExtension?: string;
}

interface ISharePointGroup {
  Id: number;
  Title: string;
}

// Add SP namespace type definitions
declare global {
  interface Window {
    SP: {
      PickerDialog: new () => unknown;
      PickerTabInfo: new (type: unknown) => unknown;
      PickerTabType: {
        images: unknown;
      };
      PickerDialogEventType: {
        FilesSelected: unknown;
      };
    };
  }
}

export interface IImageBannerbyAteaWebPartProps {
  description: string;
  targetGroupId: string;
  bannerFileUrl: IBannerFileUrl | undefined;
  linkUrl: string;
}

interface IImageBannerbyAteaProps {
  context: WebPartContext;
  targetGroupId: string;
  bannerFileUrl: IBannerFileUrl | undefined;
  linkUrl: string;
}

export default class ImageBannerbyAteaWebPart extends BaseClientSideWebPart<IImageBannerbyAteaWebPartProps> {
  private _groups: { key: string; text: string }[] = [];

  public render(): void {
    const element: React.ReactElement<IImageBannerbyAteaProps> =
      React.createElement(ImageBannerbyAtea, {
        context: this.context,
        targetGroupId: this.properties.targetGroupId || "",
        bannerFileUrl: this.properties.bannerFileUrl || undefined,
        linkUrl: this.properties.linkUrl || "",
      });

    ReactDom.render(element, this.domElement);
  }

  protected async onInit(): Promise<void> {
    await super.onInit();
    // Initialize PnPjs
    const sp = spfi().using(SPFx(this.context));
    await this._loadGroups(sp);
  }

  private async _loadGroups(sp: SPFI): Promise<void> {
    try {
      const groups = await sp.web.siteGroups();
      this._groups = groups.map((group: ISharePointGroup) => ({
        key: group.Id.toString(),
        text: group.Title,
      }));

      // If no group is selected, select the first one
      if (!this.properties.targetGroupId && this._groups.length > 0) {
        this.properties.targetGroupId = this._groups[0].key;
      }

      this.context.propertyPane.refresh();
    } catch (error) {
      console.error("Error loading groups:", error);
    }
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
    super.onDispose();
  }

  protected get dataVersion(): Version {
    return Version.parse("1.0");
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: "Konfigurer bannerinnstillinger",
          },
          groups: [
            {
              groupName: "Målgruppe",
              groupFields: [
                PropertyPaneDropdown("targetGroupId", {
                  label: "Målgruppe",
                  options: this._groups,
                  selectedKey: this.properties.targetGroupId || "",
                }),
              ],
            },
            {
              groupName: "Bannerinnstillinger",
              groupFields: [
                PropertyFieldFilePicker("bannerFileUrl", {
                  // eslint-disable-next-line @typescript-eslint/no-explicit-any
                  context: this.context as any,
                  filePickerResult: this.properties.bannerFileUrl
                    ? {
                        fileAbsoluteUrl:
                          this.properties.bannerFileUrl.fileAbsoluteUrl,
                        fileName: this.properties.bannerFileUrl.fileName || "",
                        fileNameWithoutExtension:
                          this.properties.bannerFileUrl
                            .fileNameWithoutExtension || "",
                        downloadFileContent: async () => new File([], ""),
                      }
                    : {
                        fileAbsoluteUrl: "",
                        fileName: "",
                        fileNameWithoutExtension: "",
                        downloadFileContent: async () => new File([], ""),
                      },
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  key: "bannerFileUrl",
                  label: "Velg banner-bilde",
                  buttonLabel: "Velg bilde",
                  onSave: (filePickerResult: IFilePickerResult) => {
                    if (filePickerResult && filePickerResult.fileAbsoluteUrl) {
                      this.properties.bannerFileUrl = filePickerResult;
                    }
                  },
                  onChanged: (filePickerResult: IFilePickerResult) => {
                    if (filePickerResult && filePickerResult.fileAbsoluteUrl) {
                      this.properties.bannerFileUrl = filePickerResult;
                    }
                  },
                  accepts: [".jpg", ".jpeg", ".png", ".gif"],
                  buttonIcon: "Image",
                }),
                PropertyPaneTextField("linkUrl", {
                  label: "Lenke til banner",
                  description: "Skriv inn URL som banner skal lenke til",
                }),
              ],
            },
          ],
        },
      ],
    };
  }
}
