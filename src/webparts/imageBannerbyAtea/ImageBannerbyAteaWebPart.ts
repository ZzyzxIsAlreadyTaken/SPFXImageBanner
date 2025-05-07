import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import {
  type IPropertyPaneConfiguration,
  PropertyPaneDropdown,
} from "@microsoft/sp-property-pane";
import {
  BaseClientSideWebPart,
  WebPartContext,
} from "@microsoft/sp-webpart-base";
import { spfi, SPFx } from "@pnp/sp";
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

// Add SP namespace type definitions
declare global {
  interface Window {
    SP: {
      PickerDialog: new () => any;
      PickerTabInfo: new (type: any) => any;
      PickerTabType: {
        images: any;
      };
      PickerDialogEventType: {
        FilesSelected: any;
      };
    };
  }
}

export interface IImageBannerbyAteaWebPartProps {
  description: string;
  targetGroupId: string;
  bannerFileUrl: IBannerFileUrl | null;
}

interface IImageBannerbyAteaProps {
  context: WebPartContext;
  targetGroupId: string;
  bannerFileUrl: IBannerFileUrl | null;
}

export default class ImageBannerbyAteaWebPart extends BaseClientSideWebPart<IImageBannerbyAteaWebPartProps> {
  private _groups: { key: string; text: string }[] = [];

  public render(): void {
    const element: React.ReactElement<IImageBannerbyAteaProps> =
      React.createElement(ImageBannerbyAtea, {
        context: this.context,
        targetGroupId: this.properties.targetGroupId || "",
        bannerFileUrl: this.properties.bannerFileUrl || null,
      });

    ReactDom.render(element, this.domElement);
  }

  protected async onInit(): Promise<void> {
    await super.onInit();
    // Initialize PnPjs
    const sp = spfi().using(SPFx(this.context));
    await this._loadGroups(sp);
  }

  private async _loadGroups(sp: any): Promise<void> {
    try {
      const groups = await sp.web.siteGroups();
      this._groups = groups.map((group: any) => ({
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
            description: "Configure your banner settings",
          },
          groups: [
            {
              groupName: "Targeting",
              groupFields: [
                PropertyPaneDropdown("targetGroupId", {
                  label: "Target Group",
                  options: this._groups,
                  selectedKey: this.properties.targetGroupId || "",
                }),
              ],
            },
            {
              groupName: "Banner Settings",
              groupFields: [
                PropertyFieldFilePicker("bannerFileUrl", {
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
                  label: "Select Banner Image",
                  buttonLabel: "Select Image",
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
              ],
            },
          ],
        },
      ],
    };
  }
}
