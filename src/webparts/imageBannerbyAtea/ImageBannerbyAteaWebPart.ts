import * as React from "react";
import * as ReactDom from "react-dom";
import { DisplayMode, Version } from "@microsoft/sp-core-library";
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
import "@pnp/sp/folders";
import "@pnp/sp/files/folder";
import "@pnp/sp/files";
import ImageBannerbyAtea from "./components/ImageBannerbyAtea";
import {
  PropertyFieldFilePicker,
  IFilePickerResult,
} from "@pnp/spfx-property-controls/lib/PropertyFieldFilePicker";

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
  linkUrl: string;
  filePickerResult: IFilePickerResult;
  displayMode: DisplayMode;
}

interface IImageBannerbyAteaProps {
  context: WebPartContext;
  targetGroupId: string;
  linkUrl: string;
  filePickerResult: IFilePickerResult;
  displayMode: DisplayMode;
}

export default class ImageBannerbyAteaWebPart extends BaseClientSideWebPart<IImageBannerbyAteaWebPartProps> {
  private _groups: { key: string; text: string }[] = [];

  public render(): void {
    const element: React.ReactElement<IImageBannerbyAteaProps> =
      React.createElement(ImageBannerbyAtea, {
        context: this.context,
        targetGroupId: this.properties.targetGroupId || "",
        filePickerResult: this.properties.filePickerResult,
        linkUrl: this.properties.linkUrl || "",
        displayMode: this.displayMode,
      });

    ReactDom.render(element, this.domElement);
  }

  protected async onInit(): Promise<void> {
    try {
      await super.onInit();
      // Initialize PnPjs
      const sp = spfi().using(SPFx(this.context));
      await this._loadGroups(sp);
    } catch (error) {
      console.error("onInit error:", error);
    }
  }

  private async _loadGroups(sp: SPFI): Promise<void> {
    try {
      const groups = await sp.web.siteGroups();
      this._groups = groups
        .filter(
          (group: ISharePointGroup) =>
            !group.Title.startsWith("Limited Access System Group") &&
            !group.Title.startsWith("SharingLinks")
        )
        .map((group: ISharePointGroup) => ({
          key: group.Id.toString(),
          text: group.Title,
        }));

      // If no group is selected, select the first one
      if (!this.properties.targetGroupId && this._groups.length > 0) {
        this.properties.targetGroupId = this._groups[0].key;
      }

      this.context.propertyPane.refresh();
    } catch (error) {
      console.error("_loadGroups error:", error);
      throw error; // Re-throw to be caught by onInit
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
                PropertyFieldFilePicker("filePickerResult", {
                  // eslint-disable-next-line @typescript-eslint/no-explicit-any
                  context: this.context as any,
                  filePickerResult: this.properties.filePickerResult,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  key: "filePickerResult",
                  label: "Velg banner-bilde",
                  buttonLabel: "Velg bilde",
                  onSave: (e: IFilePickerResult) => {
                    if (!e.fileAbsoluteUrl) {
                      (async () => {
                        try {
                          const fileContent = await e.downloadFileContent();
                          const sp = spfi().using(SPFx(this.context));
                          const folderUrl = `${this.context.pageContext.site.serverRelativeUrl}/SiteAssets/SitePages/bildebanner`;
                          const uploadResult = await sp.web
                            .getFolderByServerRelativePath(folderUrl)
                            .files.addUsingPath(e.fileName, fileContent, {
                              Overwrite: true,
                            });

                          // Clone the filePickerResult to avoid possible race conditions
                          const updatedResult = {
                            ...e,
                            fileAbsoluteUrl: uploadResult.ServerRelativeUrl,
                          };
                          this.properties.filePickerResult = updatedResult;
                          this.context.propertyPane.refresh();
                          this.render();
                        } catch (err) {
                          console.error("File upload failed:", err);
                        }
                      })().catch(console.error);
                    } else {
                      this.properties.filePickerResult = e;
                      this.context.propertyPane.refresh();
                      this.render();
                    }
                  },
                  onChanged: (e: IFilePickerResult) => {
                    this.properties.filePickerResult = e;
                    this.context.propertyPane.refresh();
                    this.render();
                  },
                  accepts: [".jpg", ".jpeg", ".png", ".gif"],
                  buttonIcon: "FabricPictureLibrary",
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
