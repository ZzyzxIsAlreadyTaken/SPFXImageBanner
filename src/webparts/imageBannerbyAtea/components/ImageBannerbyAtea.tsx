import * as React from "react";
import styles from "./ImageBannerbyAtea.module.scss";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/profiles";
import "@pnp/sp/site-users/web";
import { IFilePickerResult } from "@pnp/spfx-property-controls/lib/PropertyFieldFilePicker";
import { DisplayMode } from "@microsoft/sp-core-library";

interface IImageBannerbyAteaProps {
  context: WebPartContext;
  targetGroupId: string;
  filePickerResult: IFilePickerResult;
  linkUrl: string;
  displayMode: DisplayMode;
}

export default function ImageBannerbyAtea(
  props: IImageBannerbyAteaProps
): React.ReactElement {
  const [isInTargetAudience, setIsInTargetAudience] =
    React.useState<boolean>(false);
  const [error, setError] = React.useState<string | null>(null);

  React.useEffect(() => {
    const checkAudience = async (): Promise<void> => {
      try {
        // Initialize PnPjs with SPFx context
        const sp = spfi().using(SPFx(props.context));

        // Get user's groups
        const userGroups = await sp.web.siteUsers
          .getById(props.context.pageContext.legacyPageContext.userId)
          .groups();

        // Check if user is in the selected group
        const isInGroup = userGroups.some(
          (group: { Id: number }) =>
            group.Id.toString() === (props.targetGroupId || "")
        );

        setIsInTargetAudience(isInGroup);
      } catch (err) {
        console.error("Error checking audience:", err);
        setError("Failed to check audience membership");
      }
    };

    if (props.context && props.context.pageContext) {
      // eslint-disable-next-line @typescript-eslint/no-floating-promises
      checkAudience();
    }
  }, [props.context, props.targetGroupId]);

  return (
    <section className={styles.imageBannerbyAtea}>
      {error && <div className={styles.error}>{error}</div>}

      {(props.displayMode === DisplayMode.Edit || isInTargetAudience) && (
        <>
          {props.filePickerResult ? (
            <div className={styles.bannerContainer}>
              <a href={props.linkUrl} target="_blank" rel="noopener noreferrer">
                <img
                  alt="Banner"
                  src={props.filePickerResult.fileAbsoluteUrl}
                  className={styles.bannerImage}
                />
              </a>
            </div>
          ) : props.displayMode === DisplayMode.Edit ? (
            <div>🔧 Velg et bilde for å vise forsidebanneret.</div>
          ) : null}
        </>
      )}
    </section>
  );
}
