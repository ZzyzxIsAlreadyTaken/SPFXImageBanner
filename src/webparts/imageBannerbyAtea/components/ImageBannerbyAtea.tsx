import * as React from "react";
import styles from "./ImageBannerbyAtea.module.scss";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/profiles";
import "@pnp/sp/site-users/web";

interface IBannerFileUrl {
  fileAbsoluteUrl: string;
  fileName?: string;
  fileNameWithoutExtension?: string;
}

interface IImageBannerbyAteaProps {
  context: WebPartContext;
  targetGroupId: string;
  bannerFileUrl: IBannerFileUrl | undefined;
  linkUrl: string;
}

export default function ImageBannerbyAtea(
  props: IImageBannerbyAteaProps
): React.ReactElement {
  console.log("bannerFileUrl prop:", props.bannerFileUrl?.fileAbsoluteUrl);
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
      {isInTargetAudience && (
        <div className={styles.targetedContent}>
          {props.bannerFileUrl && (
            <div className={styles.bannerContainer}>
              <a href={props.linkUrl} target="_blank" rel="noopener noreferrer">
                <img
                  alt="Banner"
                  src={props.bannerFileUrl?.fileAbsoluteUrl}
                  className={styles.bannerImage}
                />
              </a>
            </div>
          )}
        </div>
      )}
    </section>
  );
}
