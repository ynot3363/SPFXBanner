import { override } from "@microsoft/decorators";
import { Log } from "@microsoft/sp-core-library";
import {
  BaseApplicationCustomizer,
  PlaceholderContent,
  PlaceholderName,
} from "@microsoft/sp-application-base";
import { Dialog } from "@microsoft/sp-dialog";

import styles from "./SpFxBannerApplicationCustomizer.module.scss";

import * as strings from "SpFxBannerApplicationCustomizerStrings";
import CustomDialog from "./CustomDialog";

const LOG_SOURCE: string = "SpFxBannerApplicationCustomizer";

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface ISpFxBannerApplicationCustomizerProperties {
  // This is an example; replace with your own property
  testMessage: string;
}

/** A Custom Action which can be run during execution of a Client Side Application */
export default class SpFxBannerApplicationCustomizer extends BaseApplicationCustomizer<ISpFxBannerApplicationCustomizerProperties> {
  private _topPlaceholder: PlaceholderContent | undefined;
  private _bottomPlaceholder: PlaceholderContent | undefined;
  private modifiedAndReviewedElement: HTMLElement;

  @override
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);

    // Wait for the placeholders to be created (or handle them being changed)
    // and then render.
    this.context.placeholderProvider.changedEvent.add(
      this,
      this._renderPlaceholders
    );

    this.context.application.navigatedEvent.add(this, this._updateData);

    return Promise.resolve();
  }

  private _onDispose(): void {
    Log.info(LOG_SOURCE, "Disposed custom top and bottom placeholders.");
  }

  private _updateData(): void {
    const currentDate: Date = new Date();

    this.modifiedAndReviewedElement = document.getElementById(
      "spce_ModifiedAndReviewed"
    );
    if (this.modifiedAndReviewedElement) {
      let placeholderModified: string = "";
      placeholderModified =
        placeholderModified +
        '<span id="spce_LastModified">' +
        this._getDateFormattedUTC(currentDate) +
        "</span>";

      let placeholderlink: string =
        '<a href="#" id="spce_btnLastReviewed" alt="Set page as reviewed">';
      placeholderlink =
        placeholderlink + '<span id="spce_LastReviewed">Not Set</span></a>';

      const placeholderBody: string = `
      <div id="spce_ModifiedAndReviewed">Last Modified: ${placeholderModified}
      | Last Reviewed: ${placeholderlink}</div>`;

      this.modifiedAndReviewedElement.innerHTML = placeholderBody;

      let currentPage: string = this.context.pageContext.site.serverRequestPath;
      let lastReviewedElement: HTMLElement = document.getElementById(
        "spce_btnLastReviewed"
      );
      lastReviewedElement.setAttribute("pagePath", currentPage);
      lastReviewedElement.addEventListener("click", this._clickeventhandler);
    }
  }

  private _clickeventhandler(): void {
    // DialogManager.instance.alert('This is a test!');
    // alert('Set the page reviewed date to today?');
    let thisElement: HTMLElement = document.getElementById(
      "spce_btnLastReviewed"
    );
    let currentPage: string = thisElement.getAttribute("pagePath");
    try {
      let dialog: CustomDialog = new CustomDialog();
      dialog.itemUrlFromExtension = currentPage;

      dialog.show().then(() => {
        Log.info(
          LOG_SOURCE,
          `Message from Custom Dialog --> ` + dialog.paramFromDialog
        );
      });
    } catch (error) {
      Log.info(LOG_SOURCE, "caught error");
    }
  }

  private _renderPlaceholders(): void {
    console.log(
      "Available placeholders: ",
      this.context.placeholderProvider.placeholderNames
        .map((name) => PlaceholderName[name])
        .join(", ")
    );

    if (!this._topPlaceholder) {
      this._topPlaceholder = this.context.placeholderProvider.tryCreateContent(
        PlaceholderName.Top,
        { onDispose: this._onDispose }
      );

      // The extension should not assume that the expected placeholder is available.
      if (!this._topPlaceholder) {
        console.error("The expected placeholder (Top) was not found.");
        return;
      }

      // if it is available, and access to domElement, update contents
      if (this._topPlaceholder.domElement) {
        this._topPlaceholder.domElement.innerHTML =
          this._getTopPlaceHolderHtml();
      }
    } // !this._topPlaceholder

    if (!this._bottomPlaceholder) {
      this._bottomPlaceholder =
        this.context.placeholderProvider.tryCreateContent(
          PlaceholderName.Bottom,
          { onDispose: this._onDispose }
        );

      // The extension should not assume that the expected placeholder is available.
      if (!this._bottomPlaceholder) {
        console.error("The expected placeholder (Bottom) was not found.");
        return;
      }

      // if it is available, and access to domElement, update contents
      if (this._bottomPlaceholder.domElement) {
        this._bottomPlaceholder.domElement.innerHTML =
          this._getBottomPlaceholderHtml();
      }
    } // !this.bottom_Placeholder
  } // _renderPlaceHolders

  /*
   *
   */
  private _getTopPlaceHolderHtml(): string {
    const placeholderBody: string = `
            <div class="${styles.app}">
                <div style="style="clear:both;>
                  <div style="float:center;">
                    <div id="spce_classification_top" class=${styles.bannas}>SOME VALUE</div>
                  </div>
                </div>
            </div>`;

    return placeholderBody;
  }

  private _getBottomPlaceholderHtml(): string {
    const placeholderBody: string = `
            <div class="${styles.app}">
                <div style="style="clear:both;>
                  <div style="font-size: 1em; width:600px; float:left;">
                    <div id="spce_ModifiedAndReviewed">Last Modified: |Last Reviewed: NOT SET</div>
                  </div>
                </div>
                <br/>
                <div style="style="clear:both;>
                  <div style="float:center;">
                    <div id="spce_bottom" class=${styles.bannas}>SOME VALUE</div>
                  </div>
                </div>
            </div>`;

    return placeholderBody;
  }

  private _getDateFormattedUTC(thisdate: Date): string {
    const dd: number = thisdate.getDate();
    const mm: number = thisdate.getMonth() + 1;
    const yyyy: number = thisdate.getFullYear();
    let day: string;
    let month: string;

    if (dd < 10) {
      day = `0${dd.toString()}`;
    } else {
      day = dd.toString();
    }

    if (mm < 10) {
      month = `0${mm.toString()}`;
    } else {
      month = mm.toString();
    }
    return `${yyyy}${month}${day}`;
  }
}
