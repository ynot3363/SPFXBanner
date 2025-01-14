import { BaseDialog, IDialogConfiguration } from '@microsoft/sp-dialog';

export default class CustomDialog extends BaseDialog {
    public itemUrlFromExtension: string;
    public otherParam: string;
    public paramFromDialog: string;
    public render(): void {
      let html: string = '';
      html += `<div style="padding: 10px;">`;
      html += `<h1>Set the page reviewed date to today?</h1>`;
      html += `<p>Page Url---> <span>` + this.itemUrlFromExtension + `</span></p>`;
      html += `<br>`;
      html += `<br>`;
      html += `<input type="button" id="OkButton" value="Submit">`;
      html += `<input type="button" id="CancelButton" value="Submit">`;
      html += `</div>`;
      this.domElement.innerHTML += html;
      this._setButtonEventHandlers();
    }
    // METHOD TO BIND EVENT HANDLER TO BUTTON CLICK
    private _setButtonEventHandlers(): void {
      // const webPart: CustomDialog = this;
      this.domElement.querySelector('#OkButton').addEventListener('click', () => {
        this.paramFromDialog = 'okButton_clicked';
        this.close();
      });
      this.domElement.querySelector('#CancelButton').addEventListener('click', () => {
        this.paramFromDialog = '';
        this.close();
      });

    }
    public getConfig(): IDialogConfiguration {
      return {
        isBlocking: false
      };
    }
    protected onAfterClose(): void {
      super.onAfterClose();
    }
  }