import { BaseDialog, IDialogConfiguration } from '@microsoft/sp-dialog';
import styles from './SendToLaserficheLoginComponent.module.scss';
import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { TempStorageKeys } from '../../../Utils/Enums';
import { Navigation } from 'spfx-navigation';

export default class SendToLaserficheCustomDialog extends BaseDialog {
  isLoading: boolean = true;
  metadataSaved = false;
  fileLink?: string;

  public render(): void {
    ReactDOM.render(
      <SendToLaserficheDialog
        loading={this.isLoading}
        metadataSaved={this.metadataSaved}
        fileLink={this.fileLink}
        closeDialog={() => this.close()}
      />,
      this.domElement
    );
  }
  public getConfig(): IDialogConfiguration {
    return {
      isBlocking: false,
    };
  }
  protected onAfterClose(): void {
    super.onAfterClose();
  }
}

function SendToLaserficheDialog(props: {
  loading: boolean;
  metadataSaved: boolean;
  fileLink?: string;
  closeDialog: () => void;
}) {
  function viewFile() {
    window.open(props.fileLink);
  }

  function redirect() {
    const pageOrigin = window.localStorage.getItem(TempStorageKeys.PageOrigin);
    const Fileurl = window.localStorage.getItem(TempStorageKeys.Fileurl);
    const Filenamewithext1 = window.localStorage.getItem(
      TempStorageKeys.Filename
    );
    const fileeee = Fileurl.replace(Filenamewithext1, '');
    const path = pageOrigin + fileeee;
    Navigation.navigate(path, true);
  }

  return (
    <div className="${styles['maindialog']}">
      <div id='overlay' className='${styles.overlay}'></div>
      <img
        src='data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAALQAAAC0CAMAAAAKE/YAAAAAUVBMVEXSXyj////HYzL/+/T/+Or/9d+yaUa9ZT2yaUj/9OG7Zj3SXybRYCj/+/b///3LYS/OYCvEZDS2aEL/89jAZTnMYS3/8dO7Zzusa02+ZTn/78wyF0DsAAABnUlEQVR4nO3ci26CMABGYQcoLRS5OTf2/g86R+KSLYUm2vxcPB8RTYzxkADRajkcAAAAAAAAAADYgbJcusCvqdtLnhfeJR/a96X7vOriarNJ/cUtHeiTnI7p26TsY+XRZ190sXSfVyA6X7rP6xZdzeweREeTGDt3IBIdTeCUR3Q0wQOxLNf3CWSr0ZvcPYiWIFqFaBWiVYhWIVqFaBWiVYhWIVqFaBWiVYhWIVqFaBWiVYhWIVqFaBWiVYhWIVqFaBWiVYhWIVqFaBWiVYhWIVqFaBWiVV4zeok/379m9BL2HO1Ckymlky0jRQc3Kqoou4f6YHzdaLX56PRzak757/JjfDS0dbOK6HM6Paf8P3st6lVE/9mAwPOpNcnqokOIJppoookmmmiiiSaaaKKJ3k30OfTFdU3RXZ+lT6qq6rbO+k4VXQ9fvT2OrH30Zo+3u/5rUI17NO3QmdPImIduxoyrUze0khEm5w6uqZNIRKNi91Hl5661dH+tdow6wts5J//BaJPRwH6IT1NxbDJ6vVc+nrXJaAAAAADALn0DBosqnCStFi4AAAAASUVORK5CYII='
        width='42'
        height='42'
      />
      {props.loading && (
        <>
          <img
            style={{ marginLeft: '28%' }}
            src='/_layouts/15/images/progress.gif'
            id='imgid'
          />
          <div>
            <p className='text' id='it'>
              Saving your document to Laserfiche
            </p>
          </div>
        </>
      )}
      {!props.loading && (
        <>
          <div>
            <p className='text' id='it'>
              {props.metadataSaved
                ? 'Document uploaded'
                : 'Document uploaded to repository, updating metadata failed due to constraint mismatch<br/> <p style="color:red">The Laserfiche template and fields were not applied to this document</p>'}
            </p>
          </div>

          <div id='divid' className='button'>
            <button id='divid1' className='button1' onClick={props.closeDialog}>
              Close
            </button>
            {props.fileLink && (
              <button
                id='divid13'
                className='button2'
                title='Click here to view the file in Laserfiche'
                onClick={viewFile}
              >
                Go to File
              </button>
            )}
            <button
              id='divid14'
              className='button2'
              title='Click here to go back to your SharePoint library'
              onClick={redirect}
            >
              Go to Library
            </button>
          </div>
        </>
      )}
    </div>
  );
}
