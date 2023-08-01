import { NgElement, WithProperties } from '@angular/elements';
import { LfLoginComponent } from '@laserfiche/types-lf-ui-components';
import * as React from 'react';
import { ISPDocumentData } from '../../Utils/Types';
import { clientId } from '../../webparts/constants';
import LoadingDialog, {
  SavedToLaserficheSuccessDialogButtons,
  SavedToLaserficheSuccessDialogText,
} from './CommonDialogs';
import {
  SaveDocumentToLaserfiche,
  SavedToLaserficheDocumentData,
} from './SaveDocumentToLaserfiche';
import styles from './SendToLaserFiche.module.scss';
import { SPComponentLoader } from '@microsoft/sp-loader';
import * as ReactDOM from 'react-dom';
import { BaseDialog } from '@microsoft/sp-dialog';
import { getRegion } from '../../Utils/Funcs';

export default class SaveToLaserficheCustomDialog extends BaseDialog {
  successful = false;

  handleSuccessSave = (successful: boolean) => {
    this.successful = successful;
  };

  closeClick = async () => {
    await this.close();
    if (this.closeParent) {
      await this.closeParent();
    }
  };

  constructor(
    private spFileData: ISPDocumentData,
    private hadToRouteToLogin: boolean,
    private closeParent?: () => Promise<void>
  ) {
    super();
  }

  public render(): void {
    const element: React.ReactElement = (
      <SaveToLaserficheDialog
        spFileMetadata={this.spFileData}
        successSave={this.handleSuccessSave}
        closeClick={this.closeClick}
        hadToRouteToLogin={this.hadToRouteToLogin}
      />
    );
    ReactDOM.render(element, this.domElement);
  }

  protected onAfterClose(): void {
    ReactDOM.unmountComponentAtNode(this.domElement);
    super.onAfterClose();
    if (this.closeParent) {
      this.closeParent();
    }
  }
}

const CLOSE = 'Close';
function SaveToLaserficheDialog(props: {
  successSave: (success: boolean) => void;
  closeClick: () => Promise<void>;
  spFileMetadata: ISPDocumentData;
  hadToRouteToLogin: boolean;
}) {
  const loginComponent = React.createRef<
    NgElement & WithProperties<LfLoginComponent>
  >();

  const region = getRegion();
  const [success, setSuccess] = React.useState<
    SavedToLaserficheDocumentData | undefined
  >();

  const saveToDialogCloseClick = async () => {
    await props.closeClick();
  };

  React.useEffect(() => {
    SPComponentLoader.loadScript(
      'https://cdn.jsdelivr.net/npm/zone.js@0.11.4/bundles/zone.umd.min.js'
    ).then(() => {
      SPComponentLoader.loadScript(
        'https://cdn.jsdelivr.net/npm/@laserfiche/lf-ui-components@14/cdn/lf-ui-components.js'
      ).then(async () => {
        if (loginComponent.current?.authorization_credentials) {
          const saveToLF = new SaveDocumentToLaserfiche(props.spFileMetadata);
          const successSaveToLF =
            await saveToLF.trySaveDocumentToLaserficheAsync();
          if (successSaveToLF) {
            props.successSave(true);
            setSuccess(successSaveToLF);
          } else {
            // TODO is this handled correctly when an error occured when deleting the file
            props.successSave(false);
            props.closeClick();
          }
        } else {
          props.successSave(false);
          props.closeClick();
        }
      });
    });
  }, []);

  return (
    <div className={styles.wrapper}>
      <div className={styles.header}>
        <div className={styles.logoHeader}>
          <lf-login
            hidden
            redirect_uri=''
            authorize_url_host_name={region}
            redirect_behavior='Replace'
            client_id={clientId}
            ref={loginComponent}
          />
          <img
            src='data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAALQAAAC0CAMAAAAKE/YAAAAAUVBMVEXSXyj////HYzL/+/T/+Or/9d+yaUa9ZT2yaUj/9OG7Zj3SXybRYCj/+/b///3LYS/OYCvEZDS2aEL/89jAZTnMYS3/8dO7Zzusa02+ZTn/78wyF0DsAAABnUlEQVR4nO3ci26CMABGYQcoLRS5OTf2/g86R+KSLYUm2vxcPB8RTYzxkADRajkcAAAAAAAAAADYgbJcusCvqdtLnhfeJR/a96X7vOriarNJ/cUtHeiTnI7p26TsY+XRZ190sXSfVyA6X7rP6xZdzeweREeTGDt3IBIdTeCUR3Q0wQOxLNf3CWSr0ZvcPYiWIFqFaBWiVYhWIVqFaBWiVYhWIVqFaBWiVYhWIVqFaBWiVYhWIVqFaBWiVYhWIVqFaBWiVYhWIVqFaBWiVYhWIVqFaBWiVV4zeok/379m9BL2HO1Ckymlky0jRQc3Kqoou4f6YHzdaLX56PRzak757/JjfDS0dbOK6HM6Paf8P3st6lVE/9mAwPOpNcnqokOIJppoookmmmiiiSaaaKKJ3k30OfTFdU3RXZ+lT6qq6rbO+k4VXQ9fvT2OrH30Zo+3u/5rUI17NO3QmdPImIduxoyrUze0khEm5w6uqZNIRKNi91Hl5661dH+tdow6wts5J//BaJPRwH6IT1NxbDJ6vVc+nrXJaAAAAADALn0DBosqnCStFi4AAAAASUVORK5CYII='
            width='30'
            height='30'
          />
          <p className={styles.dialogTitle}>Laserfiche</p>
        </div>

        <button
          className={styles.lfCloseButton}
          title='close'
          onClick={saveToDialogCloseClick}
        >
          <span className='material-icons-outlined'> close </span>
        </button>
      </div>

      <div className={styles.contentBox}>
        {!success && <LoadingDialog />}
        {success && (
          <SavedToLaserficheSuccessDialogText successfulSave={success} />
        )}
      </div>

      <div className={styles.footer}>
        {success && (
          <SavedToLaserficheSuccessDialogButtons
            successfulSave={success}
            closeClick={saveToDialogCloseClick}
            hadToRouteToLogin={props.hadToRouteToLogin}
          />
        )}
        {!success && (
          <button
            onClick={saveToDialogCloseClick}
            className='lf-button sec-button'
          >
            {CLOSE}
          </button>
        )}
      </div>
    </div>
  );
}
