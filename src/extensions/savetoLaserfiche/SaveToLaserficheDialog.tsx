import { NgElement, WithProperties } from '@angular/elements';
import { LfLoginComponent } from '@laserfiche/types-lf-ui-components';
import * as React from 'react';
import { ISPDocumentData } from '../../Utils/Types';
import { clientId } from '../../webparts/constants';
import LoadingDialog, { SavedToLaserficheSuccessDialog } from './CommonDialogs';
import { SaveDocumentToLaserfiche } from './SaveDocumentToLaserfiche';
import styles from './SendToLaserFiche.module.scss';
import { SPComponentLoader } from '@microsoft/sp-loader';
import * as ReactDOM from 'react-dom';
import { BaseDialog, IDialogConfiguration } from '@microsoft/sp-dialog';

export default class SaveToLaserficheCustomDialog extends BaseDialog {
  successful = false;

  handleCloseClickAsync = async (success: boolean) => {
    this.successful = success;
    await this.close();
  };

  constructor(private spFileData: ISPDocumentData) {
    // TODO is there a way to make this not blocking and send success true if it has been saved successfully
    super({
      isBlocking: true,
    });
  }

  public render(): void {
    const element: React.ReactElement = (
      <SaveToLaserficheDialog
        spFileMetadata={this.spFileData}
        closeClick={this.handleCloseClickAsync}
      />
    );
    ReactDOM.render(element, this.domElement);
  }

  public getConfig(): IDialogConfiguration {
    return {
      isBlocking: true,
    };
  }

  protected onAfterClose(): void {
    ReactDOM.unmountComponentAtNode(this.domElement);
    super.onAfterClose();
  }
}

function SaveToLaserficheDialog(props: {
  closeClick: (success: boolean) => Promise<void>;
  spFileMetadata: ISPDocumentData;
}) {
  const loginComponent = React.createRef<
    NgElement & WithProperties<LfLoginComponent>
  >();

  const [success, setSuccess] = React.useState<
    { fileLink: string; pathBack: string; metadataSaved: boolean } | undefined
  >();

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
            setSuccess(successSaveToLF);
          } else {
            // TODO is this handled correctly when an error occured when deleting the file
            await props.closeClick(false);
          }
        } else {
          await props.closeClick(false);
        }
      });
    });
  }, []);

  return (
    <div className={styles.maindialog}>
      <lf-login
        hidden
        redirect_uri='https://lfdevm365.sharepoint.com/sites/TestSite/Shared%20Documents/Forms/AllItems.aspx'
        authorize_url_host_name='a.clouddev.laserfiche.com'
        redirect_behavior='Replace'
        client_id={clientId}
        ref={loginComponent}
      />
      <div id='overlay' className={styles.overlay} />
      <div>
        <img
          src='data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAALQAAAC0CAMAAAAKE/YAAAAAUVBMVEXSXyj////HYzL/+/T/+Or/9d+yaUa9ZT2yaUj/9OG7Zj3SXybRYCj/+/b///3LYS/OYCvEZDS2aEL/89jAZTnMYS3/8dO7Zzusa02+ZTn/78wyF0DsAAABnUlEQVR4nO3ci26CMABGYQcoLRS5OTf2/g86R+KSLYUm2vxcPB8RTYzxkADRajkcAAAAAAAAAADYgbJcusCvqdtLnhfeJR/a96X7vOriarNJ/cUtHeiTnI7p26TsY+XRZ190sXSfVyA6X7rP6xZdzeweREeTGDt3IBIdTeCUR3Q0wQOxLNf3CWSr0ZvcPYiWIFqFaBWiVYhWIVqFaBWiVYhWIVqFaBWiVYhWIVqFaBWiVYhWIVqFaBWiVYhWIVqFaBWiVYhWIVqFaBWiVYhWIVqFaBWiVV4zeok/379m9BL2HO1Ckymlky0jRQc3Kqoou4f6YHzdaLX56PRzak757/JjfDS0dbOK6HM6Paf8P3st6lVE/9mAwPOpNcnqokOIJppoookmmmiiiSaaaKKJ3k30OfTFdU3RXZ+lT6qq6rbO+k4VXQ9fvT2OrH30Zo+3u/5rUI17NO3QmdPImIduxoyrUze0khEm5w6uqZNIRKNi91Hl5661dH+tdow6wts5J//BaJPRwH6IT1NxbDJ6vVc+nrXJaAAAAADALn0DBosqnCStFi4AAAAASUVORK5CYII='
          width='42'
          height='42'
        />
      </div>
      {!success && <LoadingDialog />}
      {success && (
        <SavedToLaserficheSuccessDialog
          successfulSave={success}
          closeClick={() => props.closeClick(true)}
        />
      )}
      {/* TODO error dialog */}
    </div>
  );
}
