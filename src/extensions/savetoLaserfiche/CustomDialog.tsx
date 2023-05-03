import { BaseDialog, IDialogConfiguration } from '@microsoft/sp-dialog';
import styles from './SendToLaserFiche.module.scss';
//import {} from './../../../lib/logo/laserfiche-logo.png';
import * as ReactDOM from 'react-dom';
import * as React from 'react';

export default class CustomDailog extends BaseDialog {
  textInside: JSX.Element = (<span>Saving your document to Laserfiche</span>);
  isLoading = true;

  public render(): void {
    const element: React.ReactElement = (
      <CustomDialog
        textInside={this.textInside}
        loading={this.isLoading}
        handleCloseClick={() => this.close()}
      />
    );
    ReactDOM.render(element, this.domElement);
  }
  public getConfig(): IDialogConfiguration {
    return {
      isBlocking: false,
    };
  }
  protected onAfterClose(): void {
    ReactDOM.unmountComponentAtNode(this.domElement);
    super.onAfterClose();
  }
}

function CustomDialog(props: {
  loading: boolean;
  textInside: JSX.Element;
  handleCloseClick: () => void;
}) {
  return (
    <div className={styles.maindialog}>
      <div id='overlay' className={styles.overlay}/>
      <div>
        <img
          src='data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAALQAAAC0CAMAAAAKE/YAAAAAUVBMVEXSXyj////HYzL/+/T/+Or/9d+yaUa9ZT2yaUj/9OG7Zj3SXybRYCj/+/b///3LYS/OYCvEZDS2aEL/89jAZTnMYS3/8dO7Zzusa02+ZTn/78wyF0DsAAABnUlEQVR4nO3ci26CMABGYQcoLRS5OTf2/g86R+KSLYUm2vxcPB8RTYzxkADRajkcAAAAAAAAAADYgbJcusCvqdtLnhfeJR/a96X7vOriarNJ/cUtHeiTnI7p26TsY+XRZ190sXSfVyA6X7rP6xZdzeweREeTGDt3IBIdTeCUR3Q0wQOxLNf3CWSr0ZvcPYiWIFqFaBWiVYhWIVqFaBWiVYhWIVqFaBWiVYhWIVqFaBWiVYhWIVqFaBWiVYhWIVqFaBWiVYhWIVqFaBWiVYhWIVqFaBWiVV4zeok/379m9BL2HO1Ckymlky0jRQc3Kqoou4f6YHzdaLX56PRzak757/JjfDS0dbOK6HM6Paf8P3st6lVE/9mAwPOpNcnqokOIJppoookmmmiiiSaaaKKJ3k30OfTFdU3RXZ+lT6qq6rbO+k4VXQ9fvT2OrH30Zo+3u/5rUI17NO3QmdPImIduxoyrUze0khEm5w6uqZNIRKNi91Hl5661dH+tdow6wts5J//BaJPRwH6IT1NxbDJ6vVc+nrXJaAAAAADALn0DBosqnCStFi4AAAAASUVORK5CYII='
          width='42'
          height='42'
        />
      </div>
      {props.loading && (
        <img
          style={{ marginLeft: '28%' }}
          src='/_layouts/15/images/progress.gif'
          id='imgid'
        />
      )}
      <div>
        <p className={styles.text} id='it'>
          {props.textInside}
        </p>
      </div>
      {!props.loading && (
        <div id='divid' className={styles.button}>
          <button
            id='divid1'
            className={styles.button1}
            onClick={props.handleCloseClick}
          >
            Close
          </button>
        </div>
      )}
    </div>
  );
}
