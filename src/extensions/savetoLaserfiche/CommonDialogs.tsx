import * as React from 'react';
import { Navigation } from 'spfx-navigation';
import styles from './SendToLaserFiche.module.scss';
import { SPComponentLoader } from '@microsoft/sp-loader';
import { SavedToLaserficheDocumentData } from './SaveDocumentToLaserfiche';

const SAVING_DOCUMENT_TO_LASERFICHE = 'Saving your document to Laserfiche';

export default function LoadingDialog() {
  return (
    <>
      <img src='/_layouts/15/images/progress.gif' />
      <br />
      <div>{SAVING_DOCUMENT_TO_LASERFICHE}</div>
    </>
  );
}

const DOCUMENT_UPLOADED =
  'Document successfully uploaded to Laserfiche with name:';
const DOCUMENT_UPLOADED_METADATA_FAILED =
  'All metadata failed to save due to an invalid field';
const TEMPLATE_FIELDS_NOT_APPLIED =
  'The Laserfiche template and fields were not applied to this document.';
const CLOSE = 'Close';
const GO_TO_FILE = 'View file in Laserfiche';
const GO_TO_LIBRARY = 'Return to SharePoint library';
const CLICK_HERE_VIEW_FILE_LASERFICHE = 'View the file in Laserfiche';
const CLICK_HERE_GO_SHAREPOINT_LIBRARY = 'Return to your SharePoint library';

export function SavedToLaserficheSuccessDialogText(props: {
  successfulSave: SavedToLaserficheDocumentData;
}) {
  React.useEffect(() => {
    SPComponentLoader.loadCss(
      'https://cdn.jsdelivr.net/npm/@laserfiche/lf-ui-components@14/cdn/indigo-pink.css'
    );
    SPComponentLoader.loadCss(
      'https://cdn.jsdelivr.net/npm/@laserfiche/lf-ui-components@14/cdn/lf-ms-office-lite.css'
    );
  });
  const metadataFailedNotice: JSX.Element = (
    <span>
      {`${DOCUMENT_UPLOADED} ${props.successfulSave.fileName}. ${DOCUMENT_UPLOADED_METADATA_FAILED}`}
      <br /> <p style={{ color: 'red' }}>{TEMPLATE_FIELDS_NOT_APPLIED}</p>
    </span>
  );

  return (
    <>
      <div>
        <p>
          {props.successfulSave.metadataSaved
            ? `${DOCUMENT_UPLOADED}  ${props.successfulSave.fileName}.`
            : metadataFailedNotice}
        </p>
      </div>
    </>
  );
}

export function SavedToLaserficheSuccessDialogButtons(props: {
  closeClick: () => Promise<void>;
  successfulSave: SavedToLaserficheDocumentData;
  hadToRouteToLogin: boolean;
}) {
  React.useEffect(() => {
    SPComponentLoader.loadCss(
      'https://cdn.jsdelivr.net/npm/@laserfiche/lf-ui-components@14/cdn/indigo-pink.css'
    );
    SPComponentLoader.loadCss(
      'https://cdn.jsdelivr.net/npm/@laserfiche/lf-ui-components@14/cdn/lf-ms-office-lite.css'
    );
  });

  function viewFile() {
    window.open(props.successfulSave.fileLink);
  }

  function redirect() {
    Navigation.navigate(props.successfulSave.pathBack, true);
  }

  return (
    <>
      {props.successfulSave.fileLink && (
        <button
          className={`lf-button primary-button ${styles.actionButton}`}
          title={CLICK_HERE_VIEW_FILE_LASERFICHE}
          onClick={viewFile}
        >
          {GO_TO_FILE}
        </button>
      )}
      {props.hadToRouteToLogin && (
        <button
          className={`lf-button primary-button ${styles.actionButton}`}
          title={CLICK_HERE_GO_SHAREPOINT_LIBRARY}
          onClick={redirect}
        >
          {GO_TO_LIBRARY}
        </button>
      )}
      <button className='lf-button sec-button' onClick={props.closeClick}>
        {CLOSE}
      </button>
    </>
  );
}
