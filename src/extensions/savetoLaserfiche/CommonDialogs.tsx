import * as React from 'react';
import { Navigation } from 'spfx-navigation';
import styles from './SendToLaserFiche.module.scss';

const SAVING_DOCUMENT_TO_LASERFICHE = 'Saving your document to Laserfiche';

export default function LoadingDialog() {
  return (
    <>
      <img
        style={{ marginLeft: '28%' }}
        src='/_layouts/15/images/progress.gif'
        id='imgid'
      />
      <div>
        <p className={styles.text}>{SAVING_DOCUMENT_TO_LASERFICHE}</p>
      </div>
    </>
  );
}

const DOCUMENT_UPLOADED = 'Document uploaded';
const DOCUMENT_UPLOADED_METADATA_FAILED =
  'Document uploaded to repository, updating metadata failed due to constraint mismatch';
const TEMPLATE_FIELDS_NOT_APPLIED =
  'The Laserfiche template and fields were not applied to this document.';
const CLOSE = 'Close';
const GO_TO_FILE = 'Go to File';
const GO_TO_LIBRARY = 'Go to Library';
const CLICK_HERE_VIEW_FILE_LASERFICHE =
  'Click here to view the file in Laserfiche';
const CLICK_HERE_GO_SHAREPOINT_LIBRARY =
  'Click here to go back to your SharePoint library';

export function SavedToLaserficheSuccessDialog(props: {
  closeClick: (success: boolean) => Promise<void>;
  successfulSave: {
    fileLink: string;
    pathBack: string;
    metadataSaved: boolean;
  };
}) {
  const metadataFailedNotice: JSX.Element = (
    <span>
      {DOCUMENT_UPLOADED_METADATA_FAILED}
      <br /> <p style={{ color: 'red' }}>{TEMPLATE_FIELDS_NOT_APPLIED}</p>
    </span>
  );

  function viewFile() {
    window.open(props.successfulSave.fileLink);
  }

  function redirect() {
    Navigation.navigate(props.successfulSave.pathBack, true);
  }

  return (
    <>
      <div>
        <p className={styles.text}>
          {props.successfulSave.metadataSaved
            ? DOCUMENT_UPLOADED
            : metadataFailedNotice}
        </p>
      </div>

      <div className={styles.button}>
        <button className={styles.button1} onClick={() => props.closeClick(true)}>
          {CLOSE}
        </button>
        {props.successfulSave.fileLink && (
          <button
            className={styles.button2}
            title={CLICK_HERE_VIEW_FILE_LASERFICHE}
            onClick={viewFile}
          >
            {GO_TO_FILE}
          </button>
        )}
        <button
          className={styles.button2}
          title={CLICK_HERE_GO_SHAREPOINT_LIBRARY}
          onClick={redirect}
        >
          {GO_TO_LIBRARY}
        </button>
      </div>
    </>
  );
}
