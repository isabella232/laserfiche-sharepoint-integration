import * as React from 'react';
import styles from './SendToLaserFiche.module.scss';
import { SPComponentLoader } from '@microsoft/sp-loader';
import { SavedToLaserficheDocumentData } from './SaveDocumentToLaserfiche';

const SAVING_DOCUMENT_TO_LASERFICHE = 'Saving document to Laserfiche...';

export default function LoadingDialog(): JSX.Element {
  return (
    <>
      <img src='/_layouts/15/images/progress.gif' />
      <br />
      <div>{SAVING_DOCUMENT_TO_LASERFICHE}</div>
    </>
  );
}

const DOCUMENT_SUCCESSFULLY_UPLOADED_TO_LASERFICHE_WITH_NAME =
  'Document successfully uploaded to Laserfiche with name:';
const METADATA_FAILED_TO_SAVE_INVALID_FIELD =
  'All metadata failed to save due to at least one invalid field.';
const CLOSE = 'Close';
const VIEW_FILE_IN_LASERFICHE = 'View file in Laserfiche';

export function SavedToLaserficheSuccessDialogText(props: {
  successfulSave: SavedToLaserficheDocumentData;
}): JSX.Element {
  React.useEffect(() => {
    SPComponentLoader.loadCss(
      'https://cdn.jsdelivr.net/npm/@laserfiche/lf-ui-components@14/cdn/indigo-pink.css'
    );
    SPComponentLoader.loadCss(
      'https://cdn.jsdelivr.net/npm/@laserfiche/lf-ui-components@14/cdn/lf-ms-office-lite.css'
    );
  }, []);

  const metadataFailedNotice: JSX.Element = (
    <>
      <div className={styles.paddingUnder}>
        <b>Warning: </b>
        {METADATA_FAILED_TO_SAVE_INVALID_FIELD}
      </div>
      <Collapsible title='Error details'>
        {props.successfulSave.failedMetadata}
      </Collapsible>
    </>
  );

  return (
    <>
      <div className={styles.successSaveToLaserfiche}>
        <div className={styles.paddingUnder}>
          {`${DOCUMENT_SUCCESSFULLY_UPLOADED_TO_LASERFICHE_WITH_NAME}  ${props.successfulSave.fileName}.`}
        </div>
        {!props.successfulSave.metadataSaved && metadataFailedNotice}
      </div>
    </>
  );
}

export function Collapsible(props: {
  open?: boolean;
  children: JSX.Element;
  title: string;
}): JSX.Element {
  const [isOpen, setIsOpen] = React.useState<boolean>(props.open ?? false);

  const handleFilterOpening: () => void = () => {
    setIsOpen((prev) => !prev);
  };

  return (
    <>
      <div className={styles.collapseBox}>
        <button
          className={styles.lfMaterialIconButton}
          onClick={handleFilterOpening}
        >
          {!isOpen ? (
            <span className='material-icons-outlined'> chevron_right </span>
          ) : (
            <span className='material-icons-outlined'> expand_less </span>
          )}
        </button>
        <span>{props.title}</span>
      </div>

      {isOpen && props.children}
    </>
  );
}

export function SavedToLaserficheSuccessDialogButtons(props: {
  closeClick: () => Promise<void>;
  successfulSave: SavedToLaserficheDocumentData;
}): JSX.Element {
  React.useEffect(() => {
    SPComponentLoader.loadCss(
      'https://cdn.jsdelivr.net/npm/@laserfiche/lf-ui-components@14/cdn/indigo-pink.css'
    );
    SPComponentLoader.loadCss(
      'https://cdn.jsdelivr.net/npm/@laserfiche/lf-ui-components@14/cdn/lf-ms-office-lite.css'
    );
  }, []);

  function viewFile(): void {
    window.open(props.successfulSave.fileLink);
  }

  return (
    <>
      {props.successfulSave?.fileLink && (
        <button
          className={`lf-button primary-button ${styles.actionButton}`}
          title={VIEW_FILE_IN_LASERFICHE}
          onClick={viewFile}
        >
          {VIEW_FILE_IN_LASERFICHE}
        </button>
      )}
      <button className='lf-button sec-button' onClick={props.closeClick}>
        {CLOSE}
      </button>
    </>
  );
}
