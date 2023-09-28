import * as React from 'react';
require('./../../../../Assets/CSS/commonStyles.css');
require('../../../../Assets/CSS/bootstrap.min.css');

export default function HomePage(): JSX.Element {
  return (
    <div className='p-3'>
      <main className='bg-white shadow-sm'>
        <p className='adminContent'>
          Use the Laserfiche Administration page to edit your SharePoint and
          Laserfiche configuration. Sign in and select the task you want to
          perform from the menu at the top of this section.
        </p>
        <p className='adminContent'>
          For more information, see the{' '}
          <a
            href='https://laserfiche.github.io/laserfiche-sharepoint-integration/'
            target='_blank'
            rel='noreferrer'
            style={{color: '#0079d6'}}
          >
            help documentation.
          </a>{' '}
          <i>Note: the help link is not live yet.</i>{' '}
        </p>
        <div className='adminContent'>
          <p>
            <strong>Profiles</strong>
          </p>
          <p style={{ marginLeft: '38px' }}>
            <span>
              Profiles govern how documents in SharePoint will be saved to
              Laserfiche. You can create multiple profiles for different
              SharePoint content types. For example, if you want applications
              stored differently than invoices, create separate
              profiles for each content type.
            </span>
          </p>
          <p>
            <strong>Profile Mapping</strong>
          </p>
          <p style={{ marginLeft: '38px' }}>
            <span>
              In this tab, you can map a specific SharePoint content type with a
              corresponding Laserfiche profile. This profile will then be used
              when saving all documents of the specified SharePoint content
              type.
            </span>
          </p>
        </div>
      </main>
    </div>
  );
}
