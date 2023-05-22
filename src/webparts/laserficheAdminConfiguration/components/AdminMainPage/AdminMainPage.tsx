import * as React from 'react';
import { NavLink } from 'react-router-dom';
import { useEffect } from 'react';
import { IAdminPageProps } from './IAdminPageProps';
import { CreateConfigurations } from '../../../../Utils/CreateConfigurations';
require('../../../../Assets/CSS/bootstrap.min.css');
require('../../adminConfig.css');

declare global {
  // eslint-disable-next-line
  namespace JSX {
    interface IntrinsicElements {
      // eslint-disable-next-line
      ['lf-login']: any;
    }
  }
}

export default function AdminMainPage(props: IAdminPageProps) {
  useEffect(() => {
    CreateConfigurations.CreateAdminConfigList(props.context);
    CreateConfigurations.CreateDocumentConfigList(props.context);
  }, []);

  const linkData: LinkInfo[] = [
    { route: '/HomePage', name: 'About' },
    { route: '/ManageConfigurationsPage', name: 'Profiles' },
    { route: '/ManageMappingsPage', name: 'Profile Mapping' },
  ];

  return (
    <div style={{ borderBottom: '3px solid #CE7A14', width: '80%' }}>
      <div>
        <span
          style={{
            marginRight: '450px',
            fontSize: '18px',
            fontWeight: '500',
          }}
        >
          Profile Editor
        </span>
        {props.loggedIn && <Links linkData={linkData} />}
      </div>
    </div>
  );
}

interface LinkInfo {
  route: string;
  name: string;
}

function Links(props: { linkData: LinkInfo[] }) {
  const linkEls = props.linkData.map((link: LinkInfo) => (
    <span key={link.name}>
      <NavLink
        to={link.route}
        activeStyle={{ fontWeight: 'bold', color: 'red' }}
        style={{
          marginRight: '25px',
          fontWeight: '500',
          fontSize: '15px',
        }}
      >
        {link.name}
      </NavLink>
    </span>
  ));
  return <div>{linkEls}</div>;
}

