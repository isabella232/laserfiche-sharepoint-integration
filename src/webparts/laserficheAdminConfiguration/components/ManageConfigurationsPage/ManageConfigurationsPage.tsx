import * as React from 'react';
import { useEffect, useState } from 'react';
import { NavLink } from 'react-router-dom';
import { IManageConfigurationPageProps } from './IManageConfigurationPageProps';
import { IListItem } from './IListItem';
import { SPHttpClient, ISPHttpClientOptions } from '@microsoft/sp-http';
require('../../../../Assets/CSS/bootstrap.min.css');
require('../../../../Assets/CSS/adminConfig.css');
require('../../../../../node_modules/bootstrap/dist/js/bootstrap.min.js');

export default function ManageConfigurationsPage(
  props: IManageConfigurationPageProps
) {
  const [configRows, setConfigRows] = useState([]);
  const [deleteModal, setDeleteModal] = useState<JSX.Element | undefined>(
    undefined
  );

  useEffect(() => {
    GetItemIdByTitle().then((results: IListItem[]) => {
      if (results != null) {
        const jsonValue = JSON.parse(results[0].JsonValue);
        if (jsonValue.length > 0) {
          setConfigRows(configRows.concat(...jsonValue));
        }
      }
    });
  }, []);

  async function GetItemIdByTitle(): Promise<IListItem[]> {
    const array: IListItem[] = [];
    const restApiUrl: string =
      props.context.pageContext.web.absoluteUrl +
      "/_api/web/lists/getByTitle('AdminConfigurationList')/Items?$select=Id,Title,JsonValue&$filter=Title eq 'ManageConfigurations'";
    try {
      const res = await fetch(restApiUrl, {
        method: 'GET',
        headers: {
          Accept: 'application/json',
          'Content-Type': 'application/json',
        },
      });
      const results = await res.json();
      if (results.value.length > 0) {
        for (let i = 0; i < results.value.length; i++) {
          array.push(results.value[i]);
        }
        return array;
      } else {
        return null;
      }
    } catch (error) {
      console.log('error occured' + error);
    }
  }

  function RemoveSpecificConfiguration(idx: number) {
    const rows = [...configRows];
    const configName = rows[idx].ConfigurationName;
    const deleteModal = (
      <DeleteModal
        configurationName={configName}
        onConfirmDelete={() => RemoveRow(idx)}
        onCancel={CloseModalUp}
      ></DeleteModal>
    );
    setDeleteModal(deleteModal);
  }

  function RemoveRow(id: number) {
    const rows = [...configRows];
    const deleteRows = [...configRows];
    rows.splice(id, 1);
    DeleteMapping(deleteRows, id);
    setDeleteModal(undefined);
  }

  function CloseModalUp() {
    setDeleteModal(undefined);
  }

  //Delete the selected configuration from the list
  function DeleteMapping(rows, idx) {
    GetItemIdByTitle().then((results: IListItem[]) => {
      if (results != null) {
        const itemId = results[0].Id;
        const profileConfigurations = JSON.parse(results[0].JsonValue);
        const profileToRemove = rows[idx].ConfigurationName;
        for (let i = 0; i < profileConfigurations.length; i++) {
          if (profileConfigurations[i].ConfigurationName == profileToRemove) {
            profileConfigurations.splice(i, 1);
            setConfigRows(profileConfigurations);
            const restApiUrl: string =
              props.context.pageContext.web.absoluteUrl +
              "/_api/web/lists/getByTitle('AdminConfigurationList')/items(" +
              itemId +
              ')';
            const jsonObject = JSON.stringify(profileConfigurations);
            const body: string = JSON.stringify({
              Title: 'ManageConfigurations',
              JsonValue: jsonObject,
            });
            const options: ISPHttpClientOptions = {
              headers: {
                Accept: 'application/json;odata=nometadata',
                'content-type': 'application/json;odata=nometadata',
                'odata-version': '',
                'IF-MATCH': '*',
                'X-HTTP-Method': 'MERGE',
              },
              body: body,
            };
            props.context.spHttpClient.post(
              restApiUrl,
              SPHttpClient.configurations.v1,
              options
            );
            break;
          }
        }
      }
    });
  }

  //Dynamically render list of configurations created in the table format
  const tableData = configRows.map((item, index) => {
    return (
      <tr id='addr0' key={index}>
        <td>{item.ConfigurationName}</td>
        <td className='text-center'>
          <span>
            <NavLink
              to={
                '/EditManageConfiguration/' +
                item.ConfigurationName
              }
              style={{
                marginRight: '18px',
                fontWeight: '500',
                fontSize: '15px',
              }}
            >
              <span className='material-icons'>edit</span>
            </NavLink>
          </span>
          <a
            href='javascript:;'
            className='ml-3'
            onClick={() => RemoveSpecificConfiguration(index)}
          >
            <span className='material-icons'>delete</span>
          </a>
        </td>
      </tr>
    );
  });

  return (
    <div>
      <div
        className='container-fluid p-3'
        style={{ maxWidth: '85%', marginLeft: '-26px' }}
      >
        <main className='bg-white shadow-sm'>
          <div className='p-3'>
            <div className='card rounded-0'>
              <div className='card-header d-flex justify-content-between pt-1 pb-1'>
                <NavLink
                  to='/AddNewManageConfiguration'
                  style={{
                    marginRight: '18px',
                    fontWeight: '500',
                    fontSize: '15px',
                  }}
                >
                  <a className='btn btn-primary pl-5 pr-5'>Add Profile</a>
                </NavLink>
              </div>
              <div className='card-body'>
                <table className='table table-bordered table-striped table-hover'>
                  <thead>
                    <tr>
                      <th className='text-center'>Profile Name</th>
                      <th className='text-center' style={{ width: '30%' }}>
                        Action
                      </th>
                    </tr>
                  </thead>
                  <tbody>{tableData}</tbody>
                </table>
              </div>
            </div>
          </div>
        </main>
      </div>
      <div>
        <div
          className='modal'
          id='deleteModal'
          hidden={!deleteModal}
          data-backdrop='static'
          data-keyboard='false'
        >
          {deleteModal}
        </div>
      </div>
    </div>
  );
}

function DeleteModal(props: {
  configurationName: string;
  onConfirmDelete: () => void;
  onCancel: () => void;
}) {
  return (
    <div className='modal-dialog modal-dialog-centered'>
      <div className='modal-content'>
        <div className='modal-header'>
          <h5 className='modal-title' id='ModalLabel'>
            Delete Confirmation
          </h5>
          <button
            type='button'
            className='close'
            data-dismiss='modal'
            aria-label='Close'
            onClick={props.onCancel}
          >
            <span aria-hidden='true'>&times;</span>
          </button>
        </div>
        <div className='modal-body'>
          Do you want to permanently delete &quot;
          {props.configurationName}&quot;?
        </div>
        <div className='modal-footer'>
          <button
            type='button'
            className='btn btn-primary btn-sm'
            data-dismiss='modal'
            onClick={props.onConfirmDelete}
          >
            OK
          </button>
          <button
            type='button'
            className='btn btn-secondary btn-sm'
            data-dismiss='modal'
            onClick={props.onCancel}
          >
            Cancel
          </button>
        </div>
      </div>
    </div>
  );
}
