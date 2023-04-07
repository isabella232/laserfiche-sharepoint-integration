import * as React from 'react';
import { SPHttpClient, ISPHttpClientOptions } from '@microsoft/sp-http';
import { DeleteModal } from '../ProfileConfigurationComponents';
import { ChangeEvent, useState } from 'react';
import { IManageMappingsPageProps } from './IManageMappingsPageProps';
import { IListItem } from '../IListItem';
require('../../../../Assets/CSS/bootstrap.min.css');
require('../../../../Assets/CSS/adminConfig.css');
require('../../../../../node_modules/bootstrap/dist/js/bootstrap.min.js');

const sharepointValidationMapping = 'Please select a content type from the SharePoint Content Type drop down';
const laserficheValidationMapping = 'Please select a content type from the Laserfiche Profile dropdown';
const validationOf = 'Already Mapping exists for this SharePoint content type';

export default function ManageMappingsPage(props: IManageMappingsPageProps) {
  const [mappingRows, setMappingRows] = useState([]);
  const [sharePointContentTypes, setSharePointContentTypes] = useState([]);
  const [laserficheContentTypes, setLaserficheContentTypes] = useState([]);
  const [deleteModal, setDeleteModal] = useState(undefined);
  const [validationMessage, setValidationMessage] = useState(undefined);

  React.useEffect(() => {
    GetAllSharePointContentTypes();
    GetAllLaserficheContentTypes();
    GetItemIdByTitle().then((results: IListItem[]) => {
      if (results != null) {
        const jsonValue = JSON.parse(results[0].JsonValue);
        if (jsonValue.length > 0) {
          setMappingRows(mappingRows.concat(jsonValue));
        }
      }
    });
  }, [props.repoClient]);

  function GetAllLaserficheContentTypes() {
    const array = [];
    GetManageConfiguration().then((results: IListItem[]) => {
      if (results != null) {
        const jsonValue = JSON.parse(results[0].JsonValue);
        if (jsonValue.length > 0) {
          for (let i = 0; i < jsonValue.length; i++) {
            array.push({
              name: jsonValue[i].ConfigurationName,
            });
          }
          setLaserficheContentTypes(array);
        }
      }
    });
  }

  async function GetManageConfiguration(): Promise<IListItem[]> {
    const data: IListItem[] = [];
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
          data.push(results.value[i]);
        }
        return data;
      } else {
        return null;
      }
    } catch (error) {
      console.log('error occured' + error);
    }
  }

  async function GetAllSharePointContentTypes() {
    const array = [];
    const restApiUrl: string =
      props.context.pageContext.web.absoluteUrl + '/_api/web/contenttypes';
    try {
      const res = await fetch(restApiUrl, {
        method: 'GET',
        headers: {
          Accept: 'application/json',
          'Content-Type': 'application/json',
        },
      });
      const results = await res.json();
      for (let i = 0; i < results.value.length; i++) {
        array.push({
          name: results.value[i].Name,
        });
      }
      array.sort((a, b) => (a.name > b.name ? 1 : -1));
      setSharePointContentTypes(array);
    } catch (error) {
      console.log('error occured' + error);
    }
  }

  function CreateNewMapping(idx, rows) {
    setValidationMessage(undefined);
    if (rows[idx].SharePointContentType == 'Select') {
      setValidationMessage(sharepointValidationMapping);
    } else if (rows[idx].LaserficheContentType == 'Select') {
      setValidationMessage(laserficheValidationMapping);
    } else {
      GetItemIdByTitle().then((results: IListItem[]) => {
        if (results != null) {
          let entryExists = false;
          const itemId = results[0].Id;
          const jsonValue = JSON.parse(results[0].JsonValue);
          if (jsonValue.length > 0) {
            for (let i = 0; i < jsonValue.length; i++) {
              if (jsonValue[i].id == rows[idx].id) {
                entryExists = true;
                break;
              }
            }
            if (entryExists == true) {
              UpdateExistingMapping(jsonValue, rows, idx, itemId);
            } else {
              AddNewInExistingMapping(jsonValue, rows, idx, itemId);
            }
          } else {
            const restApiUrl: string =
              props.context.pageContext.web.absoluteUrl +
              "/_api/web/lists/getByTitle('AdminConfigurationList')/items(" +
              itemId +
              ')';
            const row = [...mappingRows];
            const newjsonValue = [
              {
                id: row[idx].id,
                SharePointContentType: row[idx].SharePointContentType,
                LaserficheContentType: row[idx].LaserficheContentType,
                toggle: true,
              },
            ];
            const jsonObject = JSON.stringify(newjsonValue);
            const body: string = JSON.stringify({
              Title: 'ManageMapping',
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
            rows[idx].toggle = !rows[idx].toggle;
            setMappingRows(rows);
          }
        } else {
          const restApiUrl: string =
            props.context.pageContext.web.absoluteUrl +
            "/_api/web/lists/getByTitle('AdminConfigurationList')/items";
          const newRow = [...mappingRows];
          const jsonValues = [
            {
              id: newRow[idx].id,
              SharePointContentType: newRow[idx].SharePointContentType,
              LaserficheContentType: newRow[idx].LaserficheContentType,
              toggle: true,
            },
          ];
          const jsonObject = JSON.stringify(jsonValues);
          const body: string = JSON.stringify({
            Title: 'ManageMapping',
            JsonValue: jsonObject,
          });
          const options: ISPHttpClientOptions = {
            headers: {
              Accept: 'application/json;odata=nometadata',
              'content-type': 'application/json;odata=nometadata',
              'odata-version': '',
            },
            body: body,
          };
          props.context.spHttpClient.post(
            restApiUrl,
            SPHttpClient.configurations.v1,
            options
          );
          rows[idx].toggle = !rows[idx].toggle;
          setMappingRows(rows);
        }
      });
    }
  }

  function AddNewInExistingMapping(jsonValue, rows, idx, itemId) {
    let exitEntry = false;
    for (let i = 0; i < jsonValue.length; i++) {
      if (
        jsonValue[i].SharePointContentType == rows[idx].SharePointContentType
      ) {
        exitEntry = true;
        break;
      }
    }
    if (exitEntry == false) {
      const restApiUrl: string =
        props.context.pageContext.web.absoluteUrl +
        "/_api/web/lists/getByTitle('AdminConfigurationList')/items(" +
        itemId +
        ')';
      const newJsonValue = [
        ...jsonValue,
        {
          id: rows[idx].id,
          SharePointContentType: rows[idx].SharePointContentType,
          LaserficheContentType: rows[idx].LaserficheContentType,
          toggle: true,
        },
      ];
      const jsonObject = JSON.stringify(newJsonValue);
      const body: string = JSON.stringify({
        Title: 'ManageMapping',
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
      rows[idx].toggle = !rows[idx].toggle;
      setMappingRows(rows);
      if (jsonValue.length + 1 == rows.length) {
        setValidationMessage(undefined);
      }
    } else {
      setValidationMessage(validationOf);
    }
  }

  function UpdateExistingMapping(jsonValue, rows, idx, itemId) {
    let exitEntry = false;
    for (let i = 0; i < jsonValue.length; i++) {
      if (
        jsonValue[i].SharePointContentType == rows[idx].SharePointContentType
      ) {
        if (jsonValue[i].id == rows[idx].id) {
          exitEntry = false;
          break;
        } else {
          exitEntry = true;
          break;
        }
      }
    }
    if (exitEntry == false) {
      for (let j = 0; j < jsonValue.length; j++) {
        if (jsonValue[j].id == rows[idx].id) {
          jsonValue[j].SharePointContentType = rows[idx].SharePointContentType;
          jsonValue[j].LaserficheContentType = rows[idx].LaserficheContentType;
          jsonValue[j].toggle = !rows[idx].toggle;
          break;
        }
      }
      const restApiUrl: string =
        props.context.pageContext.web.absoluteUrl +
        "/_api/web/lists/getByTitle('AdminConfigurationList')/items(" +
        itemId +
        ')';
      const newJsonValue = [...jsonValue];
      const jsonObject = JSON.stringify(newJsonValue);
      const body: string = JSON.stringify({
        Title: 'ManageMapping',
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
      rows[idx].toggle = !rows[idx].toggle;
      setMappingRows(rows);
      if (rows.some((item) => item.SharePointContentType == 'Select')) {
        setValidationMessage(sharepointValidationMapping);
      } else if (rows.some((item) => item.LaserficheContentType == 'Select')) {
        setValidationMessage(laserficheValidationMapping);
      } else {
        setValidationMessage(undefined);
      }
    } else {
      setValidationMessage(validationOf);
    }
  }

  function DeleteMapping(rows, idx) {
    GetItemIdByTitle().then((results: IListItem[]) => {
      if (results != null) {
        const itemId = results[0].Id;
        const jsonValue = JSON.parse(results[0].JsonValue);
        let entryExists = -1;
        for (let i = 0; i < jsonValue.length; i++) {
          if (jsonValue[i].id == rows[idx].id) {
            entryExists = i;
            break;
          }
        }
        if (entryExists > -1) {
          jsonValue.splice(entryExists, 1);
          const restApiUrl: string =
            props.context.pageContext.web.absoluteUrl +
            "/_api/web/lists/getByTitle('AdminConfigurationList')/items(" +
            itemId +
            ')';
          const newJsonValue = [...jsonValue];
          const jsonObject = JSON.stringify(newJsonValue);
          const body: string = JSON.stringify({
            Title: 'ManageMapping',
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
          for (let q = jsonValue.length + 1; q < rows.length; q++) {
            if (
              rows[idx].SharePointContentType == rows[q].SharePointContentType
            ) {
              setValidationMessage(undefined);
              break;
            } else {
              setValidationMessage(validationOf);
            }
          }
        } else {
          if (jsonValue.length + 1 == rows.length) {
            setValidationMessage(undefined);
          } else {
            for (let j = jsonValue.length; j < rows.length; j++) {
              if (rows[j].SharePointContentType == 'Select') {
                setValidationMessage(sharepointValidationMapping);
                break;
              } else if (rows[j].LaserficheContentType == 'Select') {
                setValidationMessage(laserficheValidationMapping);
              }
            }
          }
        }
      }
    });
  }

  async function GetItemIdByTitle(): Promise<IListItem[]> {
    const array: IListItem[] = [];
    const restApiUrl: string =
      props.context.pageContext.web.absoluteUrl +
      "/_api/web/lists/getByTitle('AdminConfigurationList')/Items?$select=Id,Title,JsonValue&$filter=Title eq 'ManageMapping'";
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

  const AddNewMapping = () => {
    const id = (+new Date() + Math.floor(Math.random() * 999999)).toString(36);
    const item = {
      id: id,
      SharePointContentType: 'Select',
      LaserficheContentType: 'Select',
      toggle: false,
    };
    setMappingRows([...mappingRows, item]);
  };

  const RemoveSpecificMapping = (idx) => {
    const rows = [...mappingRows];
    const delModal = (
      <DeleteModal
        onCancel={CloseModalUp}
        onConfirmDelete={() => RemoveRow(idx)}
        configurationName={rows[idx].SharePointContentType}
      ></DeleteModal>
    );
    setDeleteModal(delModal);
  };

  function RemoveRow(id: number) {
    const rows = [...mappingRows];
    const deleteRows = [...mappingRows];
    rows.splice(id, 1);
    setMappingRows(rows);
    DeleteMapping(deleteRows, id);
    setDeleteModal(undefined);
  }

  const EditSpecificMapping = (idx) => {
    const rows = [...mappingRows];
    rows[idx].toggle = !rows[idx].toggle;
    setMappingRows(rows);
  };

  const SaveSpecificMapping = (idx) => {
    const rows = [...mappingRows];
    CreateNewMapping(idx, rows);
  };

  const handleChange = (event: ChangeEvent<HTMLSelectElement>, idx: number) => {
    const item = {
      id: event.target.id,
      name: event.target.name,
      value: event.target.value,
    };
    const newRows = [...mappingRows];
    if (item.name == 'SharePointContentType') {
      newRows[idx].SharePointContentType = item.value;
    } else if (item.name == 'LaserficheContentType') {
      newRows[idx].LaserficheContentType = item.value;
    }
    setMappingRows(newRows);
  };

  function CloseModalUp() {
    setDeleteModal(undefined);
  }

  const Reset = () => {
    setDeleteModal(undefined);
    GetAllSharePointContentTypes();
    GetAllLaserficheContentTypes();
    GetItemIdByTitle().then((results: IListItem[]) => {
      if (results != null) {
        const jsonValue = JSON.parse(results[0].JsonValue);
        if (jsonValue.length > 0) {
          setMappingRows(jsonValue);
        }
      }
    });
    setValidationMessage(undefined);
  };

  const SharePointContents = sharePointContentTypes.map((v) => (
    <option value={v.name}>{v.name}</option>
  ));
  const LaserficheContents = laserficheContentTypes.map((v) => (
    <option value={v.name}>{v.name}</option>
  ));
  const renderTableData = mappingRows.map((item, index) => {
    if (item.toggle) {
      return (
        <tr id='addr0' key={index}>
          <td>
            <select
              name='SharePointContentType'
              disabled
              className='custom-select'
              value={mappingRows[index].SharePointContentType}
              id={mappingRows[index].id}
              onChange={(e) => handleChange(e, index)}
            >
              <option>Select</option>
              {SharePointContents}
            </select>
          </td>
          <td>
            <select
              name='LaserficheContentType'
              disabled
              className='custom-select'
              value={mappingRows[index].LaserficheContentType}
              id={mappingRows[index].id}
              onChange={(e) => handleChange(e,index)}
            >
              <option>Select</option>
              {LaserficheContents}
            </select>
          </td>
          <td className='text-center'>
            <a
              href='javascript:;'
              className='ml-3'
              onClick={() => EditSpecificMapping(index)}
            >
              <span className='material-icons'>edit</span>
            </a>
            <a
              href='javascript:;'
              className='ml-3'
              onClick={() => RemoveSpecificMapping(index)}
            >
              <span className='material-icons'>delete</span>
            </a>
          </td>
        </tr>
      );
    } else {
      return (
        <tr id='addr0' key={index}>
          <td>
            <select
              name='SharePointContentType'
              className='custom-select'
              value={mappingRows[index].SharePointContentType}
              id={mappingRows[index].id}
              onChange={(e) => handleChange(e, index)}
            >
              <option>Select</option>
              {SharePointContents}
            </select>
          </td>
          <td>
            <select
              name='LaserficheContentType'
              className='custom-select'
              value={mappingRows[index].LaserficheContentType}
              id={mappingRows[index].id}
              onChange={(e) => handleChange(e, index)}
            >
              <option>Select</option>
              {LaserficheContents}
            </select>
          </td>
          <td className='text-center'>
            <a
              href='javascript:;'
              className='ml-3'
              onClick={() => SaveSpecificMapping(index)}
            >
              <span className='material-icons'>save</span>
            </a>
            <a
              href='javascript:;'
              className='ml-3'
              onClick={() => RemoveSpecificMapping(index)}
            >
              <span className='material-icons'>delete</span>
            </a>
          </td>
        </tr>
      );
    }
  });
  const viewSharePointContentTypes =
    props.context.pageContext.web.absoluteUrl + '/_layouts/15/mngctype.aspx';

  return (
    <>
      <div
        className='container-fluid p-3'
        style={{ maxWidth: '85%', marginLeft: '-26px' }}
      >
        <div className='p-3'>
          <div className='card rounded-0'>
            <div className='card-header d-flex justify-content-between'>
              <div>
                <h6 className='mb-0'>Content Type Mappings Laserfiche</h6>
              </div>
              <div>
                <a
                  href=''
                  onClick={() => window.open(viewSharePointContentTypes)}
                  target='_blank'
                >
                  View SharePoint Content Types
                </a>
              </div>
            </div>
            <div className='card-body'>
              <table className='table table-sm'>
                <thead>
                  <tr>
                    <th className='text-center' style={{ width: '45%' }}>
                      SharePoint Content Type
                    </th>
                    <th className='text-center' style={{ width: '45%' }}>
                      Laserfiche Profile
                    </th>
                    <th className='text-center'>Action</th>
                  </tr>
                </thead>
                <tbody>{renderTableData}</tbody>
              </table>
            </div>

            {validationMessage && <div id='sharePointValidationMapping' style={{ color: 'red' }}>
              <span>
                {validationMessage}
              </span>
            </div>}
            <div id='laserficheValidationMapping' style={{ color: 'red' }}>
              <span>
                
              </span>
            </div>
            <div id='validationOfMapping' style={{ color: 'red' }}>
              <span>
                
              </span>
            </div>
            <div className='card-footer bg-transparent'>
              <a
                className='btn btn-primary pl-5 pr-5 float-right'
                style={{ marginLeft: '10px' }}
                onClick={Reset}
              >
                Reset
              </a>
              <a
                href='javascript:;'
                className='btn btn-primary pl-5 pr-5 float-right'
                onClick={AddNewMapping}
              >
                Add
              </a>
            </div>
          </div>
        </div>
      </div>
      <div
        className='modal'
        id='deleteModal'
        hidden={!deleteModal}
        data-backdrop='static'
        data-keyboard='false'
      >
        {deleteModal}
      </div>
    </>
  );
}
