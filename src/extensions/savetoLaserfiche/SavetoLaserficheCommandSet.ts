import { Log } from '@microsoft/sp-core-library';
import CustomDailog from './CustomDailog';
import {
  BaseListViewCommandSet,
  Command,
  IListViewCommandSetListViewUpdatedParameters,
  IListViewCommandSetExecuteEventParameters,
} from '@microsoft/sp-listview-extensibility';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { SPComponentLoader } from '@microsoft/sp-loader';
import * as React from 'react';
import { Navigation } from 'spfx-navigation';

SPComponentLoader.loadScript(
  'https://cdnjs.cloudflare.com/ajax/libs/jquery/3.1.1/jquery.min.js',
  {
    globalExportsName: 'jQuery',
  }
);

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface ISendToLfCommandSetProperties {
  // This is an example; replace with your own properties
  sampleTextOne: string;
  sampleTextTwo: string;
}

const LOG_SOURCE = 'SendToLfCommandSet';
const dialog: CustomDailog = new CustomDailog();
let staticFieldNames;
let fieldDataStatic = [];
let fieldDataStaticAll = [];
let fieldDataInternal = [];
let fieldDataDisplay = [];
let allFieldValueStore;
let webpartconfigurations = '';
let webpartconfigurationsAdmin = '';
const Redirectpagelink = '/SitePages/LaserficheSpSignIn.aspx';
export default class SendToLfCommandSet extends BaseListViewCommandSet<ISendToLfCommandSetProperties> {
  public fieldContainer: React.RefObject<any>;
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, 'Initialized SendToLfCommandSet');
    this.fieldContainer = React.createRef();
    return Promise.resolve();
  }

  public onListViewUpdated(
    event: IListViewCommandSetListViewUpdatedParameters
  ): void {
    const compareOneCommand: Command = this.tryGetCommand('COMMAND_1');
    if (compareOneCommand) {
      // This command should be hidden unless exactly one row is selected.
      compareOneCommand.visible =
        event.selectedRows.length === 1 &&
        event.selectedRows[0]['_values'].get('ContentType') !== 'Folder';
    }
  }
  public async onExecute(
    event: IListViewCommandSetExecuteEventParameters
  ): Promise<void> {
    const libraryUrl = this.context.pageContext.list.title;
    const allfieldsvalues = event.selectedRows[0]['_values'];
    const fileId = allfieldsvalues.get('ID');
    const fileSize = allfieldsvalues.get('File_x0020_Size');
    const fileUrl = event.selectedRows[0]['_values'].get('FileRef');
    const fileName = event.selectedRows[0]['_values'].get('FileLeafRef');
    await this.GetAllFieldsProperties(libraryUrl);
    await this.GetAllFieldsValues(libraryUrl, fileId);
    await this.pageConfigurationCheck();
    const filecontenttypename =
      event.selectedRows[0]['_values'].get('ContentType');
    const fileNamelength = fileName.split('.').length;
    let fileSplitValue = '';
    const fileExtensionOnly = fileName.split('.')[fileNamelength - 1];
    for (let j = 0; j < fileNamelength - 1; j++) {
      fileSplitValue += fileName.split('.')[j] + '.';
    }
    const fileNoName = fileSplitValue.slice(0, -1);
    const siteurl = window.location.origin;
    const requestUrl = siteurl + fileUrl;
    const isCheckedOut = allfieldsvalues.get('CheckoutUser');
    if (filecontenttypename === 'Folder') {
      alert('Cannot Send a Folder To Laserfiche');
    } else if (fileNoName === '') {
      alert(
        'Please add a filename to the selected file before trying to save to Laserfiche.'
      );
    } else if (fileExtensionOnly === 'url') {
      alert('Cannot send the .url file to Laserfiche');
    } else if (isCheckedOut != '') {
      alert(
        'The selected file is checked out. Please discard the checkout or check the file back in before trying to save to Laserfiche.'
      );
    } else if (fileSize > 100000000) {
      alert('Please select a file below 100MB size');
    } else if (webpartconfigurations != 'True') {
      alert(
        'Missing "LaserficheSpSignIn" SharePoint page. Please refer to the admin guide and complete configuration steps exactly as described.'
      );
    } else if (webpartconfigurationsAdmin != 'True') {
      alert(
        'Missing "LaserficheSpAdministration" SharePoint page. Please refer to the admin guide and complete configuration steps exactly as described.'
      );
    } else {
      dialog.show();
      this.getAdminData(
        fileName,
        filecontenttypename,
        fileUrl,
        requestUrl,
        fileExtensionOnly,
        siteurl
      );
    }
  }
  //checking whether the Sign-in Page configured or not
  public async pageConfigurationCheck(): Promise<any> {
    webpartconfigurations = '';
    try {
      const res = await fetch(
        this.context.pageContext.web.absoluteUrl +
          `/_api/web/lists/getbytitle('Site Pages')/items`,
        {
          method: 'GET',
          headers: {
            Accept: 'application/json',
            'Content-Type': 'application/json',
          },
        }
      );
      const resultsrr = await res.json();
      console.log(resultsrr);
      for (let o = 0; o < resultsrr.value.length; o++) {
        const pageName = resultsrr['value'][o]['Title'];
        if (pageName === 'LaserficheSpSignIn') {
          webpartconfigurations = 'True';
        } else if (pageName === 'LaserficheSpAdministration') {
          webpartconfigurationsAdmin = 'True';
        }
      }
    } catch (error) {
      // TODO
    }
  }
  // getting All Fields from the library and other properties
  public async GetAllFieldsProperties(libraryUrl): Promise<any> {
    let dataStatic: object = {};
    let dataDisplay: object = {};
    let dataInternal: object = {};
    try {
      const res = await fetch(
        this.context.pageContext.web.absoluteUrl +
          `/_api/web/lists/getbytitle('${libraryUrl}')/Fields?$filter=Group ne '_Hidden'`,
        {
          method: 'GET',
          headers: {
            Accept: 'application/json',
            'Content-Type': 'application/json',
          },
        }
      );
      const results = await res.json();
      for (let i = 0; i < results.value.length; i++) {
        const fieldStaticName = results.value[i]['StaticName'];
        const fieldDisplayName = results.value[i]['Title'];
        const fieldInternalName = results.value[i]['InternalName'];
        dataStatic = { [fieldStaticName]: fieldInternalName };
        dataDisplay = { [fieldStaticName]: fieldDisplayName };
        dataInternal = { [fieldInternalName]: fieldDisplayName };
        fieldDataDisplay.push(dataDisplay);
        fieldDataStaticAll.push(dataStatic);
        //if(uniqueArray.includes(fieldInternalName)){
        staticFieldNames.push(fieldStaticName);
        fieldDataStatic.push(dataStatic);
        fieldDataInternal.push(dataInternal);
        //}
      }
      console.log(staticFieldNames);
      return results;
    } catch (error) {
      console.log('error occured' + error);
    }
  }

  //getting all the Fields Values for the Selected file
  public async GetAllFieldsValues(libraryUrl, fileId): Promise<any> {
    try {
      const res = await fetch(
        this.context.pageContext.web.absoluteUrl +
          `/_api/web/lists/getbytitle('${libraryUrl}')/items(${fileId})/FieldValuesForEdit`,
        {
          method: 'GET',
          headers: {
            Accept: 'application/json',
            'Content-Type': 'application/json',
          },
        }
      );
      const results = await res.json();
      const responseEncoded = JSON.stringify(results)
        .split('_x005f_')
        .join('_');
      const responseRemoveOdata = responseEncoded.split('OData_').join('');
      allFieldValueStore = JSON.parse(responseRemoveOdata);
      console.log(allFieldValueStore);
      return allFieldValueStore;
    } catch (error) {
      console.log('error occured' + error);
    }
  }

  //Processing Admin Data and making All further validations to upload file with metadata
  public getAdminData(
    fileName,
    filecontenFttypename,
    fileUrl,
    requestUrl,
    fileExtensionOnly,
    siteurl
  ) {
    dialog.show();
    const siteUrl = this.context.pageContext.web.absoluteUrl;
    const spRequiredfieldValuesCheck = [];
    const allSpFieldsFromAdmin = [];
    const spRequiredfieldsFromAdmin = [];
    const laserficheFieldsFromAdmin = [];
    let table = '';
    const spfarray = [];
    let fieldvalue;
    const requiredFields = [];
    const requiredFieldsName = [];
    const missigRequiredFieldsValues = [];
    const missigRequiredFieldsValuesNames = [];

    this.context.spHttpClient
      .get(
        siteUrl +
          "/_api/web/lists/getByTitle('AdminConfigurationList')/items?$filter=Title eq 'ManageMapping'&$top=1",
        SPHttpClient.configurations.v1,
        {
          headers: {
            Accept: 'application/json',
          },
        }
      )
      .then((response: SPHttpClientResponse) => {
        return response.json();
      })
      .then(async (response) => {
        const manageMappingDetails = JSON.parse(response.value[0]['JsonValue']);
        for (let i = 0; i < manageMappingDetails.length; i++) {
          const Maping = manageMappingDetails[i]['SharePointContentType'];
          spfarray.push(Maping);

          // we check whether the contentype of selected file is have a mapping or not
          if (filecontenttypename === Maping) {
            const laserficheProfile =
              manageMappingDetails[i]['LaserficheContentType'];

            this.context.spHttpClient
              .get(
                siteUrl +
                  "/_api/web/lists/getByTitle('AdminConfigurationList')/items?$filter=Title eq 'ManageConfigurations'&$top=1",
                SPHttpClient.configurations.v1,
                {
                  headers: {
                    Accept: 'application/json',
                  },
                }
              )
              .then((response1: SPHttpClientResponse) => {
                return response1.json();
              })
              .then((response1) => {
                const Config = JSON.parse(response1.value[0]['JsonValue']);
                for (let j = 0; j < Config.length; j++) {
                  const configname = Config[j]['ConfigurationName'];

                  //below we get Laserfiche Profile details that is mapped to the content type of selected file
                  if (laserficheProfile === configname) {
                    const mappedProfileDocumentName = Config[j]['DocumentName'];
                    const mappedProfileDestinationFolder = Config[j]['EntryId'];
                    const mappedProfileAction = Config[j]['Action'];
                    const mappedProfileTemplate = Config[j]['DocumentTemplate'];
                    if (mappedProfileTemplate !== 'None') {
                      const SPFields = Config[j]['SharePointFields'];
                      const LFFields = Config[j]['LaserficheFields'];

                      for (let k = 0; k < SPFields.length; k++) {
                        const spFieldName = SPFields[k].split('[')[0];
                        allSpFieldsFromAdmin.push(spFieldName);
                        const lfFieldName = LFFields[k].split('[')[0];
                        laserficheFieldsFromAdmin.push(lfFieldName);
                        const lffieldrequired = LFFields[k].split('[')[2];
                        const lffieldrequired1 = lffieldrequired.slice(0, -1);
                        const lffieldrequired2 = lffieldrequired1.split(':')[1];
                        if (lffieldrequired2 == 'true') {
                          const spfieldsrequired = SPFields[k].split('[')[0];
                          spRequiredfieldsFromAdmin.push(spfieldsrequired);
                        }
                      }

                      //checking whether the required fields mapped in admin configuration is present in Library
                      const requiredFieldsCheckinLibrary =
                        spRequiredfieldsFromAdmin.filter((element) =>
                          staticFieldNames.includes(element)
                        );
                      //missing Required fields from Library
                      const requiredFieldsmissing = $(spRequiredfieldsFromAdmin)
                        .not(requiredFieldsCheckinLibrary)
                        .get();

                      /* if(requiredFieldsmissing.length!=0){
                  for(let l=0; l<requiredFieldsmissing.length; l++){
                    var requiredStaticName=requiredFieldsmissing[l];
                    for(let f=0; f<fieldDataDisplay.length; f++){
                      if(fieldDataDisplay[f][requiredStaticName]!=undefined){
                        missigRequiredFields.push(fieldDataDisplay[f][requiredStaticName]);
                      }
                    }
                  }
                } */

                      const missingRequiredFieldsNames = [
                        ...new Set(requiredFieldsmissing),
                      ];

                      //getting internal names of required fields present in the library
                      for (const q of requiredFieldsCheckinLibrary /* =0; q<requiredFieldsCheckinLibrary.length; q++ */) {
                        const nameField =
                          q; /* requiredFieldsCheckinLibrary[q] */
                        for (const w of Object.keys(
                          fieldDataStatic
                        ) /* =0; w<fieldDataStatic.length; w++ */) {
                          if (fieldDataStatic[w][nameField] != undefined) {
                            const requiredFieldInternal =
                              fieldDataStatic[w][nameField];
                            requiredFields.push(requiredFieldInternal);
                          }
                        }
                      }

                      const requiredfieldsCountFromAdmin =
                        spRequiredfieldsFromAdmin.length;
                      const requiredfieldsCountFromLibrary =
                        requiredFieldsCheckinLibrary.length;

                      // checking whether all the required sharepoint fields are present in the library
                      if (
                        requiredfieldsCountFromAdmin ==
                        requiredfieldsCountFromLibrary
                      ) {
                        // checking whether all the required fields have values
                        for (
                          let b = 0;
                          b < requiredfieldsCountFromLibrary;
                          b++
                        ) {
                          const spfieldname2 = requiredFields[b];
                          const fieldvaluerequired =
                            allFieldValueStore[spfieldname2];
                          if (fieldvaluerequired != '') {
                            spRequiredfieldValuesCheck.push(fieldvaluerequired);
                          } else {
                            missigRequiredFieldsValues.push(requiredFields[b]);
                          }
                        }

                        // getting display names of Required SharePoint fields to show in the error dialog
                        if (missigRequiredFieldsValues.length != 0) {
                          for (const z of missigRequiredFieldsValues /* =0; z<missigRequiredFieldsValues.length; z++ */) {
                            const missingRequiredvalueFieldInternalName =
                              z; /* missigRequiredFieldsValues[z] */
                            for (const s of Object.keys(
                              fieldDataInternal
                            ) /* =0; s<fieldDataInternal.length; s++ */) {
                              if (
                                fieldDataInternal[s][
                                  missingRequiredvalueFieldInternalName
                                ] != undefined
                              ) {
                                missigRequiredFieldsValuesNames.push(
                                  fieldDataInternal[s][
                                    missingRequiredvalueFieldInternalName
                                  ]
                                );
                              }
                            }
                          }
                        }

                        const missigRequiredFieldsvaluesNames = [
                          ...new Set(missigRequiredFieldsValuesNames),
                        ];

                        // checking if all the reuired fields present in the Library doesn't have Blanks
                        if (
                          requiredfieldsCountFromAdmin ==
                          spRequiredfieldValuesCheck.length
                        ) {
                          // getting internal names of all SharePoint fields From Admin Profile Mapping
                          for (
                            let y = 0;
                            y < allSpFieldsFromAdmin.length;
                            y++
                          ) {
                            const nameFieldAll = allSpFieldsFromAdmin[y];
                            for (
                              let v = 0;
                              v < fieldDataStaticAll.length;
                              v++
                            ) {
                              if (
                                fieldDataStaticAll[v][nameFieldAll] != undefined
                              ) {
                                const requiredFieldInternalName =
                                  fieldDataStaticAll[v][nameFieldAll];
                                requiredFieldsName.push(
                                  requiredFieldInternalName
                                );
                              } else {
                                if (
                                  v == fieldDataStaticAll.length - 1 &&
                                  requiredFieldsName.length < y + 1
                                ) {
                                  requiredFieldsName.push(nameFieldAll);
                                }
                              }
                            }
                          }

                          // for every mapping getting Values and assigning to Laserfiche Field
                          for (
                            let m = 0;
                            m < allSpFieldsFromAdmin.length;
                            m++
                          ) {
                            const spfieldname1 = requiredFieldsName[m];
                            const spFieldIndex = m;
                            fieldvalue = allFieldValueStore[spfieldname1];

                            if (fieldvalue != undefined || fieldvalue != null) {
                              const Fieldvaluelength = fieldvalue.length;
                              const LFfield =
                                laserficheFieldsFromAdmin[spFieldIndex]; // Laserfiche Field name

                              //Checking Laserfiche Field Type
                              for (let o = 0; o < LFFields.length; o++) {
                                const LFFields1 = LFFields[o];
                                const result = LFFields1.startsWith(LFfield);
                                if (result == true) {
                                  const Lflength = LFFields1.split('[')[3];
                                  const Lflength1 = Lflength.slice(0, -1);
                                  const Lflength2 = Lflength1.split(':')[1];
                                  const LFFieldtype1 = LFFields1.split('[')[1];
                                  const LFFieldtype = LFFieldtype1.slice(0, -1);
                                  if (Lflength2 != 0) {
                                    if (Fieldvaluelength > Lflength2) {
                                      fieldvalue = fieldvalue.slice(
                                        0,
                                        Lflength2
                                      );
                                    }
                                  } else if (
                                    LFFieldtype != 'DateTime' ||
                                    LFFieldtype != 'Time' ||
                                    LFFieldtype != 'Date'
                                  ) {
                                    if (LFFieldtype == 'ShortInteger') {
                                      const extractOnlynumbers =
                                        fieldvalue.replace(/[^0-9]/g, '');
                                      const extractOnlynumberslength =
                                        extractOnlynumbers.length;
                                      if (extractOnlynumberslength > 5) {
                                        fieldvalue = extractOnlynumbers.slice(
                                          0,
                                          5
                                        );
                                      } else {
                                        fieldvalue = extractOnlynumbers;
                                      }
                                    } else if (LFFieldtype == 'LongInteger') {
                                      const extractOnlynumbersLonginteger =
                                        fieldvalue.replace(/[^0-9]/g, '');
                                      const extractOnlynumbersLongintegerlength =
                                        extractOnlynumbersLonginteger.length;
                                      if (
                                        extractOnlynumbersLongintegerlength > 10
                                      ) {
                                        fieldvalue =
                                          extractOnlynumbersLonginteger.slice(
                                            0,
                                            10
                                          );
                                      } else {
                                        fieldvalue =
                                          extractOnlynumbersLonginteger;
                                      }
                                    } else if (LFFieldtype == 'Number') {
                                      const valueOnlyNumbers =
                                        fieldvalue.replace(/[^0-9.]/g, '');
                                      const valueOnlyNumberssplit =
                                        valueOnlyNumbers.split('.');
                                      if (valueOnlyNumberssplit.length === 1) {
                                        const valueOnlyNumbersLimitcheck =
                                          valueOnlyNumbers.split('.')[0];
                                        if (
                                          valueOnlyNumbersLimitcheck.length > 13
                                        ) {
                                          fieldvalue =
                                            valueOnlyNumbersLimitcheck.slice(
                                              0,
                                              13
                                            );
                                        } else {
                                          fieldvalue = valueOnlyNumbers;
                                        }
                                      } else {
                                        const valueOnlyNumbersbfrPeriod =
                                          valueOnlyNumbers.split('.')[0];
                                        const valueOnlyNumbersafrPeriod =
                                          valueOnlyNumbers.split('.')[1];
                                        if (
                                          valueOnlyNumbersbfrPeriod.length <=
                                            13 &&
                                          valueOnlyNumbersafrPeriod.length <= 5
                                        ) {
                                          fieldvalue = valueOnlyNumbers;
                                        } else {
                                          const valueOnlyNumbersbfrPeriod1 =
                                            valueOnlyNumbersbfrPeriod.slice(
                                              0,
                                              13
                                            );
                                          const valueOnlyNumbersafrPeriod1 =
                                            valueOnlyNumbersafrPeriod.slice(
                                              0,
                                              5
                                            );
                                          fieldvalue =
                                            valueOnlyNumbersbfrPeriod1 +
                                            '.' +
                                            valueOnlyNumbersafrPeriod1;
                                        }
                                      }
                                    }
                                  }
                                }
                              }
                              fieldvalue = fieldvalue.replace(/[\\]/g, `\\\\`);
                              fieldvalue = fieldvalue.replace(/["]/g, `\\"`);
                              table +=
                                '"' +
                                LFfield +
                                '"' +
                                ':{"values": [{"value":' +
                                '"' +
                                fieldvalue +
                                '"' +
                                ',"position": 1}]},';
                              fieldvalue = '';
                            } else {
                              fieldvalue = '';
                            }
                          }
                          console.log(table);
                          const table1 = table.slice(0, -1);
                          //var fieldmetadata='{"template":'+'"'+MapDocTemplate+'"'+',"metadata": {"fields": {'+table+'}}}';
                          const fieldmetadata =
                            '{"metadata": {"fields": {' +
                            table1 +
                            '}},"template":' +
                            '"' +
                            mappedProfileTemplate +
                            '"' +
                            '}';
                          console.log(fieldmetadata);
                          window.localStorage.setItem(
                            'Filemetadata',
                            fieldmetadata
                          );
                          window.localStorage.setItem('Filename', fileName);
                          window.localStorage.setItem(
                            'Documentname',
                            mappedProfileDocumentName
                          );
                          window.localStorage.setItem(
                            'DocTemplate',
                            mappedProfileTemplate
                          );
                          window.localStorage.setItem(
                            'Action',
                            mappedProfileAction
                          );
                          window.localStorage.setItem('Fileurl', fileUrl);
                          window.localStorage.setItem(
                            'Destinationfolder',
                            mappedProfileDestinationFolder
                          );
                          window.localStorage.setItem(
                            'Filedataurl',
                            requestUrl
                          );
                          window.localStorage.setItem(
                            'Fileextension',
                            fileExtensionOnly
                          );
                          window.localStorage.setItem('Siteurl', siteUrl);
                          window.localStorage.setItem('SiteUrl', siteurl);
                          window.localStorage.setItem('Maping', Maping);
                          window.localStorage.setItem(
                            'Filecontenttype',
                            filecontenttypename
                          );
                          window.localStorage.setItem(
                            'LContType',
                            laserficheProfile
                          );
                          window.localStorage.setItem('configname', configname);
                          Navigation.navigate(siteUrl + Redirectpagelink, true);
                        } else {
                          document.getElementById('it').innerHTML =
                            'The following SharePoint field values are blank and are mapped to required Laserfiche fields:<br/>&ensp;-' +
                            missigRequiredFieldsvaluesNames.join(
                              '<br/>&ensp;-'
                            ) +
                            '<br/><br/>Please fill out these required fields and try again.';
                          document.getElementById('imgid').style.display =
                            'none';
                          //document.getElementById("ref").style.display='block';
                          document.getElementById('divid').style.display =
                            'block';
                          document.getElementById('divid1').onclick = this.Dc;
                          document.getElementById('divid13').style.display =
                            'none';
                          staticFieldNames = [];
                          fieldDataDisplay = [];
                          fieldDataStatic = [];
                          fieldDataInternal = [];
                          fieldDataStaticAll = [];
                          allFieldValueStore = [];
                        }
                      } else {
                        document.getElementById('it').innerHTML =
                          'The following SharePoint fields are not available in the library and are mapped to required Laserfiche fields:<br/>&ensp;-' +
                          missingRequiredFieldsNames.join('<br/>&ensp;-') +
                          '<br/><br/>Note:These are the internal names of the SharePoint fields';
                        document.getElementById('imgid').style.display = 'none';
                        //document.getElementById("ref").style.display='block';
                        document.getElementById('divid').style.display =
                          'block';
                        document.getElementById('divid1').onclick = this.Dc;
                        document.getElementById('divid13').style.display =
                          'none';
                        staticFieldNames = [];
                        fieldDataDisplay = [];
                        fieldDataStatic = [];
                        fieldDataInternal = [];
                        fieldDataStaticAll = [];
                        allFieldValueStore = [];
                      }
                    } else {
                      window.localStorage.setItem('Filename', fileName);
                      window.localStorage.setItem(
                        'Documentname',
                        mappedProfileDocumentName
                      );
                      window.localStorage.setItem(
                        'DocTemplate',
                        mappedProfileTemplate
                      );
                      window.localStorage.setItem(
                        'Action',
                        mappedProfileAction
                      );
                      window.localStorage.setItem('Fileurl', fileUrl);
                      window.localStorage.setItem(
                        'Destinationfolder',
                        mappedProfileDestinationFolder
                      );
                      window.localStorage.setItem('Filedataurl', requestUrl);
                      window.localStorage.setItem(
                        'Fileextension',
                        fileExtensionOnly
                      );
                      window.localStorage.setItem('Siteurl', siteUrl);
                      window.localStorage.setItem('SiteUrl', siteurl);
                      window.localStorage.setItem('Maping', Maping);
                      window.localStorage.setItem(
                        'Filecontenttype',
                        filecontenttypename
                      );
                      window.localStorage.setItem(
                        'LContType',
                        laserficheProfile
                      );
                      window.localStorage.setItem('configname', configname);
                      Navigation.navigate(siteUrl + Redirectpagelink, true);
                    }
                  }
                }
              });
          }
        }
        console.log(spfarray);
        if (spfarray.indexOf(filecontenttypename) === -1) {
          window.localStorage.setItem('Filename', fileName);
          // window.localStorage.setItem('Maping', Maping); // TODO this doesn't exist because it is per manageMappingDetails
          window.localStorage.setItem('Filecontenttype', filecontenttypename);
          window.localStorage.setItem('Fileurl', fileUrl);
          window.localStorage.setItem('Filedataurl', requestUrl);
          window.localStorage.setItem('Fileextension', fileExtensionOnly);
          window.localStorage.setItem('Siteurl', siteUrl);
          window.localStorage.setItem('SiteUrl', siteurl);
          // window.localStorage.setItem('LContType', laserficheProfile); // TODO this doesn't exist because it is per manageMappingDetails
          Navigation.navigate(siteUrl + Redirectpagelink, true);
        }
      });
  }
  //

  private Dc() {
    dialog.close();
  }
}
