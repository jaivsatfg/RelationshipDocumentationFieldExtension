import * as React from 'react';
import * as ReactDOM from 'react-dom';

import { Guid, Log } from '@microsoft/sp-core-library';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import {
  BaseFieldCustomizer,
  type IFieldCustomizerCellEventParameters
} from '@microsoft/sp-listview-extensibility';


import * as strings from 'RelatedDocumentFieldExtensionFieldCustomizerStrings';
import { IFieldConfig } from './IFieldConfig';
import { ISoapTaxonomyResponse } from './components/ISoapTaxonomyResponse';
import { ITaxonomyElementT, ITaxonomyElementTL, ITaxonomyElementTM, ITaxonomyResponse } from './components/ITaxonomyResponse';
import { XMLParser } from 'fast-xml-parser';
import RelatedDocumentFieldExtension, { IRelatedDocumentFieldExtensionProps } from './components/RelatedDocumentFieldExtension';
import { dateAdd, PnPClientStorage } from "@pnp/core";

/**
 * If your field customizer uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */


export interface IRelatedDocumentFieldExtensionFieldCustomizerProperties {
  // This is an example; replace with your own property
  termStoreId: string;
  termSetId: string;
  servicesTermId: string;
  notificationTypeTermId: string;
}

const LOG_SOURCE: string = 'RelatedDocumentFieldExtensionFieldCustomizer';

export default class RelatedDocumentFieldExtensionFieldCustomizer
  extends BaseFieldCustomizer<IRelatedDocumentFieldExtensionFieldCustomizerProperties> {

  private fieldConfig: IFieldConfig = {
    termStoreId: '',
    termSetId: '',
    serveiTermId: '',
    tipologiaTermId: '',
    managerDocumentsUrl: '',
    managerDocumentsListId: ''
  };
  private terms: ISoapTaxonomyResponse[];

  public async onInit(): Promise<void> {
    try {
      return await new Promise<void>((resolve, reject) => {
        // Add your custom initialization to this method.  The framework will wait
        // for the returned promise to resolve before firing any BaseFieldCustomizer events.
        Log.info(LOG_SOURCE, 'Activated DocumentsRelacionatsFieldCustomizer with properties:');
        Log.info(LOG_SOURCE, JSON.stringify(this.properties, undefined, 2));
        Log.info(LOG_SOURCE, `The following string should be equal: "DocumentsRelacionatsFieldCustomizer" and "${strings.Title}"`);

        this.getFieldConfig().then((fcfg: IFieldConfig) => {
          this.fieldConfig.termStoreId = fcfg.termStoreId || this.properties.termStoreId || 'fd67ea3b-3045-4cd7-a213-76a221c6482e';
          this.fieldConfig.termSetId = fcfg.termSetId || this.properties.termSetId || 'fd67ea3b-3045-4cd7-a213-76a221c6482e';
          this.fieldConfig.serveiTermId = fcfg.serveiTermId || this.properties.servicesTermId || 'f5011dce-f5f6-49c4-b641-f07c065454da';
          this.fieldConfig.tipologiaTermId = fcfg.tipologiaTermId || this.properties.notificationTypeTermId || 'e8247b94-53d4-46bf-9472-9dacda3d2a0c';
          this.fieldConfig.managerDocumentsUrl = fcfg.managerDocumentsUrl;
          this.fieldConfig.managerDocumentsListId = fcfg.managerDocumentsListId;

          const keyLocalStore = "fldTaxonomyTree".concat('_',
            this.context && this.context.pageContext && this.context.pageContext.list ? this.context.pageContext.list.id.toString() : '');
          const store = new PnPClientStorage();
          const taxonomyTreeValues: ISoapTaxonomyResponse[] = store.local.get(keyLocalStore);
          if (Array.isArray(taxonomyTreeValues) && taxonomyTreeValues.length > 0) {
            this.terms = taxonomyTreeValues;
            resolve();
            return;
          }

          this.LoadTerms({
            termStoreId: this.fieldConfig.termStoreId,
            lcid: 1033,
            termId: this.fieldConfig.tipologiaTermId,
            termSetId: this.fieldConfig.termSetId
          }).then((tipologiaTerms: ISoapTaxonomyResponse[]) => {
            this.LoadTerms({
              termStoreId: this.fieldConfig.termStoreId,
              lcid: 1033,
              termId: this.fieldConfig.serveiTermId,
              termSetId: this.fieldConfig.termSetId,
              tipologiasTerms: tipologiaTerms
            }).then((terms: ISoapTaxonomyResponse[]) => {
              this.terms = terms;
              store.local.put(keyLocalStore, terms, dateAdd(new Date(), "minute", 30));
              resolve();
            }).catch(() => {
              reject("ERROR TRATANDO DE BUSCAR EL NODO SERVEI.");
            });
          }).catch(() => {
            reject("ERROR TRATANDO DE BUSCAR EL NODO TIPOLOGIA.");
          });
        }).catch((a: string) => {
          console.log('Error in getFieldConfig():' + a);
        });
      });
    } catch (e) {
      console.log('Error in getFieldConfig():' + e);
    }
  }

  public onRenderCell(event: IFieldCustomizerCellEventParameters): void {
    // Use this method to perform your custom cell rendering.
    // const text: string = `${this.properties.sampleText}: ${event.fieldValue}`;
    let unitUrl: string = '';
    let documentsManagerUrl: string = '';
    let documentsManagerRelativeUrl: string = '';
    let documentsManagerListId: string = '';
    let documentsManagerListTitle: string = '';
    let documentsManagerListItemId: number = 0;
    let elementName = event.listItem.getValueByName('FileRef').split('/')[event.listItem.getValueByName('FileRef').split('/').length - 1];

    if (this.context && this.context.pageContext && this.context.pageContext.list &&
      (this.context.pageContext.list.title.toLowerCase() === 'notificaciones'
        || this.context.pageContext?.list.title.toLowerCase() === 'DocumentosTrabajo')) {
      unitUrl = this.context.pageContext.web.absoluteUrl;
      documentsManagerUrl = this.fieldConfig.managerDocumentsUrl;
      documentsManagerRelativeUrl = [''].concat(documentsManagerUrl.split('/').slice(3)).join('/');
      documentsManagerListTitle = 'DocumentosPublicados';
      documentsManagerListId = this.fieldConfig.managerDocumentsListId;
      if (this.context.pageContext.list.title.toLowerCase() === 'docstreball') {
        documentsManagerListItemId = event.listItem.getValueByName('PublicDocumentId');
      } else {
        elementName = event.listItem.getValueByName('Title');
      }
    } else if (this.context && this.context.pageContext && this.context.pageContext.list) {
      documentsManagerUrl = this.context.pageContext.web.absoluteUrl;
      documentsManagerRelativeUrl = this.context.pageContext.web.serverRelativeUrl;
      documentsManagerListTitle = this.context.pageContext.list.title;
      documentsManagerListId = this.context.pageContext.list.id.toString();
      documentsManagerListItemId = event.listItem.getValueByName('ID');
    }

    const props: IRelatedDocumentFieldExtensionProps = {
      terms: this.terms,
      pageContext: this.context.pageContext,
      webUrl: unitUrl,
      documentsManagerWebUrl: documentsManagerUrl,
      documentsManagerRelativeWebUrl: documentsManagerRelativeUrl,
      documentsManagerListId: documentsManagerListId,
      documentsManagerListTitle: documentsManagerListTitle,
      documentsManagerListItemId: documentsManagerListItemId,
      localListTitle: this.context && this.context.pageContext && this.context.pageContext.list ? this.context.pageContext.list.title : '',
      localListId: this.context && this.context.pageContext && this.context.pageContext.list ? this.context.pageContext.list.id : Guid.empty,
      localListItemId: event.listItem.getValueByName('ID'),
      elementName: elementName,
      textValue: event.listItem.getValueByName('DocuRelaJson'),
      FSObjType: event.listItem.getValueByName('FSObjType'),
      spHttpClient: this.context.spHttpClient,
      spHttpConfiguration: SPHttpClient.configurations.v1,
      insertSharePointItem: this.insertItem,
      updateSharePointItem: this.updatetItem,
      deleteSharePointItem: this.deleteItem
    };

    const relatedDocument: React.ReactElement<{}> =
      React.createElement(RelatedDocumentFieldExtension, props);

    ReactDOM.render(relatedDocument, event.domElement);
  }

  public onDisposeCell(event: IFieldCustomizerCellEventParameters): void {
    // This method should be used to free any resources that were allocated during rendering.
    // For example, if your onRenderCell() called ReactDOM.render(), then you should
    // call ReactDOM.unmountComponentAtNode() here.
    ReactDOM.unmountComponentAtNode(event.domElement);
    super.onDisposeCell(event);
  }

  private async LoadTerms(param: { termStoreId: string; lcid: number; termId: string; termSetId: string; tipologiasTerms?: ISoapTaxonomyResponse[]; }): Promise<ISoapTaxonomyResponse[]> {
    const values: ISoapTaxonomyResponse[] = await this.GetChildTermsInTerm(param);
    for (let i: number = 0; i < values.length; i++) {
      if (values[i].hasChildren) {
        const p = Object.assign({}, param, { termId: values[i].id });
        values[i].children = await this.LoadTerms(p);
      } else {

        values[i].children = Object.assign([], param.tipologiasTerms ? param.tipologiasTerms.map((t: ISoapTaxonomyResponse) => {
          const currentObj: ISoapTaxonomyResponse = this.deepCopy(t);

          if (currentObj.info && values && currentObj.info.length > 0 && values.length > 0) {
            let currentValue: ISoapTaxonomyResponse = values[i];
            if (currentValue.info && currentValue.info.length > 0) {
              currentObj.info[0].parentLabel = '' + currentValue.info[0].parentLabel + ';' + values[i].text;
            }
          }
          return currentObj;
        }) : []);
      }
    }
    return values;
  }

  private async GetChildTermsInTerm(param: { termStoreId: string; lcid: number; termId: string; termSetId: string; }): Promise<ISoapTaxonomyResponse[]> {
    return new Promise<ISoapTaxonomyResponse[]>((resolve, reject) => {
      let siteUrl: string = this.context.pageContext.web.absoluteUrl;
      if (siteUrl.charAt(siteUrl.length - 1) === '/') siteUrl = siteUrl.substring(0, siteUrl.length - 1);
      const webServiceURL: string = siteUrl + "/_vti_bin/TaxonomyClientService.asmx?op=GetChildTermsInTerm";
      const soapMessage =
        '<soap12:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap12="http://www.w3.org/2003/05/soap-envelope">'
        + '<soap12:Body>'
        + '<GetChildTermsInTerm xmlns="http://schemas.microsoft.com/sharepoint/taxonomy/soap/">'
        + '<sspId>' + param.termStoreId + '</sspId>'
        + '<lcid>' + param.lcid + '</lcid>'
        + '<termId>' + param.termId + '</termId>'
        + '<termSetId>' + param.termSetId + '</termSetId>'
        + '</GetChildTermsInTerm>'
        + '</soap12:Body>'
        + '</soap12:Envelope>';


      this.context.spHttpClient.post(webServiceURL, SPHttpClient.configurations.v1, {
        headers: {
          'Accept': 'xml',
          'Content-type': "text/xml; charset=\"utf-8\""
        },
        body: soapMessage
      }).then((response: SPHttpClientResponse) => {
        if (response.ok) {
          response.text().then((xml: string) => {
            const parser = new XMLParser({
              ignoreAttributes: false,
              attributeNamePrefix: "_",
              // alwaysCreateTextNode: true
            });
            const soapResponse = parser.parse(xml);
            const soapDataResponse: ITaxonomyResponse = parser.parse(soapResponse['soap:Envelope']['soap:Body'].GetChildTermsInTermResponse.GetChildTermsInTermResult);
            const values: ISoapTaxonomyResponse[] = [];
            if (Array.isArray(soapDataResponse.TermStore.T)) {
              soapDataResponse.TermStore.T.map((t: ITaxonomyElementT) => {
                const currentValue: ISoapTaxonomyResponse = {
                  selected: false,
                  description: t.TD ? t.TD.a11 : '',
                  labels: t.LS && Object.keys(t.LS).map((l: string, index: number) => {
                    const current: ITaxonomyElementTL = t.LS.TL;
                    return {
                      value: current._a32,
                      isDefault: current._a31 === 'true' ? true : false
                    };
                  }),
                  text: t.LS && Object.keys(t.LS).map((l: string, index: number) => {
                    const current: ITaxonomyElementTL = t.LS.TL;
                    return current._a32;
                  })[0],
                  info: t.TMS && Object.keys(t.TMS).map((i: string, index: number) => {
                    const current: ITaxonomyElementTM = t.TMS.TM;
                    return {
                      parentId: current._a25,
                      parentLabel: current._a40,
                      termPath: current._a45,
                      children: current._a67 ? current._a67.split(':') : [],
                      hasChildren: current._a69 === 'true' ? true : false,
                      termSetId: current._a24,
                      termSetLabel: current._a12
                    }
                  }),
                  id: t._a9,
                  isDeprecated: t._a21 === 'true' ? true : false,
                  internalId: t._a61,
                  hasChildren: false,
                  children: []
                };
                if (Array.isArray(currentValue.info) && currentValue.info[0] && currentValue.info[0].hasChildren) {
                  currentValue.hasChildren = true;
                }
                values.push(currentValue);
              });
              resolve(values);
            } else {
              const t: ITaxonomyElementT = soapDataResponse.TermStore.T;
              const currentValue: ISoapTaxonomyResponse = {
                selected: false,
                description: t.TD ? t.TD.a11 : '',
                labels: t.LS && Object.keys(t.LS).map((l: string, index: number) => {
                  const current: ITaxonomyElementTL = t.LS.TL;
                  return {
                    value: current._a32,
                    isDefault: current._a31 === 'true' ? true : false
                  };
                }),
                text: t.LS && Object.keys(t.LS).map((l: string, index: number) => {
                  const current: ITaxonomyElementTL = t.LS.TL;
                  return current._a32;
                })[0],
                info: t.TMS && Object.keys(t.TMS).map((i: string, index: number) => {
                  const current: ITaxonomyElementTM = t.TMS.TM;
                  return {
                    parentId: current._a25,
                    parentLabel: current._a40,
                    termPath: current._a45,
                    children: current._a67 ? current._a67.split(':') : [],
                    hasChildren: current._a69 === 'true' ? true : false,
                    termSetId: current._a24,
                    termSetLabel: current._a12
                  }
                }),
                id: t._a9,
                isDeprecated: t._a21 === 'true' ? true : false,
                internalId: t._a61,
                hasChildren: false,
                children: []
              };
              if (Array.isArray(currentValue.info) && currentValue.info[0] && currentValue.info[0].hasChildren) {
                currentValue.hasChildren = true;
              }
              values.push(currentValue);
              resolve(values);
            }
          })
            .catch((ex: string) => {
              console.error(ex);
              reject([]);
            });
        }
        else {
          console.error('response with error.');
          console.error(response);
          reject([]);
        }
      }).catch((ex: SPHttpClientResponse) => {
        console.error(ex);
        reject([]);
      });
    });
  }

  private async getFieldConfig(): Promise<IFieldConfig> {
    return new Promise<IFieldConfig>((resolve, reject) => {
      const store = new PnPClientStorage();

      const keyLocalStore = "fldConfig".concat('_', this.context && this.context.pageContext && this.context.pageContext.list ? this.context.pageContext.list.id.toString() : '');
      const cfgFld: IFieldConfig = store.local.get(keyLocalStore);
      if (cfgFld !== null) {
        resolve(cfgFld);
        return;
      }

      const listName = 'AppParametros';

      /* El filtro se hace a la lista AppParametros, 
        pero hay que saber que puede ser consultada en el sitecollection de la unidad
        o el sitecollection de la gestion documental. 
        Si se hace desde la unidad hay que determinar cual es el id y title de la lista en la gestion documental
        por lo que el filtro a title igual a DocumentPublicUrl y title igual a DocumentPublicListId son los que darán esa info
        si se hace directamente sobre la gestion documental la info será extraida de Title igual al id de la lista

        Tener claro esto.
      */

      const itemTitle = this.context && this.context.pageContext && this.context.pageContext.list ? this.context.pageContext.list.id : null;
      const params = {
        '$select': 'Title,ItemValor',
        '$filter': `Title eq '${itemTitle}' or Title eq 'PublicLibraryUrl' or Title eq 'PublicLibraryId'`
      };


      const currentUrl = this.context.pageContext.web.absoluteUrl.concat(`/_api/web/Lists/GetByTitle('`, encodeURIComponent(listName), `')`,
        '/items?', jQuery.param(params));

      this.context.spHttpClient.get(currentUrl, SPHttpClient.configurations.v1)
        .then((response: SPHttpClientResponse) => {
          if (response.ok) {
            type itemValue = { Title: string; ItemValor: string; };
            type responseQuery = { value: itemValue[] };
            response.json().then((r: responseQuery) => {
              const termsIds: itemValue = r.value.filter((it: itemValue) => {
                const itemTitle = this.context && this.context.pageContext && this.context.pageContext.list ? this.context.pageContext.list.id : '';
                return it.Title.toLowerCase() === itemTitle.toString().toLocaleLowerCase();
              })[0];
              const documentsPublicsUrl: itemValue = r.value.filter((it: itemValue) => {
                return it.Title.toLocaleLowerCase() === 'publiclibraryurl';
              })[0];
              const documentsPublicsId: itemValue = r.value.filter((it: itemValue) => {
                return it.Title.toLocaleLowerCase() === 'publiclibraryid';
              })[0];
              const currentValue: IFieldConfig = termsIds && JSON.parse(termsIds.ItemValor) || {};
              currentValue.managerDocumentsUrl = documentsPublicsUrl && documentsPublicsUrl.ItemValor || '';
              currentValue.managerDocumentsListId = documentsPublicsId && documentsPublicsId.ItemValor || '';

              if (currentValue.managerDocumentsUrl !== ''
                && currentValue.managerDocumentsListId !== ''
                && currentValue.serveiTermId !== undefined
                && currentValue.termSetId !== undefined
                && currentValue.termStoreId !== undefined
                && currentValue.tipologiaTermId !== undefined) {
                store.local.put(keyLocalStore, currentValue, dateAdd(new Date(), "minute", 30));
              }
              resolve(currentValue);
            }).catch((e: SPHttpClientResponse) => {
              console.error(e);
              reject('ERROR');
            });
          } else {
            response.json().then((r: { error: { code: string; message: string; } }) => {
              reject('ERROR response: Status: '.concat(
                response.status.toString(),
                response.statusText ? ' - '.concat(response.statusText) : '',
                r.error ? ' details: '.concat(r.error.code + " -- " + r.error.message) : ''));
            }).catch((e: SPHttpClientResponse) => {
              console.error(e);
              reject('ERROR');
            });
          }
        })
        .catch((e: SPHttpClientResponse) => {
          console.error(e);
          reject('ERROR');
        });

    });
  }

  private deepCopy<T>(instance: T): T {
    if (instance === null) {
      return instance;
    }

    // handle Dates
    if (instance instanceof Date) {
      return new Date(instance.getTime()) as any;
    }

    // handle Array types
    if (instance instanceof Array) {
      const cloneArr = [] as any[];
      (instance as any[]).forEach((value) => { cloneArr.push(value) });
      // for nested objects
      return cloneArr.map((value: any) => this.deepCopy<any>(value)) as any;
    }
    // handle objects
    if (instance instanceof Object) {
      const copyInstance = {
        ...(instance as { [key: string]: any }
        )
      } as { [key: string]: any };
      for (const attr in instance) {
        if ((instance as Record<string, any>).hasOwnProperty(attr)) {
          copyInstance[attr] = this.deepCopy<any>((instance as Record<string, any>)[attr]);
        }
      }
      return copyInstance as T;
    }
    // handling primitive data types
    return instance;
  }

  private deleteItem(webUrl: string, listTitle: string, id: number): Promise<void> {
    return new Promise<void>((resolve, reject) => {
      const queryUrl = `${webUrl}/_api/web/Lists/GetByTitle('${encodeURIComponent(listTitle)}')/items(${id})`;
      this.context.spHttpClient && this.context.spHttpClient.post(queryUrl, SPHttpClient.configurations.v1, {
        headers: {
          'Accept': 'application/json;odata=nometadata',
          'Content-type': 'application/json',
          'If-Match': '*',
          'X-HTTP-Method': 'DELETE',
          'odata-version': ''
        }
      }).then((response: SPHttpClientResponse) => {
        //console.log('item deleted.')
        if (response.ok) {
          resolve();
          return;
        } else {
          response.text().then((r: string) => {
            const e: TypeError = new TypeError('error deleting item.');
            e.stack = 'response: Status: '.concat(
              response.status.toString(),
              response.statusText ? ' - '.concat(response.statusText) : '',
              ' details: '.concat(r)
            );
            reject(e);
          }).catch(() => {
            reject(new TypeError("ERROR Reading response.text() in deleteItem."));
          });
        }
      }).catch((e: SPHttpClientResponse) => {
        console.error(e);
        reject(new TypeError(e.status.toString()));
      });
    });
  }

  private updatetItem<T>(webUrl: string, listTitle: string, id: number, objToUpdate: T): Promise<void> {
    return new Promise<void>((resolve, reject) => {
      const queryUrl = `${webUrl}/_api/web/Lists/GetByTitle('${encodeURIComponent(listTitle)}')/items(${id})`;
      const obj: T = Object.assign({ '__metadata': { 'type': 'SP.ListItem' } }, objToUpdate);
      this.context.spHttpClient.post(queryUrl, SPHttpClient.configurations.v1, {
        headers: {
          'Accept': 'aapplication/json;odata=verbose',
          'Content-type': 'application/json;odata=verbose',
          'If-Match': '*',
          'X-HTTP-Method': 'MERGE',
          'odata-version': ''
        },
        body: JSON.stringify(obj)
      }).then((response: SPHttpClientResponse) => {
        //console.log('item updated.')
        if (response.ok) {
          resolve();
          return;
        } else {
          response.text().then((r: string) => {
            const e: TypeError = new TypeError('error updating item.');
            e.stack = 'response: Status: '.concat(
              response.status.toString(),
              response.statusText ? ' - '.concat(response.statusText) : '',
              ' details: '.concat(r)
            );
            reject(e);
          }).catch(() => {
            reject(new TypeError("ERROR Reading response.text() in updatetItem."));
          });
        }
      }).catch((e: SPHttpClientResponse) => {
        console.error(e);
        reject(new TypeError(e.status.toString()));
      });
    });
  }

  private insertItem<T>(webUrl: string, listTitle: string, objToInsert: T): Promise<number> {
    //https://learn.microsoft.com/en-us/sharepoint/dev/sp-add-ins/working-with-lists-and-list-items-with-rest
    return new Promise<number>((resolve, reject) => {
      const queryUrl = `${webUrl}/_api/web/Lists/GetByTitle('${encodeURIComponent(listTitle)}')/items`;

      const obj: T = Object.assign({ '__metadata': { 'type': 'SP.Data.' + listTitle + 'ListItem' } }, objToInsert);

      this.context.spHttpClient.post(queryUrl, SPHttpClient.configurations.v1, {
        headers: {
          'Accept': 'aapplication/json;odata=verbose',
          'Content-type': 'application/json;odata=verbose',
          'If-Match': '*',
          'odata-version': ''
        },
        body: JSON.stringify(obj)
      }).then((response: SPHttpClientResponse) => {
        if (response.ok) {
          response.text().then((r: string) => {
            resolve(0);
            return;
          }).catch(() => {
            reject(new TypeError("ERROR Reading response.text() in insertItem."));
          });
        } else {
          response.text().then((r: string) => {
            const e: TypeError = new TypeError('error inserting item.');
            e.stack = 'response: Status: '.concat(
              response.status.toString(),
              response.statusText ? ' - '.concat(response.statusText) : '',
              ' details: '.concat(r)
            );
            reject(e);
          }).catch(() => {
            reject(new TypeError("ERROR Reading response.text() in insertItem."));
          });
        }
        //console.log('item updated.')
      }).catch((e: SPHttpClientResponse) => {
        console.error(e);
        reject(new TypeError(e.status.toString()));
      });
    });
  }
}
