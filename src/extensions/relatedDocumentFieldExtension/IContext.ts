import * as React from "react";
import { SPHttpClient, SPHttpClientConfiguration } from '@microsoft/sp-http';
import { Guid } from "@microsoft/sp-core-library";
import { IRelatedDocument } from "./components/IRelatedDocument";

export interface IAppContext {
  webUrl: string;
  documentsManagerWebUrl: string;                 //Web Url donde están los documentos públicos
  documentsManagerRelativeWebUrl: string;         //Web Url Relativa donde están los documentos públicos
  documentsManagerListId: string;                 //Id de la lista con nombre {nnit}Documents donde están los documentos públicos. Ejemplo: ScoDocuments
  documentsManagerListTitle: string;              //Title de la lista con nombre {nnit}Documents donde están los documentos públicos. Ejemplo: ScoDocuments
  documentsManagerListItemId: number;              //Id del item que estoy relacionando de la lista con nombre {nnit}Documents donde están los documentos públicos. Ejemplo: ScoDocuments
  localListTitle: string;                         //Title de la lista desde donde estoy realizando la relación. Puede ser desde Notificacions, DocsTreball o {Unit}Documents
  localListId: Guid;                             //Id de la lista desde donde estoy realizando la relación.
  localListItemId: number;                        //Id del item de la lista desde donde estoy realizando la relación. Aquí se actualizarán los campos DocRelaJson, LinksDocumentos y IdsDocumentosRelacionados
  spHttpClient: SPHttpClient | null;
  spHttpConfiguration: SPHttpClientConfiguration | null;
  insertSharePointItem?: (webUrl: string, listTitle: string, T: IRelatedDocument) => Promise<number>;
  updateSharePointItem?: (webUrl: string, listTitle: string, id: number, T: { LinksDocumentosRelacionados?: string; DocuRelaIds: string; DocuRelaJson: string; }) => Promise<void>;
  deleteSharePointItem?: (webUrl: string, listTitle: string, id: number) => Promise<void>;
}

const appCtx: IAppContext = {
  webUrl: '',
  documentsManagerWebUrl: '',
  documentsManagerRelativeWebUrl: '',
  documentsManagerListId: '',
  documentsManagerListTitle: '',
  documentsManagerListItemId: 0,
  localListTitle: '',  
  localListId: Guid.empty,
  localListItemId: 0,
  spHttpClient: null,
  spHttpConfiguration: null
};

export const AppContext = React.createContext(appCtx);