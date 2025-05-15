import { Guid, Log } from '@microsoft/sp-core-library';
import * as React from 'react';

import styles from './RelatedDocumentFieldExtension.module.scss';
import { DocInfo } from './IDocInfo';
import { ISoapTaxonomyResponse } from './ISoapTaxonomyResponse';
import { IRelatedDocument } from './IRelatedDocument';
import { SPHttpClient, SPHttpClientConfiguration } from '@microsoft/sp-http';
import { AppContext, IAppContext } from '../IContext';
import { IconButton, IIconProps, IStackTokens, ITooltipHostStyles, Stack, StackItem, TooltipHost } from 'office-ui-fabric-react';
import LinkDocuments from './LinkDocuments';


export interface IRelatedDocumentFieldExtensionProps {
  textValue: string;
  FSObjType: string;
  elementName: string;
  terms: ISoapTaxonomyResponse[];
  pageContext: any;
  webUrl: string;
  documentsManagerWebUrl: string;                 //Web Url donde están los documentos públicos
  documentsManagerRelativeWebUrl: string;         //Web Url Relativa donde están los documentos públicos
  documentsManagerListId: string;                 //Id de la lista con nombre 'Documentos públicos' donde están los documentos públicos. Ejemplo: ScoDocuments
  documentsManagerListTitle: string;              //Title de la lista con nombre 'Documentos públicos' donde están los documentos públicos. Ejemplo: ScoDocuments
  documentsManagerListItemId: number;             //Id del item que estoy relacionando de la lista con nombre 'Documentos públicos' donde están los documentos públicos. Ejemplo: ScoDocuments
  localListTitle: string;                         //Title de la lista desde donde estoy realizando la relación. Puede ser desde Notificaciones, Documentos de trabajo o Documentos públicos
  localListId: Guid;                              //Id de la lista desde donde estoy realizando la relación.
  localListItemId: number;                        //Id del item de la lista desde donde estoy realizando la relación. Aquí se actualizarán los campos DocRelaJson, LinksDocumentos y IdsDocumentosRelacionados
  spHttpClient: SPHttpClient;
  spHttpConfiguration: SPHttpClientConfiguration;
  insertSharePointItem: (webUrl: string, listTitle: string, T: IRelatedDocument) => Promise<number>;
  updateSharePointItem: (webUrl: string, listTitle: string, id: number, T: { LinksDocumentosRelacionados?: string; DocuRelaIds: string; DocuRelaJson: string; }) => Promise<void>;
  deleteSharePointItem: (webUrl: string, listTitle: string, id: number) => Promise<void>;
}

export interface IRIRelatedDocumentFieldExtensionStates {
  cellValue: DocInfo[];
  showModal: boolean;
}

const LOG_SOURCE: string = 'RelatedDocumentFieldExtension';

export default class RelatedDocumentFieldExtension extends React.Component<IRelatedDocumentFieldExtensionProps, IRIRelatedDocumentFieldExtensionStates> {
  static modalIsOpen: boolean;
  private ctx: IAppContext = {
    webUrl: this.props.webUrl,
    documentsManagerWebUrl: this.props.documentsManagerWebUrl,
    documentsManagerRelativeWebUrl: this.props.documentsManagerRelativeWebUrl,
    documentsManagerListId: this.props.documentsManagerListId,
    documentsManagerListTitle: this.props.documentsManagerListTitle,
    documentsManagerListItemId: this.props.documentsManagerListItemId,
    localListTitle: this.props.localListTitle,
    localListId: this.props.localListId,
    localListItemId: this.props.localListItemId,
    spHttpClient: this.props.spHttpClient,
    spHttpConfiguration: this.props.spHttpConfiguration,
    insertSharePointItem: this.props.insertSharePointItem,
    updateSharePointItem: this.props.updateSharePointItem,
    deleteSharePointItem: this.props.deleteSharePointItem
  };

  public constructor(props: IRelatedDocumentFieldExtensionProps | Readonly<IRelatedDocumentFieldExtensionProps>) {
    super(props);
    this.onDocRelaLink = this.onDocRelaLink.bind(this);
    this.closeModal = this.closeModal.bind(this);

    let values: DocInfo[] = [];
    try {
      if (this.props.textValue) { values = JSON.parse(this.props.textValue); }
    }
    catch (e: unknown) {
      console.error(e);
    }

    this.state = {
      cellValue: values,
      showModal: false
    }
  }

  public componentDidMount(): void {
    Log.info(LOG_SOURCE, 'React Element: RelatedDocumentFieldExtension mounted');
  }

  public componentWillUnmount(): void {
    Log.info(LOG_SOURCE, 'React Element: RelatedDocumentFieldExtension unmounted');
  }

  public render(): React.ReactElement<{}> {
    const stackTokens: IStackTokens = { childrenGap: 40, padding: '3px;' };
    const calloutProps = { gapSpace: 0 };
    const hostStyles: Partial<ITooltipHostStyles> = { root: { display: 'inline-block' } };
    const textDocIcon: IIconProps = { iconName: 'Link' };
    if (this.props.FSObjType !== '0') {
      return (
        <AppContext.Provider value={this.ctx}>
          <Stack key={this.props.localListItemId} />
        </AppContext.Provider>
      );
    } else {
      return (
        <AppContext.Provider value={this.ctx}>
          <Stack>
            <Stack horizontal className={styles.relatedDocumentFieldExtension} tokens={stackTokens}>
              <StackItem className={styles.documentsRelacionatsItem} key={this.props.localListItemId}>
                <TooltipHost
                  content="Asignación de documentos relacionados"
                  calloutProps={calloutProps}
                  styles={hostStyles}
                  setAriaDescribedBy={false}>
                  <IconButton iconProps={textDocIcon} aria-label="TextDocumentSetting" onClick={this.onDocRelaLink} />
                </TooltipHost>
                {this.state.cellValue === undefined || this.state.cellValue.length === 0 ? <span>sin asignar</span> :
                  <>
                    <ul>
                      {
                        this.state.cellValue.map((d: DocInfo) => {
                          return (<>
                            <li>
                              <div>
                                <a href={(this.ctx.localListTitle === this.ctx.documentsManagerListTitle ? '' : '') + d.url} rel="noreferrer" target="_blank" data-interception="off">{d.name}</a>
                              </div>
                            </li>
                          </>);
                        })
                      }
                    </ul>
                  </>
                }
              </StackItem>
            </Stack>
            {!RelatedDocumentFieldExtension.modalIsOpen ?
              <LinkDocuments elementName={this.props.elementName}
                isModalOpen={this.state.showModal}
                terms={this.props.terms}
                closeModal={this.closeModal} items={this.state.cellValue}
              />
              : <></>
            }
          </Stack>
        </AppContext.Provider>
      );
    }
  }

  private closeModal(savedItems: DocInfo[]): void {
    console.log('closeModal');
    const stateValues: IRIRelatedDocumentFieldExtensionStates = {
      showModal: false,
      cellValue: this.state.cellValue
    };
    if (savedItems !== undefined) {
      stateValues.cellValue = savedItems;
    }
    RelatedDocumentFieldExtension.modalIsOpen = false;
    this.setState(Object.assign({}, stateValues));
  }

  private onDocRelaLink(): void {
    if (this.state.showModal) {
      RelatedDocumentFieldExtension.modalIsOpen = false;
      this.setState({ showModal: false });
    } else {
      this.setState({ showModal: true });
    }
  }
}
