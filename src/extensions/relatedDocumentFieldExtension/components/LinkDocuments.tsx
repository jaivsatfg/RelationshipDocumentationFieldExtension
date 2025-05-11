import * as React from "react";
import { CommandBarButton, ContextualMenu, DefaultButton, FontWeights, getFocusStyle, getTheme, IButtonStyles, IconButton, IDragOptions, IIconProps, IStackProps, Label, mergeStyleSets, Modal, ProgressIndicator, Stack, StackItem } from "office-ui-fabric-react";
import 'bootstrap/dist/css/bootstrap.min.css';
import "gijgo";
import "gijgo/css/gijgo.min.css";
import * as jQuery from "jquery";
import 'datatables.net'


import { SPHttpClientResponse } from '@microsoft/sp-http';
import { ISoapTaxonomyResponse } from "./ISoapTaxonomyResponse";
import { AppContext, IAppContext } from "../IContext";
import { DocInfo } from "./IDocInfo";
import DocumentsTable from "./DocumentsTable";
import { DomSelector } from "datatables.net-bs5";
import { IRelatedDocument } from "./IRelatedDocument";
import styles from "./RelatedDocumentFieldExtension.module.scss";
import RelatedDocumentFieldExtension from "./RelatedDocumentFieldExtension";

export interface ILinkDocumentsProps {
    elementName: string;
    items: DocInfo[];
    isModalOpen: boolean;
    terms: ISoapTaxonomyResponse[];
    closeModal: (savedItems: DocInfo[]) => void;
}
export interface ILinkDocumentsStates {
    modalVisible: boolean;
    modalTreeUpdating: boolean;
    itemsToSelect: DocInfo[];
    relatedItems: DocInfo[];
    saving: boolean;
    errorMessage: string;
}

class LinkDocuments extends React.Component<ILinkDocumentsProps, ILinkDocumentsStates> {
    static contextType = AppContext;
    private treeInitialization: boolean = false;
    private modalClosing: boolean = false;
    private tree: Types.Tree;
    private selectedItems: DocInfo[];
    private savedItems: DocInfo[];
    private prevSelectedNode: DomSelector;
    private isCallFromNotificationsList: boolean;
    private isCallFromDocsTreballList: boolean;

    public constructor(props: ILinkDocumentsProps | Readonly<ILinkDocumentsProps>) {
        super(props);
        this.onDismiss = this.onDismiss.bind(this);
        this.state = {
            modalVisible: this.props.isModalOpen,
            modalTreeUpdating: false,
            itemsToSelect: [],
            relatedItems: this.props.items,
            saving: false,
            errorMessage: ''
        }
        this.handleSelectedDocuments = this.handleSelectedDocuments.bind(this);
        this.insertSelectedItems = this.insertSelectedItems.bind(this);
        this.itemsSave = this.itemsSave.bind(this);
        this.selectedItemToBeRemoved = this.selectedItemToBeRemoved.bind(this);
        this.removedSelectedItem = this.removedSelectedItem.bind(this);
    }

    private handleSelectedDocuments(docs: DocInfo[]): void {
        const ctx: IAppContext = this.context;
        this.selectedItems = docs.map((d: DocInfo) => {
            let currentUrl: string = d.url;
            if (currentUrl.indexOf(ctx.documentsManagerRelativeWebUrl) != -1) {
                currentUrl = '/' + d.url.toString().split('/').splice(3).join('/');
            }
            return Object.assign(d, { url: currentUrl });
        });
    }

    private selectedItemToBeRemoved(sender: React.MouseEvent<HTMLDivElement>): void {
        const currentIdx: number = this.state.relatedItems.findIndex((doc: DocInfo) => {
            return '' + doc.id === sender.currentTarget.dataset.docid;
        });
        if (currentIdx !== -1) {
            const items = [...this.state.relatedItems];
            items[currentIdx].selected = !items[currentIdx].selected;
            this.setState({ relatedItems: items });
        }
    }
    private removedSelectedItem = (): void => {
        this.setState({
            relatedItems: this.state.relatedItems.filter((doc: DocInfo) => {
                return !doc.selected;
            }),
        });
    }

    private insertSelectedItems(): void {
        const items: DocInfo[] = this.selectedItems.filter((d: DocInfo) => {
            return this.state.relatedItems.findIndex((doc: DocInfo) => {
                return doc.id === d.id;
            }) === -1;
        });
        if (items.length > 0) {
            this.setState({ relatedItems: this.state.relatedItems.concat(items) });
        }
    }

    private itemsSave(): void {
        this.setState({
            saving: true,
            errorMessage: ''
        });

        const ctx: IAppContext = this.context;
        const documentsIds: string = this.state.relatedItems.map((d: DocInfo) => { return d.id }).join(',');
        const documentsLinks: string = this.state.relatedItems.map((d: DocInfo) => {
            let currentUrl = d.url;
            if (ctx.documentsManagerRelativeWebUrl
                && currentUrl.toLocaleLowerCase().startsWith("/")
                && !currentUrl.toLocaleLowerCase().startsWith(ctx.documentsManagerRelativeWebUrl.toLowerCase())) {
                currentUrl = ctx.documentsManagerRelativeWebUrl.concat(currentUrl);
            }
            return `<a href="${encodeURI(currentUrl)}" target="_blank" data-interception="off">${d.name}</a>`;
        }).join('</br>');
        type relaJson = {
            id: number,
            name: string,
            url: string
        };
        const docRelaJson: relaJson[] = this.state.relatedItems.map((d: DocInfo) => {
            return { id: d.id, url: d.url, name: d.name };
        });
        this.manageRelatedDocumentsList().then(async () => {
            /* Actualizo los campos LinksDocumentsRelacionats,DocuRelaIds y DocuRelaJson */
            type updateValue = {
                LinksDocumentsRelacionats?: string;
                DocuRelaIds: string;
                DocuRelaJson: string;
            }
            let objectToUpdate: updateValue;
            if (this.isCallFromNotificationsList) {
                objectToUpdate = {
                    'DocuRelaIds': documentsIds,
                    'DocuRelaJson': JSON.stringify(docRelaJson)
                };
            }
            else {
                objectToUpdate = {
                    'LinksDocumentsRelacionats': documentsLinks,
                    'DocuRelaIds': documentsIds,
                    'DocuRelaJson': JSON.stringify(docRelaJson)
                };
            }
            ctx.updateSharePointItem && await ctx.updateSharePointItem(this.isCallFromNotificationsList || this.isCallFromDocsTreballList ? ctx.webUrl : ctx.documentsManagerWebUrl, ctx.localListTitle, ctx.localListItemId, objectToUpdate).then(() => {
                this.savedItems = Array.from(this.state.relatedItems);
                this.onDismiss();
            }).catch((error: TypeError) => {
                const errorMessage = "Error in updateSharePointItem assign Related Document to the library documents."
                console.error(errorMessage);
                this.onError(error);
            });
        }).catch((error: string) => {
            this.onError(error);
        });
    }

    public componentDidMount(): void {
        const ctx: IAppContext = this.context;
        this.isCallFromNotificationsList = ctx.webUrl !== '' && ctx.localListTitle.toLocaleLowerCase() === 'notificaciones' ? true : false;
        this.isCallFromDocsTreballList = ctx.webUrl !== '' && ctx.localListTitle.toLocaleLowerCase() === 'DocumentosTrabajo' ? true : false;
        this.savedItems = this.props.items;
        if (this.props.isModalOpen) {
            this.setState({
                modalTreeUpdating: true
            });
        }
    }

    public componentDidUpdate(prevProps: Readonly<ILinkDocumentsProps>, prevState: Readonly<ILinkDocumentsStates>): void {
        if (this.modalClosing) {
            this.modalClosing = false;
            return;
        }
        if (this.state.saving) {
            return;
        }

        if (!this.state.modalVisible && !prevState.modalVisible) {
            if (this.state.modalTreeUpdating || this.state.modalVisible !== this.props.isModalOpen) {
                this.setState({
                    modalVisible: this.props.isModalOpen,
                    modalTreeUpdating: false
                });
            }
        } else {
            if (!this.state.modalTreeUpdating) {
                this.setState({
                    modalTreeUpdating: true
                });
            } else {
                this.createTree();
            }
        }
    }

    componentWillUnmount(): void {
        console.log('componentWillUnmount');
    }

    public onDismiss = (): void => {
        this.modalClosing = true;
        this.setState({ saving: false, modalVisible: false });
        this.props.closeModal(this.savedItems);
    }

    public render(): React.ReactElement<{}> {
        const cancelIcon: IIconProps = { iconName: 'Cancel' };
        const attachIcon: IIconProps = { iconName: 'Attach' };
        const theme = getTheme();
        const contentStyles = mergeStyleSets({
            container: {
                display: 'flex',
                flexFlow: 'column nowrap',
                alignItems: 'stretch',
                borderRadius: '10px',
                width: '95%'
            },
            header: [
                theme.fonts.xLargePlus,
                {
                    flex: '1 1 auto',
                    borderBottom: `1px solid #dee2e6`,
                    color: theme.palette.neutralPrimary,
                    display: 'flex',
                    justifyContent: 'space-between',
                    alignItems: 'center',
                    fontWeight: FontWeights.semibold,
                    fontSize: '18px',
                    padding: '16px',
                },
            ],
            heading: {
                color: '#005257',
                fontWeight: FontWeights.bold,
                fontSize: 'inherit',
                margin: '0',
            },
            body: {
                flex: '4 4 auto',
                background: '#ffffff',
                padding: '25px',
                overflowY: 'hidden',
                marginBottom: 0,
                width: '100%',
                minHeight: 300
            },
            footer: {
                display: 'flex',
                borderTop: '1px solid #dee2e6',
                padding: '16px',
                alignItems: 'flex-end'
            },
        });

        const iconButtonStyles: Partial<IButtonStyles> = {
            root: {
                color: theme.palette.neutralPrimary,
                marginLeft: 'auto',
                marginTop: '4px',
                marginRight: '2px',
            },
            rootHovered: {
                color: theme.palette.neutralDark,
            },
        };
        const iconButtonStylesAttach: Partial<IButtonStyles> = {
            root: {
                color: theme.palette.neutralPrimary,
                margin: '0',
                pointerEvents: 'none',
                span: {
                    alignItems: 'center',
                },
            },

        };
        const { palette, fonts } = theme;
        const classNames = mergeStyleSets({
            itemLeft: [
                getFocusStyle(theme, { inset: -1 }),
                {

                    padding: 0,
                    width: '60%',
                },
            ],
            itemRight: [
                getFocusStyle(theme, { inset: -1 }),
                {
                    padding: 0,
                    width: '40%',
                },
            ],
            itemTitle: {
                'font-size': '16px',
                'font-weight': 'bold',
                'color': '#005257'
            },
            itemContent: {
                display: 'flex',
                padding: "10px",
                cursor: 'pointer',
                selectors: {
                    '&:hover': { background: palette.neutralLight },
                },
            },
            itemName:
            {
                whiteSpace: 'nowrap',
                overflow: 'hidden',
                textOverflow: 'ellipsis',
            },
            itemText:
            {

            },
            itemIndex: {
                fontSize: fonts.small.fontSize,
                color: palette.neutralTertiary,
                marginBottom: 5,
            },
            panelfooterContent: {
                'width': '100%',
                padding: 15,
            },
        });


        const stackProps: Partial<IStackProps> = {
            horizontal: true,
            tokens: { childrenGap: 20 },
            styles: {
                root: {
                    'width': '100%'
                }
            }
        };

        const stackItemPropsTree: Partial<IStackProps> = {
            styles: {
                root: {
                    'min-width': '270px'
                }
            },
        };

        const stackItemPropsDatatable: Partial<IStackProps> = {
            styles: {
                root: {
                    'width': 'calc(100% - 270px)',
                    'overflow': 'hidden'
                }
            },
        };

        const stackTitle: Partial<IStackProps> = {
            horizontal: true,
            horizontalAlign: 'space-between',
            styles: {
                root: {
                    'padding-bottom': '15px'
                }
            }
        };

        const stackLeft: Partial<IStackProps> = {
            horizontal: true,
            styles: {
                root: {
                    'width': '100%'
                }
            }
        };
        const dragOptions: IDragOptions = {
            moveMenuItemText: 'Move',
            closeMenuItemText: 'Close',
            menu: ContextualMenu,
            keepInBounds: false,
            dragHandleSelector: '.ms-Modal-scrollableContent > div:first-child',
        };

        const titleId = 'modalTreeDocRela';
        const addIcon: IIconProps = { iconName: 'IncreaseIndentArrow' };
        const removeIcon: IIconProps = { iconName: 'DecreaseIndentArrow' };
        return (
            <Modal
                isModeless={true}
                titleAriaId={titleId}
                isOpen={this.state.modalVisible}
                onDismiss={this.onDismiss}
                isBlocking={false}

                containerClassName={contentStyles.container}
                dragOptions={dragOptions}
            >
                <div className={contentStyles.header}>
                    <h2 className={contentStyles.heading} id={titleId}>
                        Assignació de documents relacionats de {this.props.elementName}
                    </h2>
                    <IconButton
                        styles={iconButtonStyles}
                        iconProps={cancelIcon}
                        ariaLabel="Close popup modal"
                        onClick={this.onDismiss}
                    />
                </div>
                <Stack className={contentStyles.body}>
                    <Stack {...stackProps}>
                        <StackItem className={classNames.itemLeft}>
                            <Stack {...stackLeft}>
                                <StackItem {...stackItemPropsTree} >
                                    <div className="treeDocuments" />
                                </StackItem>
                                <StackItem {...stackItemPropsDatatable} data-is-focusable={true}>
                                    <Stack {...stackTitle}>
                                        <StackItem>
                                            <span className={classNames.itemTitle}>Documents relacionats disponibles</span>
                                        </StackItem>
                                        <StackItem>
                                            <CommandBarButton iconProps={addIcon} text="Agregar" className={styles.btnAction} onClick={this.insertSelectedItems} />
                                        </StackItem>
                                    </Stack>
                                    <DocumentsTable values={this.state.itemsToSelect} handleSelectedItems={this.handleSelectedDocuments} />
                                </StackItem>
                            </Stack>
                        </StackItem>
                        <StackItem className={classNames.itemRight} data-is-focusable={true}>
                            <Stack tokens={{ childrenGap: 74 }}>
                                <StackItem>
                                    <Stack {...stackTitle}>
                                        <StackItem>
                                            <span className={classNames.itemTitle}>Documents seleccionats</span>
                                        </StackItem>
                                        <StackItem>
                                            <CommandBarButton iconProps={removeIcon} text="Quitar" className={styles.btnAction} onClick={this.removedSelectedItem} />
                                        </StackItem>
                                    </Stack>
                                </StackItem>
                                <StackItem>
                                    {
                                        this.state.relatedItems.map((doc: DocInfo, index: number) => {
                                            return (
                                                <div key={index} className={`${classNames.itemContent}  ${doc.selected ? styles.rowActive : ''}`} data-docid={doc.id} onClick={this.selectedItemToBeRemoved}>
                                                    <IconButton
                                                        styles={iconButtonStylesAttach}
                                                        iconProps={attachIcon}
                                                        ariaLabel="Atacch"
                                                    />
                                                    <div className={classNames.itemText}>
                                                        <div className={classNames.itemName}>{doc.name}</div>
                                                        <div className={classNames.itemIndex}>{doc.url}</div>
                                                    </div>

                                                </div>
                                            )
                                        })
                                    }
                                </StackItem>
                            </Stack>
                        </StackItem>
                    </Stack>
                </Stack>
                <Stack horizontal horizontalAlign="space-between" className={contentStyles.footer}>
                    {this.state.saving ?
                        <ProgressIndicator className={styles.progressBar} label="Asignando las relaciones" description="Espere" />
                        : <></>
                    }
                    <Label className={styles.errorLabel}>{this.state.errorMessage}</Label>
                    <DefaultButton text="Guardar" className={styles.btnActionFooter} onClick={this.itemsSave} />
                </Stack>
            </Modal>
        );
    }

    private createTree(): void {
        RelatedDocumentFieldExtension.modalIsOpen = true;
        const currentInstance: LinkDocuments = this;
        const $ = require('jquery');
        $.DataTable = require('datatables.net');
        this.tree = $('.treeDocuments').tree({
            primaryKey: 'id',
            uiLibrary: 'bootstrap5',
            dataSource: this.props.terms,
            cascadeCheck: false,
            checkboxes: true,
            checkedField: 'selected',
            textField: 'text'
        });
        this.tree.off('checkboxChange');
        this.tree.on('checkboxChange', (e, node, record, state) => {
            if (state !== 'checked' || currentInstance.treeInitialization) {
                return;
            }

            $(this.prevSelectedNode).find('input[type="checkbox"]:first').prop('checked', false);
            this.prevSelectedNode = node;
            currentInstance.treeInitialization = true;


            currentInstance.getDocuments(record).then((values: DocInfo[]) => {
                currentInstance.treeInitialization = false;
                currentInstance.setState({
                    itemsToSelect: values
                });
            }).catch(() => {
                currentInstance.treeInitialization = false;
            });
        });
    }

    private getDocuments(node: ISoapTaxonomyResponse): Promise<DocInfo[]> {
        const ctx: IAppContext = this.context;
        return new Promise<DocInfo[]>((resolve, reject) => {
            const re = /\'/gi;
            if (node && node.info && node.info.length > 0) {
                const folderName = (node.info[0].parentLabel.split(';').splice(1).join('/') + '/' + node.text).replace(re, '\'\'');
                const params = {
                    '$select': 'Id,File/Name,File/ServerRelativeUrl,File/LinkingUri',
                    '$filter': ''.concat(`FSObjType eq 0 and FileDirRef eq '${ctx.documentsManagerRelativeWebUrl.concat('/', ctx.documentsManagerListTitle, '/Documents privats/', folderName)}'`,
                        ' or ', `FSObjType eq 0 and FileDirRef eq '${ctx.documentsManagerRelativeWebUrl.concat('/', ctx.documentsManagerListTitle, '/Documents publics/', folderName)}'`),
                    '$expand': 'File',
                    '$top': 5000
                };

                const currentUrl = ctx.documentsManagerWebUrl.concat(`/_api/web/Lists/GetByTitle('`, encodeURIComponent(ctx.documentsManagerListTitle), `')`,
                    '/items?', jQuery.param(params));

                ctx.spHttpClient && ctx.spHttpConfiguration && ctx.spHttpClient.get(currentUrl, ctx.spHttpConfiguration)
                    .then((response: SPHttpClientResponse) => {
                        if (response.ok) {
                            type itemValue = { ID: number; File: { Name: string; ServerRelativeUrl: string; LinkingUri: string; } };
                            type itemValues = { value: itemValue[]; };
                            response.json().then((r: itemValues) => {
                                const values: DocInfo[] = [];
                                r.value.map((it: itemValue) => {
                                    let fileUrl = it.File.LinkingUri ? '/' + it.File.LinkingUri.split('/').splice(3).join('/') : '';
                                    if (!fileUrl && it.File.ServerRelativeUrl) {
                                        const filePath: string = encodeURIComponent(it.File.ServerRelativeUrl.split('/').splice(0, it.File.ServerRelativeUrl.split('/').length - 1).join('/'));
                                        fileUrl = ctx.documentsManagerRelativeWebUrl.concat('/', ctx.documentsManagerListTitle, '/Forms/AllItems.aspx?id=', it.File.ServerRelativeUrl, '&parent=', filePath);
                                        if (ctx.documentsManagerRelativeWebUrl && fileUrl.toLocaleLowerCase().indexOf(ctx.documentsManagerRelativeWebUrl.toLocaleLowerCase()) == -1) {
                                            fileUrl = ctx.documentsManagerRelativeWebUrl.concat(fileUrl);
                                        }
                                    }
                                    values.push({
                                        id: it.ID,
                                        name: it.File.Name,
                                        url: fileUrl,
                                        selected: false,
                                    });
                                });
                                resolve(values);
                            }).catch((e: any) => {
                                console.error(e);
                                reject();
                            });
                        } else {
                            response.json().then((r: TypeError): void => {
                                reject('ERROR response: Status: '.concat(
                                    response.status.toString(),
                                    response.statusText ? ' - '.concat(response.statusText) : '',
                                    r.message ? ' details: '.concat(r.message) : ''));
                            }).catch((e) => {
                                console.error(e);
                                reject();
                            });
                        }
                    })
                    .catch((e) => {
                        console.error(e);
                        reject();
                    });
            } else {
                reject();
            }
        });
    }

    private async manageRelatedDocumentsList(): Promise<void> {
        const ctx: IAppContext = this.context;
        return new Promise<void>((resolve, reject) => {

            /* Si lo llamo de DocsTreball no debo crear nada en la lista DocumentsRelacionats
               porque el documento o nueva version del documento no está publicado
               y quien se encarga de completar esta lista a la hora de la publicación es el powerAutomate.
            */
            if (this.isCallFromDocsTreballList) {
                resolve();
                return;
            }


            /* La lista DocumentsRelacionats juega dos papeles diferentes a la hora agregar o eliminar items
               Escenario 1: Se relaciona uno o varios documentos a una notificación
               Escenario 2: Se relaciona uno o varios documentos a otro documento

               Es por ello, para el escenario 1 tenemos en cuenta el campo 'IdNotificacio' 
               y en el escenario 2 siempre utilizamos los items con 'IdNotificacio' eq 0
            */

            const listName = 'DocumentsRelacionats';

            const params = {
                '$select': 'ID,IdDocumentRelacionat,IdDocument',
                '$filter': ''.concat(`Title eq '${ctx.documentsManagerListTitle}'`,
                    ` and IdBibliotecaDocuments eq '${ctx.documentsManagerListId}'`,
                    this.isCallFromNotificationsList ? '' : ` and IdDocument eq ${ctx.documentsManagerListItemId.toString()} `,
                    ` and IdNotificacio eq ${this.isCallFromNotificationsList ? ctx.localListItemId : 0}`),
                '$top': 5000
            };

            //Busco todos los elementos relacionados al item ID que estoy modificando en la lista DocumentsRelacionats
            const currentUrl = ctx.documentsManagerWebUrl.concat(`/_api/web/Lists/GetByTitle('`, encodeURIComponent(listName), `')`,
                '/items?', jQuery.param(params));

            ctx.spHttpClient && ctx.spHttpConfiguration && ctx.spHttpClient.get(currentUrl, ctx.spHttpConfiguration)
                .then((response: SPHttpClientResponse) => {
                    if (response.ok) {
                        type itemValues = { value: IRelatedDocument[]; };
                        response.json().then((r: itemValues) => {
                            type relatedItemState = { idItem: number, idDoc: number, state: string };
                            let idsOnListDocRelacionats: relatedItemState[] = [];
                            idsOnListDocRelacionats = r.value
                                .filter((it: IRelatedDocument) => {
                                    return typeof it.IdDocumentoRelacionado !== 'undefined' && it.IdDocumentoRelacionado !== null;
                                }).map((it: IRelatedDocument): relatedItemState => {
                                    return {
                                        idItem: it.ID,
                                        idDoc: it.IdDocumentoRelacionado !== null ? parseInt(it.IdDocumentoRelacionado.toString()) : 0,
                                        state: ''
                                    };
                                });

                            //Recorro la nueva asignación y doy de alta o elimino según corresponda
                            const itemsToCreate: IRelatedDocument[] = [];
                            this.state.relatedItems.forEach((docToMakeRelations: DocInfo) => {
                                //Si el item no existe en la lista DocumentsRelacionats lo agrego                                
                                const currentRelaDoc: relatedItemState = idsOnListDocRelacionats.filter((d: relatedItemState) => {
                                    return d.idDoc === docToMakeRelations.id;
                                })[0];
                                if (!currentRelaDoc) {
                                    const newRelationItem: IRelatedDocument = {
                                        Title: ctx.documentsManagerListTitle,
                                        IdBibliotecaDocumentos: ctx.documentsManagerListId,
                                        ID: 0,
                                        IdDocumento: 0,
                                        IdDocumentoRelacionado: 0,
                                        IdNotificacion: 0
                                    };
                                    //Estoy relacionando un documento a una notificacion (Escenario 1)
                                    if (this.isCallFromNotificationsList) {
                                        newRelationItem.IdDocumento = 0;
                                        newRelationItem.IdDocumentoRelacionado = docToMakeRelations.id;
                                        newRelationItem.IdNotificacion = ctx.localListItemId;
                                    } else { //Estoy relacionando un documento con otro (Escenario 2)
                                        newRelationItem.IdDocumento = ctx.documentsManagerListItemId;
                                        newRelationItem.IdDocumentoRelacionado = docToMakeRelations.id;
                                        newRelationItem.IdNotificacion = 0;
                                    }
                                    itemsToCreate.push(newRelationItem);
                                } else { //Si existe lo marco que ya está en la lista
                                    currentRelaDoc.state = 'exists';
                                }
                            });

                            (async function () {
                                await Promise.all(
                                    [
                                        //Creo items
                                        await Promise.all(
                                            itemsToCreate.map((newRelationItem: IRelatedDocument) => {
                                                return ctx.insertSharePointItem && ctx.insertSharePointItem(ctx.documentsManagerWebUrl, 'DocumentosRelacionados', newRelationItem);
                                            })
                                        ),
                                        await Promise.all(
                                            //Elimino los que sobran
                                            idsOnListDocRelacionats.filter((r: relatedItemState) => { return r.state === ''; })
                                                .map((r: relatedItemState) => {
                                                    return ctx.deleteSharePointItem && ctx.deleteSharePointItem(ctx.documentsManagerWebUrl, 'DocumentosRelacionados', r.idItem);
                                                })
                                        )
                                    ]);
                            })().then(() => {
                                resolve();
                                return;
                            }).catch(() => {
                                const errorMessage: string = "Error in Promise.all from DocumentosRelacionados items assingment.";
                                this.onError(errorMessage);
                                reject(errorMessage);
                            });
                        }).catch((e: string) => {
                            this.onError(e);
                        });
                    } else {
                        response.json().then((r: TypeError) => {
                            const errorMessage: string = 'ERROR response: Status: '.concat(
                                response.status.toString(),
                                response.statusText ? ' - '.concat(response.statusText) : '',
                                r.message ? ' details: '.concat(r.message) : '');
                            this.onError(errorMessage);
                            reject(errorMessage);
                        }).catch((e: string) => {
                            this.onError(e);
                        });
                    }
                })
                .catch((e: SPHttpClientResponse) => {
                    this.onHttpResponseError(e);
                    reject();
                });
        });
    }

    private onError(error: (TypeError | string)): void {
        let errorMessage: string = '';
        if (typeof error === 'object' && error.name === 'TypeError') {
            errorMessage = error.message.concat(" --> stack" + error.stack,
                //error.error ? ' details: '.concat(error.error.code + " -- " + error.error.message) : ''
            );
        } else {
            errorMessage = error.toString();
        }
        console.log(errorMessage);
        this.setState({
            saving: false,
            errorMessage: 'ERROR: '.concat(errorMessage)
        });
    }
    private onHttpResponseError(response: SPHttpClientResponse): void {
        const errorMessage: string = 'ERROR response: Status: '.concat(
            response.status.toString(),
            response.statusText ? ' - '.concat(response.statusText) : '',
        );
        console.log(errorMessage);
        this.setState({ errorMessage: errorMessage });
    }
}

LinkDocuments.contextType = AppContext;
export default LinkDocuments

