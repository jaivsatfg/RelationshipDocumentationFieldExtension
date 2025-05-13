import * as React from "react";
import { CommandBarButton, ContextualMenu, DefaultButton, FontWeights, getFocusStyle, getTheme, IButtonStyles, IconButton, IDragOptions, IIconProps, IStackProps, Label, mergeStyleSets, Modal, ProgressIndicator, Stack, StackItem } from "office-ui-fabric-react";
import { Tree } from 'primereact/tree';
import { DataTable, DataTableSelectionMultipleChangeEvent } from 'primereact/datatable';
import { Column } from 'primereact/column';

import { SPHttpClientResponse } from '@microsoft/sp-http';
import { ISoapTaxonomyResponse } from "./ISoapTaxonomyResponse";
import { AppContext, IAppContext } from "../IContext";
import { DocInfo } from "./IDocInfo";
import { IRelatedDocument } from "./IRelatedDocument";
import 'bootstrap/dist/css/bootstrap.min.css';
import styles from "./RelatedDocumentFieldExtension.module.scss";
import 'primeicons/primeicons.css';
import 'primereact/resources/primereact.css';
import 'primereact/resources/themes/lara-light-indigo/theme.css';
import RelatedDocumentFieldExtension from "./RelatedDocumentFieldExtension";
import { TreeNode } from "primereact/treenode";


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
    selectedTreeKeys: TreeNode[];
    itemsToSelect: DocInfo[];
    selectedItemsToAdd: DocInfo[];
    relatedItems: DocInfo[];
    saving: boolean;
    errorMessage: string;
}


class LinkDocuments extends React.Component<ILinkDocumentsProps, ILinkDocumentsStates> {
    static contextType = AppContext;
    private treeInitialization: boolean = false;
    private modalClosing: boolean = false;
    private savedItems: DocInfo[];
    private isCallFromNotificationsList: boolean;
    private isCallFromDocsTrabajolList: boolean;

    public constructor(props: ILinkDocumentsProps | Readonly<ILinkDocumentsProps>) {
        super(props);

        this.onDismiss = this.onDismiss.bind(this);
        this.state = {
            modalVisible: this.props.isModalOpen,
            modalTreeUpdating: false,
            itemsToSelect: [],
            selectedItemsToAdd: [],
            selectedTreeKeys: [],
            relatedItems: this.props.items,
            saving: false,
            errorMessage: ''
        }
        this.insertSelectedItems = this.insertSelectedItems.bind(this);
        this.itemsSave = this.itemsSave.bind(this);
        this.selectedItemToBeRemoved = this.selectedItemToBeRemoved.bind(this);
        this.removedSelectedItem = this.removedSelectedItem.bind(this);
        this.treeNodeSelect = this.treeNodeSelect.bind(this);
        this.itemsToSelectionChange = this.itemsToSelectionChange.bind(this);
    }

    private selectedItemToBeRemoved(sender: React.MouseEvent<HTMLDivElement>): void {
        const currentIdx: number = this.state.relatedItems.findIndex((doc: DocInfo) => {
            return '' + doc.id === sender.currentTarget.dataset.docid;
        });
        if (currentIdx !== -1) {
            this.setState(prevState => {
                const items = [...prevState.relatedItems];
                const currentItem = items[currentIdx];

                items[currentIdx] = {
                    ...currentItem,
                    selected: !currentItem.selected
                };

                return { relatedItems: items };
            });
        }
    }

    private removedSelectedItem = (): void => {
        this.setState(prevState => ({
            relatedItems: prevState.relatedItems.filter((doc: DocInfo) => !doc.selected)
        }));
    };

    private insertSelectedItems(): void {
        const items: DocInfo[] = this.state.selectedItemsToAdd.filter((d: DocInfo) => {
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
            /* Actualizo los campos LinksDocumentosRelacionados,DocuRelaIds y DocuRelaJson */
            type updateValue = {
                LinksDocumentosRelacionados?: string;
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
                    'LinksDocumentosRelacionados': documentsLinks,
                    'DocuRelaIds': documentsIds,
                    'DocuRelaJson': JSON.stringify(docRelaJson)
                };
            }
            ctx.updateSharePointItem && await ctx.updateSharePointItem(this.isCallFromNotificationsList || this.isCallFromDocsTrabajolList ? ctx.webUrl : ctx.documentsManagerWebUrl, ctx.localListTitle, ctx.localListItemId, objectToUpdate).then(() => {
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
        this.isCallFromDocsTrabajolList = ctx.webUrl !== '' && ctx.localListTitle.toLocaleLowerCase() === 'DocumentosTrabajo' ? true : false;
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
                    modalTreeUpdating: false,
                    selectedTreeKeys: []
                });
            }
        } else {
            if (!this.state.modalTreeUpdating) {
                this.setState({
                    modalTreeUpdating: true
                });
            } else {
                RelatedDocumentFieldExtension.modalIsOpen = true;
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
                    'overflow': 'hidden',
                    'padding-left': '15px'
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


        function handleTreeNodeClick(node: TreeNode): void {
            console.log('Nodo clickeado:', node);
        }


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
                        Assignación de documentos relacionados de {this.props.elementName}
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
                                    <div className="treeDocuments">
                                        <Tree value={this.props.terms}
                                            nodeTemplate={(node, options) => (
                                                <div onClick={() => handleTreeNodeClick(node)}>
                                                    <input
                                                        type="checkbox"
                                                        checked={this.isTreeNodeChecked(node)}
                                                        onChange={(e) => this.treeNodeSelect(e, node)}

                                                        style={{
                                                            'marginRight': '0.5rem',
                                                            'border': '2px solid #d1d5db',
                                                            'borderRadius': '6px',
                                                            'transform': 'scale(1.5)'
                                                        }}
                                                    />
                                                    <span>{node.label}</span>
                                                </div>
                                            )} className="w-full md:w-30rem" />
                                    </div>
                                </StackItem>
                                <StackItem {...stackItemPropsDatatable} data-is-focusable={true}>
                                    <Stack {...stackTitle}>
                                        <StackItem>
                                            <span className={classNames.itemTitle}>Documentos relacionados disponibles</span>
                                        </StackItem>
                                        <StackItem>
                                            <CommandBarButton iconProps={addIcon} text="Agregar" className={styles.btnAction} onClick={this.insertSelectedItems} />
                                        </StackItem>
                                    </Stack>
                                    <DataTable value={this.state.itemsToSelect} paginator rows={7} selectionMode={null}
                                        dataKey="id" onSelectionChange={this.itemsToSelectionChange}
                                        selection={this.state.selectedItemsToAdd}
                                        tableStyle={{ minWidth: '50rem' }}>
                                        <Column selectionMode="multiple" headerStyle={{ width: '3rem' }}></Column>
                                        <Column field="name" header="Nombre"></Column>
                                        <Column field="url" header="Ruta"></Column>
                                    </DataTable>
                                </StackItem>
                            </Stack>
                        </StackItem>
                        <StackItem className={classNames.itemRight} data-is-focusable={true}>
                            <Stack tokens={{ childrenGap: 74 }}>
                                <StackItem>
                                    <Stack {...stackTitle}>
                                        <StackItem>
                                            <span className={classNames.itemTitle}>Documentos seleccionados</span>
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

    private isTreeNodeChecked(node: TreeNode): boolean {
        if (!node.key) return false;
        return this.state.selectedTreeKeys.filter((t) => { return t.key === node.key?.toString() }).length > 0
    }


    private treeNodeSelect = (e: React.ChangeEvent<HTMLInputElement>, node: TreeNode): void => {
        const currentInstance: LinkDocuments = this;
        const isChecked = e.currentTarget.checked;

        if (currentInstance.treeInitialization) {
            return;
        }

        if (isChecked && node.key != undefined) {
            //Busco los documentos
            currentInstance.treeInitialization = true;
            currentInstance.getDocuments(node.key.toString()).then((values: DocInfo[]) => {
                currentInstance.treeInitialization = false;
                currentInstance.setState({
                    itemsToSelect: values
                });
            }).catch(() => {
                currentInstance.treeInitialization = false;
            });
            this.setState({
                selectedTreeKeys: [node],
                selectedItemsToAdd: []
            });
        }
    }
   

    private itemsToSelectionChange = (e: DataTableSelectionMultipleChangeEvent<DocInfo[]>): void => {
        const newSelection = e.value || []; // selección actualizada
        const prevSelection = this.state.selectedItemsToAdd;

        // Detectar si fue una selección o deselección
        const added = newSelection.filter(item => !prevSelection.some(p => p.id === item.id));
        const removed = prevSelection.filter(item => !newSelection.some(n => n.id === item.id));

        // Solo actualizar si hubo cambio
        if (added.length > 0 || removed.length > 0) {
            this.setState({
                selectedItemsToAdd: newSelection
            });
        }
    }



    private getDocuments(keyTerm: string): Promise<DocInfo[]> {
        const ctx: IAppContext = this.context;
        //Buscamos el TemrId de este key
        const node: ISoapTaxonomyResponse | undefined = this.findNodeByKey(this.props.terms, keyTerm);
        return new Promise<DocInfo[]>((resolve, reject) => {
            const re = /\'/gi;
            if (node && node.info && node.info.length > 0) {
                let folderName = (node.info[0].parentLabel.split(';').splice(1).join('/') + '/' + node.text).replace(re, '\'\'');
                if (!folderName.startsWith('/')) {
                    folderName = '/' + folderName;
                }
                let params = {
                    '$select': 'Id,File/Name,File/ServerRelativeUrl,File/LinkingUri',
                    '$filter': 'startswith(FileDirRef,'.concat(`'${ctx.documentsManagerRelativeWebUrl.concat('/', ctx.documentsManagerListTitle, folderName)}')`, ' and FSObjType eq 0'),
                    '$expand': 'File',
                    '$top': '5000'
                };
                if (node.children.length === 0) {
                    params = {
                        '$select': 'Id,File/Name,File/ServerRelativeUrl,File/LinkingUri',
                        '$filter': `FSObjType eq 0 and FileDirRef eq '${ctx.documentsManagerRelativeWebUrl.concat('/', ctx.documentsManagerListTitle, folderName)}'`,
                        '$expand': 'File',
                        '$top': '5000'
                    };
                }


                const currentUrl = ctx.documentsManagerWebUrl.concat(`/_api/web/Lists/GetByTitle('`, encodeURIComponent(ctx.documentsManagerListTitle), `')`,
                    '/items?', new URLSearchParams(params).toString());

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

    private findNodeByKey(tree: ISoapTaxonomyResponse[], keyToFind: string): ISoapTaxonomyResponse | undefined {
        for (const node of tree) {
            if (node.key === keyToFind) {
                return node;
            }
            if (node.children) {
                const found = this.findNodeByKey(node.children, keyToFind);
                if (found) {
                    return found;
                }
            }
        }
        return undefined;
    }

    private async manageRelatedDocumentsList(): Promise<void> {
        const ctx: IAppContext = this.context;
        return new Promise<void>((resolve, reject) => {

            /* Si lo llamo de DocumentosTrabajo no debo crear nada en la lista DocumentsRelacionats
               porque el documento o nueva version del documento no está publicado
               y quien se encarga de completar esta lista a la hora de la publicación es el powerAutomate.
            */
            if (this.isCallFromDocsTrabajolList) {
                resolve();
                return;
            }


            /* La lista DocumentosRelacionados juega dos papeles diferentes a la hora agregar o eliminar items
               Escenario 1: Se relaciona uno o varios documentos a una notificación
               Escenario 2: Se relaciona uno o varios documentos a otro documento

               Es por ello, para el escenario 1 tenemos en cuenta el campo 'IdNotificacion' 
               y en el escenario 2 siempre utilizamos los items con 'IdNotificacion' eq 0
            */

            const listName = 'DocumentosRelacionados';

            const params = {
                '$select': 'ID,IdDocumentoRelacionado,IdDocumento',
                '$filter': ''.concat(`Title eq '${ctx.documentsManagerListTitle}'`,
                    ` and IdBibliotecaDocumentos eq '${ctx.documentsManagerListId}'`,
                    this.isCallFromNotificationsList ? '' : ` and IdDocumento eq ${ctx.documentsManagerListItemId.toString()} `,
                    ` and IdNotificacion eq ${this.isCallFromNotificationsList ? ctx.localListItemId : 0}`),
                '$top': '5000'
            };

            //Busco todos los elementos relacionados al item ID que estoy modificando en la lista DocumentsRelacionats
            const currentUrl = ctx.documentsManagerWebUrl.concat(`/_api/web/Lists/GetByTitle('`, encodeURIComponent(listName), `')`,
                '/items?', new URLSearchParams(params).toString());

            ctx.spHttpClient && ctx.spHttpConfiguration && ctx.spHttpClient.get(currentUrl, ctx.spHttpConfiguration)
                .then((response: SPHttpClientResponse) => {
                    if (response.ok) {
                        type itemValues = { value: IRelatedDocument[]; };
                        response.json().then((r: itemValues) => {
                            type relatedItemState = { idItem: number, idDoc: number, state: string };
                            let idsOnListDocRelacionados: relatedItemState[] = [];
                            idsOnListDocRelacionados = r.value
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
                                const currentRelaDoc: relatedItemState = idsOnListDocRelacionados.filter((d: relatedItemState) => {
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
                                            idsOnListDocRelacionados.filter((r: relatedItemState) => { return r.state === ''; })
                                                .map((r: relatedItemState) => {
                                                    return ctx.deleteSharePointItem && ctx.deleteSharePointItem(ctx.documentsManagerWebUrl, 'DocumentosRelacionados', r.idItem);
                                                })
                                        )
                                    ]);
                            })().then(() => {
                                resolve();
                                return;
                            }).catch((e) => {
                                const errorMessage: string = "Error in Promise.all from DocumentosRelacionados items assignment.";
                                this.onError(errorMessage);
                                reject(errorMessage);
                            });
                        }).catch((e) => {
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

