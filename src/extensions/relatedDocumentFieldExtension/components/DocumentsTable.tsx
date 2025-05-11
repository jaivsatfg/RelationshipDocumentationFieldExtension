import * as React from 'react';
import 'datatables.net-bs5/css/dataTables.bootstrap5.css';
import { DocInfo } from './IDocInfo';
import 'datatables.net-select-dt';
import { AppContext, IAppContext } from '../IContext';
import styles from './RelatedDocumentFieldExtension.module.scss';


const $ = require('jquery');
$.DataTable = require('datatables.net-bs5');



interface IDocumentsTableProps {
    values: DocInfo[];
    handleSelectedItems: (docs: DocInfo[]) => void;
}

class DocumentsTable extends React.Component<IDocumentsTableProps> {
    static contextType = AppContext;
    private dataTable: any;

    private columns: any[] = [
        /*
            {
                defaultContent: '',
                className: 'select-chgulpeckbox',
                orderable: false,
                width: 15,
            },*/
        {
            title: 'Nom',
            width: 120,
            data: 'name'
        },
        {
            title: 'Ruta',
            data: 'url',
            render: (data: string, type: string) => {
                const ctx: IAppContext = this.context;
                const currentData:string = data.toString();
                let docName: string = '/' + currentData.split('/').splice(3).join('/');
                const url = docName;
                const maxLength: number = 80;
                if (docName.length > maxLength) {
                    let splitString = docName.split("");
                    splitString = splitString.splice(splitString.length - maxLength);
                    docName = '...' + splitString.join('');
                }
                return `<a href="${ctx.documentsManagerWebUrl}${encodeURIComponent(url)}" target="_blank">${docName}</a>`;
            }
        },
    ];

    public constructor(props: IDocumentsTableProps | Readonly<IDocumentsTableProps>) {
        super(props);
        this.createDataTable = this.createDataTable.bind(this);
        this.columns[1].render = this.columns[1].render.bind(this);
    }

    componentDidMount(): void {
        this.createDataTable();
    }
    componentDidUpdate(prevProps: Readonly<IDocumentsTableProps>): void {
        this.createDataTable();
    }

    componentWillUnmount(): void {
        $('.data-table-wrapper').find('table').DataTable().destroy(true);
    }

    render(): React.ReactElement<{}> {
        return (
            <div>
                <table style={{width:'100%'}} id="tblDocumentsAvailablesToSelect" className={`table table-striped table-bordered ${styles.defaultColor}`} />
            </div>);
    }

    private createDataTable(): void {
        const thisInstance = this;
        
        if (this.dataTable !== undefined) {
            $('#tblDocumentsAvailablesToSelect').DataTable().destroy();
        }
        
        this.dataTable = $('#tblDocumentsAvailablesToSelect').DataTable({
            dom: '<"' + styles.dtHeader + '"lf><"data-table-wrapper docRelacionatsTable"t><"table__pager"ip>',
            data: this.props.values,
            columns: this.columns,
            pageLength: 5,
            lengthMenu: [[5, 10, 25, 50, 100, -1], [5, 10, 25, 50, 100, "All"]],
            destroy: true,
            ordering: true,
            scrollX: true,
            order: [0],
            responsive: true,
            select: {
                style: 'multi'
            },
            selected: true,
            language: {
                "decimal": "",
                "emptyTable": "No se encuentran documentos en el nodo seleccionado.",
                "info": "Mostrando _START_ de _END_ of _TOTAL_ documentos",
                "infoEmpty": "Mostrando 0 de 0 of 0 documentos",
                "infoFiltered": "(filtado de _MAX_ documentos totales)",
                "infoPostFix": "",
                "thousands": ",",
                "lengthMenu": "Mostrando _MENU_ documentos",
                "loadingRecords": "Cargando...",
                "processing": "Procesando...",
                "search": "Buscar",
                "zeroRecords": "No se han encontrado registros coincidentes",
                "paginate": {
                    "first": "Primero",
                    "last": "Último",
                    "next": "Próximo",
                    "previous": "Anterior"
                },
                "aria": {
                    "sortAscending": ": activate to sort column ascending",
                    "sortDescending": ": activate to sort column descending"
                }
            }
        });
        
        this.dataTable.off('select');
        this.dataTable.on('select', function (e: any, dt: { rows: (arg0: { selected: boolean; }) => { (): any; new(): any; data: { (): any; new(): any; }; }; }, type: string, indexes: any) {
            if (type === 'row') {
                const selectedItems: DocInfo[] = Array.from(dt.rows({ selected: true }).data());
                thisInstance.props.handleSelectedItems(selectedItems);
            }
        });
        $('#tblDocumentsAvailablesToSelect .dataTables_empty').text('No existen documentos en el nodo seleccionado.').css('color', '#333');
    }
}
DocumentsTable.contextType = AppContext;
export default DocumentsTable;