/* eslint-disable react/jsx-key */
/* eslint-disable @typescript-eslint/no-array-constructor */
/* eslint-disable @typescript-eslint/no-this-alias */
import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { IInputs } from '../generated/ManifestTypes';
import iPropsInput from '../interfaces/iPropsInput';
import iCreateField from '../interfaces/iCreateField';
// import GetSampleConfig from "../sampledata/config";
import iField from '../interfaces/iField';
import { ReactTabulator } from 'react-tabulator';
import { Bars } from 'react-loader-spinner';
import { Callout, StackItem } from '@fluentui/react';
import iView from '../interfaces/iView';
import { useStyles } from './StylesV9';
import {
    makeStyles,
    mergeClasses,
    Button,
    FluentProvider,
    webLightTheme,
    Input,
    InputProps,
    useId,
    shorthands,
    tokens,
    ColorPaletteTokens,
    Overflow,
} from '@fluentui/react-components';
class CalloutControlComponent extends React.PureComponent<iPropsInput> {
    ref: any = null;
    _context: ComponentFramework.Context<IInputs>;
    _tmpField: iCreateField;
    _currentFocus = -1;
    _divSearchForm = 'divSearchForm';

    _SearchText = '';
    _columns: any[] = [];
    _data: any[] = [];
    _recordsThreshHoldLimit: number = 0;

    _originalLookupId: string = '';
    _originalLookupText: string = '';
    _placeHolder: string = 'Search in ';
    _entityId: string = '';
    _entityTypeName: string = '';

    _txtSearchId = 'txtSearch';
    _divLookupId = 'divLookup';
    _txtDummyId = 'txtDummy';
    _ddlView = 'ddlView';
    _divValidations = 'divValidations';
    _linkLookupText = 'linkLookupText';
    _divTextbox = 'divTextbox';
    _entitySymbol = '';

    state = {
        expandValidations: false,
        isLookupOpen: false,
        lookupField: undefined,
        selectedLookupField: undefined,
        data: this._data,
        filterText: '',
        calloutVisible: false,
        lookupText: '',
        lookupId: '',
        selectedView: 0,
        showSpinner: false,
    };
    private _useStyles = useStyles;

    constructor(props: iPropsInput) {
        super(props);
        this._context = props.context;
        this._tmpField = JSON.parse(this._context.parameters.ConfigJSON.raw ?? '');
        this._entityId = (this._context as any)?.page?.entityId ?? '';
        this._entityTypeName = (this._context as any)?.page?.entityTypeName ?? '';
        this._recordsThreshHoldLimit = this._tmpField.recordsThreshHoldLimit ?? 0;

        this._txtSearchId = 'txtSearch' + this._tmpField.name;
        this._divLookupId = 'divLookup' + this._tmpField.name;
        this._txtDummyId = 'txtDummy' + this._tmpField.name;
        this._ddlView = 'ddlView' + this._tmpField.name;
        this._divValidations = 'divValidations' + this._tmpField.name;
        this._linkLookupText = 'linkLookupText' + this._tmpField.name;
        this._divTextbox = 'divTextbox' + this._tmpField.name;
        this._entitySymbol = 'crmSymbolFont entity-symbol ' + this._tmpField.entitySymbol ?? 'Account';

        this.LoadColumns();
        this._tmpField.lookUpCol?.fitlerTextFields?.forEach((tmpField: iField, index) => {
            this._placeHolder = this._placeHolder + tmpField.displayText;
            if (index < (this._tmpField.lookUpCol?.fitlerTextFields?.length ?? 0) - 1) {
                this._placeHolder = this._placeHolder + ', ';
            }
        });
    }

    LoadInitialData = () => {
        const thisRef = this;

        if (this._entityId.length > 0) {
            this._context.webAPI.retrieveRecord(this._entityTypeName, this._entityId, '?$select=_' + this._tmpField.name + '_value').then(
                function success(result) {
                    const tmpLookupId = result['_' + thisRef._tmpField.name + '_value'];
                    const tmpLookupText = result['_' + thisRef._tmpField.name + '_value@OData.Community.Display.V1.FormattedValue'];

                    thisRef.setState({
                        lookupId: tmpLookupId,
                        lookupText: tmpLookupText,
                    });
                },
                function (error) {
                    console.log(error.message);
                },
            );
        }
        this.LoadData(0);
    };
    /* istanbul ignore next */
    OnFilterTextChange = (event: any) => {
        this.setState({ filterText: event.target.value });
    };
    OnLookupUpClick = (event: any, tmpField: iCreateField) => {
        this.OpenLookupDialog(tmpField);
    };
    OpenLookupDialog = (tmpColDef?: iCreateField) => {
        this.setState({
            isLookupOpen: true,
            lookupField: tmpColDef?.lookUpCol,
            selectedLookupField: tmpColDef?.name,
        });
    };
    CloseLookupDialog = (lookupObj: any) => {
        this.setState({
            isLookupOpen: false,
            selectedLookupField: 'unknown',
        });
    };

    GetLookupText = (tmpField: iCreateField) => {
        return '---';
    };

    LoadColumns = () => {
        this._columns.push({
            formatter: (cell: any) => this.SelectFormatter(cell),
            align: 'center',
            headerSort: false,
            width: 80,
            field: 'select',
        });

        this._tmpField.lookUpCol?.fieldsToShow?.forEach((tmp: iField) => {
            const tmpCol = {
                title: tmp.displayText,
                field: tmp.name,
                headerFilter: 'input',
            };
            this._columns.push(tmpCol);
        });
    };
    SelectFormatter = (cell: any) => {
        // eslint-disable-next-line @typescript-eslint/no-this-alias
        const thisRef = this;
        const tmpAnchor = document.createElement('a');
        tmpAnchor.href = '#';
        tmpAnchor.id = 'lnk' + cell.getRow().getPosition(true);
        tmpAnchor.tabIndex = -1;
        tmpAnchor.innerText = 'Select';
        tmpAnchor.className = 'editable_grid_title_link';
        tmpAnchor.setAttribute('aria-label', 'Select Record');
        tmpAnchor.onclick = (e) => {
            e.preventDefault();
            const tmpSelectedItem = cell.getRow().getData()[this._tmpField.lookUpCol?.primaryFeild ?? ''];
            const tmpSelectedItemId = cell.getRow().getData()[this._tmpField.lookUpCol?.primaryKey ?? ''];

            thisRef.setState({
                lookupText: tmpSelectedItem,
                lookupId: tmpSelectedItemId,
            });

            // eslint-disable-next-line @typescript-eslint/no-array-constructor
            const lookupValue = new Array();
            lookupValue[0] = new Object();
            lookupValue[0].id = tmpSelectedItemId;
            lookupValue[0].name = tmpSelectedItem;
            lookupValue[0].entityType = thisRef._tmpField.lookUpCol?.entity;

            // @ts-ignore
            const tmpLookupField = Xrm.Page.getAttribute(thisRef._tmpField.name);
            tmpLookupField.setValue(lookupValue);

            thisRef.SetEditability(false);
            thisRef.SetLookupText();
            thisRef.CloseCallOut();
        };
        return tmpAnchor;
    };
    LoadData = (selectedViewId: number) => {
        // eslint-disable-next-line @typescript-eslint/no-this-alias
        const thisRef = this;
        this._SearchText = (document.getElementById(this._txtSearchId) as HTMLInputElement).value ?? '';
        if (this._tmpField.exterCall) {
            this._context.webAPI
                .retrieveMultipleRecords('webresource', "?$filter=name eq '" + this._tmpField.exterCall.webResource + "'&$top=1")
                .then(
                    function success(result) {
                        if (result.entities.length > 0) {
                            const config = result.entities[0];
                            const copyFunction = new Function('return ' + atob(config.content))();
                            copyFunction(thisRef);
                        }
                    },
                    function (error) {
                        console.log(error);
                    },
                );
        } else {
            if (this._tmpField?.lookUpCol?.views) {
                const tmpView = this._tmpField?.lookUpCol?.views[selectedViewId];
                if (tmpView.fetchXml) this.LoadDataFromFetchXML(selectedViewId);
                else this.LoadDataFromNonFetchXML();
            }
        }
    };

    LoadDataFromFetchXML = (selectedViewId: number) => {
        const thisref = this;
        if (this._tmpField?.lookUpCol?.views) {
            let fetchXml = this._tmpField?.lookUpCol?.views[selectedViewId]?.fetchXml ?? '';
            let tmpConditions = '';
            if (this._SearchText) {
                const tmpSearchVals = this._SearchText.split(',');
                this._tmpField.lookUpCol.fitlerTextFields?.forEach((tmpField: iField, index) => {
                    let tmpSearchVal = '';
                    if (tmpSearchVals.length > index) tmpSearchVal = tmpSearchVals[index]?.trim();
                    if (tmpSearchVal?.trim().length > 0) {
                        tmpConditions =
                            tmpConditions +
                            "<condition attribute='" +
                            tmpField.name +
                            "' operator='like' value='%" +
                            tmpSearchVal +
                            "%' />";
                    }
                });
            }
            fetchXml = fetchXml.replace('<MORECONDITIONS/>', tmpConditions);
            fetchXml = '?fetchXml=' + encodeURIComponent(fetchXml);
            this._context.webAPI.retrieveMultipleRecords(this._tmpField.lookUpCol?.entity ?? '', fetchXml).then(
                function success(result) {
                    thisref.PopulateData(result);
                },
                function (error) {
                    console.log(error.message);
                },
            );
        }
    };
    LoadDataFromNonFetchXML = () => {
        const tmpSearchText = (document.getElementById(this._txtSearchId) as HTMLInputElement).value ?? '';

        let tmpSelect = '?$select=';
        this._tmpField.lookUpCol?.fieldsToShow?.forEach((tmp: iField) => {
            tmpSelect = tmpSelect + tmp.name + ',';
        });
        tmpSelect = tmpSelect + this._tmpField.lookUpCol?.primaryKey;
        let tmpQueryFilter = '';
        const tmpFilters: string[] = [];
        if (tmpSearchText.length > 0) {
            tmpQueryFilter = '&$filter=';
            const tmpSearchVals = tmpSearchText.split(',');
            this._tmpField?.lookUpCol?.fitlerTextFields?.forEach((tmpField: iField, index) => {
                let tmpSearchVal = '';
                if (tmpSearchVals.length > index) tmpSearchVal = tmpSearchVals[index]?.trim();
                if (tmpSearchVal?.trim().length > 0) {
                    tmpFilters.push('contains(' + tmpField.name + ", '" + tmpSearchVal + "')");
                }
            });
        }

        tmpFilters.forEach((tmpstr: string, index) => {
            tmpQueryFilter = tmpQueryFilter + tmpstr;
            if (index < tmpFilters.length - 1) {
                tmpQueryFilter = tmpQueryFilter + ' and ';
            }
        });

        const thisRef = this;

        let tmpOptions = tmpSelect;
        if (tmpQueryFilter) tmpOptions = tmpOptions + tmpQueryFilter;

        this._context.webAPI.retrieveMultipleRecords(this._tmpField.lookUpCol?.entity ?? '', tmpOptions).then(
            function success(result) {
                thisRef.PopulateData(result);
            },
            function (error) {
                console.log(error.message);
            },
        );
    };

    PopulateData = (result: any) => {
        const tmpData = [];
        for (let i = 0; i < result.entities.length; i++) {
            const tmpItem: any = {};
            tmpItem['select'] = 'select';
            tmpItem[this._tmpField.lookUpCol?.primaryKey ?? ''] = result.entities[i][this._tmpField.lookUpCol?.primaryKey ?? ''];
            this._tmpField.lookUpCol?.fieldsToShow?.forEach((tmp: iField) => {
                if (tmp.name) {
                    tmpItem[tmp.name] = result.entities[i][tmp.name];
                }
            });
            tmpData.push(tmpItem);
        }
        this.setState({ data: tmpData, showSpinner: false });
    };

    OpenCallOut = () => {
        this.setState({ calloutVisible: true });
    };
    CloseCallOut = () => {
        this.setState({ calloutVisible: false });
    };
    componentDidMount = () => {
        const thisref = this;
        if (this._tmpField.openSearchPanelOnKeyDown) {
            const txtSearchBox = document.getElementById(this._txtSearchId);

            txtSearchBox?.addEventListener('input', function (e) {
                txtSearchBox?.focus();
                thisref.OpenCallOut();
                thisref.SetFilter();
            });
        }
        this.LoadInitialData();
    };

    componentDidUpdate = () => {
        this.SetLookupText();
    };

    SetFilter = () => {
        const tmpSearchText = (document.getElementById(this._txtSearchId) as HTMLInputElement).value ?? '';

        if (tmpSearchText) {
            const tmpFinalFilters: any = [];
            tmpFinalFilters.push({ field: 'select', type: '=', value: 'select' });

            const tmpFilters: any = [];

            this._tmpField.lookUpCol?.fieldsToShow?.forEach((tmp: iField) => {
                const tmpFilter: any = {};
                tmpFilter['field'] = tmp.name;
                tmpFilter['type'] = 'like';
                tmpFilter['value'] = tmpSearchText;

                tmpFilters.push(tmpFilter);
            });
            tmpFinalFilters.push(tmpFilters);
            this.ref.table?.setFilter(tmpFinalFilters);
        } else {
            this.ref.table?.clearFilter();
        }
    };

    OnSearchClick = () => {
        this.setState({ data: [], showSpinner: true });
        this.OpenCallOut();
        const ddlView = document.getElementById(this._ddlView) as HTMLSelectElement;
        this.LoadData(this.state.selectedView);
    };

    SetEditability = (boolEditable: boolean) => {
        const divLookup = document.getElementById(this._divLookupId);
        const divTextbox = document.getElementById(this._divTextbox);
        const divValidations = document.getElementById(this._divValidations);

        if (boolEditable) {
            divLookup?.style.setProperty('display', 'none');
            divTextbox?.style.setProperty('display', 'inline');
            if (this._tmpField.validation?.type === 'Required') {
                divValidations?.classList.remove('egmtValidationDivHide');
                divValidations?.classList.add('egmtValidationDiv');
            }
        } else {
            divLookup?.style.setProperty('display', 'inline');
            divTextbox?.style.setProperty('display', 'none');
            divValidations?.classList.remove('egmtValidationDiv');
            divValidations?.classList.add('egmtValidationDivHide');
        }
    };

    SetLookupText = () => {
        const linkLookupText = document.getElementById(this._linkLookupText);
        if (linkLookupText) linkLookupText.innerText = this.state.lookupText ?? '';
    };
    OnViewChange = () => {
        this.setState({ data: [], showSpinner: true });
        const ddlView = document.getElementById(this._ddlView) as HTMLSelectElement;
        this.setState({ selectedView: Number.parseInt(ddlView.value) });
        this.LoadData(Number.parseInt(ddlView.value));
    };
    OnNewClick = () => {
        const thisRef = this;
        this._context.navigation
            .openForm({
                entityName: this._tmpField.lookUpCol?.entity ?? '',
                useQuickCreateForm: true,
            })
            .then(
                function (success) {
                    const lookupValue = new Array();
                    lookupValue[0] = new Object();
                    lookupValue[0].id = success.savedEntityReference[0].id;
                    lookupValue[0].name = success.savedEntityReference[0].name;
                    lookupValue[0].entityType = (success.savedEntityReference[0] as any).entityType;

                    // @ts-ignore
                    const tmpLookupField = Xrm.Page.getAttribute(thisRef._tmpField.name);
                    tmpLookupField.setValue(lookupValue);

                    thisRef.setState({
                        lookupText: success.savedEntityReference[0].name,
                        lookupId: success.savedEntityReference[0].id,
                    });

                    thisRef.SetEditability(false);
                    thisRef.SetLookupText();
                },
                function (error) {
                    console.log(error);
                },
            );
    };
    OnTextClick = (e: any) => {
        e.preventDefault();
        if (this._tmpField.lookUpCol?.pageUrl) {
            const tmpLink = this._tmpField.lookUpCol?.pageUrl + this._entityId;
            window.open(tmpLink, '_self');
        }
    };

    public render() {
        const classes = this._useStyles();
        const stackClasses = mergeClasses(classes.stack, classes.stackHorizontal);
        const overflowClass = mergeClasses(classes.overflow, classes.stackitem);
        const inputClass = mergeClasses(classes.input, classes.stackitem);
        const iconClass = mergeClasses(classes.icon, classes.stackitem);
        const options = {
            layoutColumnsOnNewData: true,
            tooltips: true, //show tool tips on cells
            addRowPos: 'top', //when adding a new row, add it to the top of the table
            history: true, //allow undo and redo actions on the table
            resizableRows: true, //allow row order to be changed
            height: 590,
            pagination: 'local',
            paginationSize: 10,
            placeholder: 'No Data Available',
        };
        return (
            <FluentProvider theme={webLightTheme}>
                <div>
                    <div>
                        <div className={classes.stackVertical} id='divCreateDetail'>
                            <div className={classes.stackitem}>Hallo</div>
                        </div>
                    </div>
                </div>
            </FluentProvider>
        );
    }
}
export default CalloutControlComponent;
