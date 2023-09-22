import { IInputs, IOutputs } from './generated/ManifestTypes';
import DataSetInterfaces = ComponentFramework.PropertyHelper.DataSetApi;
type DataSet = ComponentFramework.PropertyTypes.DataSet;
import iPropsInput from './interfaces/iPropsInput';
import CalloutControlComponent from './components/CalloutControlComponent';
import { createRoot, Root } from 'react-dom/client';
import { createElement } from 'react';

export class GenericLookupPCFComponent implements ComponentFramework.StandardControl<IInputs, IOutputs> {
    private _container: HTMLDivElement;
    private _context: ComponentFramework.Context<IInputs>;
    private _optionSets: any[];
    private _config: any;
    private _root: Root;

    private props: iPropsInput;

    constructor() {}

    public init(
        context: ComponentFramework.Context<IInputs>,
        notifyOutputChanged: () => void,
        state: ComponentFramework.Dictionary,
        container: HTMLDivElement,
    ) {
        this._container = container;
        this._context = context;
        this.props = {
            context: this._context,
            optionSets: this._optionSets,
            gridConfig: this._config,
        };
        this._root = createRoot(container!);
    }

    public async updateView(context: ComponentFramework.Context<IInputs>) {
        this._context = context;
        this.props.context = this._context;
        this._root.render(createElement(CalloutControlComponent, this.props));
    }

    /**
     * It is called by the framework prior to a control receiving new data.
     * @returns an object based on nomenclature defined in manifest, expecting object[s] for property marked as “bound” or “output”
     */
    public getOutputs(): IOutputs {
        return {};
    }

    /**
     * Called when the control is to be removed from the DOM tree. Controls should use this call for cleanup.
     * i.e. cancelling any pending remote calls, removing listeners, etc.
     */
    public destroy(): void {
        // Add code to cleanup control if necessary
    }
}
