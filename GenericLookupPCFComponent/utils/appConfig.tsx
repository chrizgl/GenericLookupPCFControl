/* eslint-disable @typescript-eslint/no-namespace */
import { IAppConfig } from './appConfig.types';
import { IInputs } from '../generated/ManifestTypes';
export namespace AppConfig {
    export let Configuration: IAppConfig;
    /* istanbul ignore next */
    export async function intialize(context: ComponentFramework.Context<IInputs>) {
        const webresourceName = context.parameters.ConfigFileName.raw
            ? context.parameters.ConfigFileName.raw.toString()
            : 'msp_editable_grid_config';

        let jsonConfig: any = {};
        if (webresourceName) {
            const webResources = await context.webAPI
                .retrieveMultipleRecords('webresource', "?$filter=name eq '" + webresourceName + "'&$top=1")
                .catch((error) => {
                    console.error(
                        'Not able to read editable grid configuration from file: ' + webresourceName + ', Error : ' + error.message,
                    );
                });
            if (webResources && webResources.entities && webResources.entities[0]) {
                const config = webResources.entities[0];
                jsonConfig = JSON.parse(atob(config.content));
            }
        }

        Configuration = {
            Context: context,
            config: jsonConfig,
        };
    }
}
