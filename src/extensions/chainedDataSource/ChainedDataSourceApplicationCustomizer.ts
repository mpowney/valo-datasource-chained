import { override } from '@microsoft/decorators';
import { BaseApplicationCustomizer } from '@microsoft/sp-application-base';
import { DataSourceService } from '@valo/extensibility';

import { ChainedDataSource } from './dataSource/ChainedDataSource';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IChainedDataSourceApplicationCustomizerProperties {
  // This is an example; replace with your own property
  testMessage: string;
}

/** A Custom Action which can be run during execution of a Client Side Application */
export default class ChainedDataSourceApplicationCustomizer extends BaseApplicationCustomizer<IChainedDataSourceApplicationCustomizerProperties> {

    private dataSourceService: DataSourceService = null;

    @override
    public async onInit(): Promise<void> {

        console.log(`Initialized ChainedDataSourceApplicationCustomizer`);

        this.dataSourceService = DataSourceService.getInstance();

        this.dataSourceService.registerDataSource({
            dataSourcePrototype: ChainedDataSource.prototype,
            id: "ChainedDataSource",
            name: "Chained"
        });

        return Promise.resolve();
    }

}
