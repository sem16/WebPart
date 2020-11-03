import { ConvertToXlsx } from './ConvertToXlsx';
import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseListViewCommandSet,
  Command,
  IListViewCommandSetListViewUpdatedParameters,
  IListViewCommandSetExecuteEventParameters
} from '@microsoft/sp-listview-extensibility';
import { Dialog } from '@microsoft/sp-dialog';
import {sp} from '@pnp/pnpjs';
import * as strings from 'HelloWorldCommandSetStrings';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IHelloWorldCommandSetProperties {
  // This is an example; replace with your own properties
  sampleTextOne: string;
  sampleTextTwo: string;
}

const LOG_SOURCE: string = 'HelloWorldCommandSet';

export default class HelloWorldCommandSet extends BaseListViewCommandSet<IHelloWorldCommandSetProperties> {

  @override
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, 'Initialized HelloWorldCommandSet');
    return Promise.resolve();
  }
  data: {}[];
  @override
  public onListViewUpdated(event: IListViewCommandSetListViewUpdatedParameters): void {
    this.data = [];
    console.log(event.selectedRows);
    console.log(this.context.pageContext.list.serverRelativeUrl);
    event.selectedRows.forEach((row,i) => {
      let values: any = {};
      sp.web.lists.get
      row.fields.forEach((field) => {
        let keyName = field.displayName;
        values[keyName]  = row.getValue(field)
      });
      try{
      values['Nome società'] = row.getValueByName('Nome_societa_quick');
      }
      catch{}
      this.data[i] = values;
    });
    console.log(this.data);
  }


  @override
  public onExecute(event: IListViewCommandSetExecuteEventParameters): void {
    switch (event.itemId) {
      case 'COMMAND_2':
        ConvertToXlsx.convertToXslx(this.data);
        break;
      default:
        throw new Error('Unknown command');
    }
  }
}
