import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {sp, Web} from '@pnp/sp-commonjs';
import {
  BaseListViewCommandSet,
  Command,
  IListViewCommandSetListViewUpdatedParameters,
  IListViewCommandSetExecuteEventParameters
} from '@microsoft/sp-listview-extensibility';
import { Dialog } from '@microsoft/sp-dialog';
import { CsvDataService } from './ConvertToCsv';
import * as strings from 'TestCommandSetStrings';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface ITestCommandSetProperties {
  // This is an example; replace with your own properties
  sampleTextOne: string;
  sampleTextTwo: string;
}

const LOG_SOURCE: string = 'TestCommandSet';

export default class TestCommandSet extends BaseListViewCommandSet<ITestCommandSetProperties> {

  @override
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, 'Initialized TestCommandSet');
    return Promise.resolve();
  }

  @override
  public onListViewUpdated(event: IListViewCommandSetListViewUpdatedParameters): void {
    const compareOneCommand: Command = this.tryGetCommand('COMMAND_1');
    if (compareOneCommand) {
      // This command should be hidden unless exactly one row is selected.
      compareOneCommand.visible = event.selectedRows.length === 1;
    }
  }
  getUrl(){
    const str = window.location.href;
    const res = str.substring(str.indexOf("sites") ,str.indexOf("/AllItems")).split("/");
    res.forEach(element => {
      console.log(element)
    });
    return res;
  }
  getFiltersFromUrl(){
    const str = window.location.href;
    const res = str.substring(str.indexOf("useFiltersInViewXml=1&") + 22 ,str.indexOf("FilterOp1=In") +12);
      console.log(res)
    return res;
  }

  async getList() {
    const str = window.location.href;
    const F = this.getUrl();
    return fetch(this.context.pageContext.web.absoluteUrl +
      `/_api/web/GetListUsingPath(DecodedUrl=@a1)/RenderListDataAsStream?@a1=%27%2F${F[0]}%2F${F[1]}%2F${F[2]}%2F${F[3]}%27&View=${await sp.web.lists.getByTitle(this.context.pageContext.list.title).defaultView.get().then(res => {return res.Id})}&TryNewExperienceSingle=TRUE&&${this.getFiltersFromUrl()}`,
      {method: 'post'})
    .then(res =>  {return res.json()})
  }

  async convertToCsv(){
    const json: object[] = await this.getList().then(res => {return res.Row});
    console.log(json[0]);
    console.log(json);
    try{
    CsvDataService.exportToCsv(this.context.pageContext.list.title,json);
    }
    catch(error){

      console.log("error: "+error);
    }
  }
  @override
  public onExecute(event: IListViewCommandSetExecuteEventParameters): void {

    switch (event.itemId) {
      case 'COMMAND_1':
        Dialog.alert(`${this.properties.sampleTextOne}`);

        break;
      case 'COMMAND_2':
        this.getUrl();
        this.getFiltersFromUrl();
        this.convertToCsv();
      default:
        throw new Error('Unknown command');
    }
  }
}


