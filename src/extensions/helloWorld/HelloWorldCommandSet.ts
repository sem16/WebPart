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
import * as React from 'react';
import {Message} from './msg';
import * as ReactDOM from 'react-dom';

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
  data: {}[];
  dialogPlaceHolder = document.body.appendChild(document.createElement("div"));

  @override
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, 'Initialized HelloWorldCommandSet');
    sp.setup({pageContext: {web: {absoluteUrl: this.context.pageContext.web.absoluteUrl} }})
    return Promise.resolve();
    ;
  }

  @override
  public onListViewUpdated(event: IListViewCommandSetListViewUpdatedParameters): void {
    this.data = [];
    console.log(event.selectedRows);
    console.log(this.context);
    console.log(sp.web.getParentWeb());
    console.log(this.context.dynamicDataProvider.getAvailableSources().map(el => el.metadata.instanceId))
    event.selectedRows.forEach((row,i) => {
      let values: any = {};

      try{
      values['Nome società'] = row.getValueByName('Nome_societa_quick');
      }

      catch{}
      row.fields.forEach((field) => {
        let keyName = field.displayName;
        values[keyName]  = row.getValue(field)
      });
      if(values['Nome società'] === undefined){
        delete values['Nome società'];
      }
      try{
        delete values['Attachments'];
        delete values['Allegati']
      }catch{}
      this.data[i] = values;
    });
    console.log(this.data);
  }


  @override
  public onExecute(event: IListViewCommandSetExecuteEventParameters): void {
    switch (event.itemId) {
      case 'COMMAND_2':
        let url= this.context.pageContext.site.serverRequestPath;
        let arrOfStr:string[] = url.split("/");
        let listName: String;
        console.log(arrOfStr);
        for(let I=0; I<arrOfStr.length;I++){​​​​
          if(arrOfStr[I]==="Lists" || arrOfStr[I]==="SitePages"){​​​​
          console.log(arrOfStr[I+1])
           listName=arrOfStr[I+1];
           listName = listName.replace(".aspx","");
          }​​​​
        }​​​​

        if(event.selectedRows.length === 0){
          try{
            ReactDOM.unmountComponentAtNode(this.dialogPlaceHolder);
          }catch{}

          const element: React.ReactElement<{}> = React.createElement(
            Message,{
              show: true
            }
          );

          ReactDOM.render(element,this.dialogPlaceHolder);
        }else{
          try{
            ReactDOM.unmountComponentAtNode(this.dialogPlaceHolder);
          }catch{}
        }
        // const listName = this.context.dynamicDataProvider.getAvailableSources()[1].metadata.title;
        // sp.web.lists.getByTitle(listName).get().then(res => console.log(res));
        ConvertToXlsx.convertToXslx(this.data,listName);
        break;
      default:
        throw new Error('Unknown command');
    }
  }
}
