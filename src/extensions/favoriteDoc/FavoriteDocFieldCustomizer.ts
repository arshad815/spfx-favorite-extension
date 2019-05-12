import * as React from 'react';
import * as ReactDOM from 'react-dom';

import { Log } from '@microsoft/sp-core-library';
import { override } from '@microsoft/decorators';
import {
  BaseFieldCustomizer,
  IFieldCustomizerCellEventParameters
} from '@microsoft/sp-listview-extensibility';

import * as strings from 'FavoriteDocFieldCustomizerStrings';
import FavoriteDoc from './components/FavoriteDoc';
import { IFavoriteDocProps } from './components/IFavoriteDocProps';

import {sp, Item, ItemAddResult} from '@pnp/sp';
import { CurrentUser } from '@pnp/sp/src/siteusers';

import "@pnp/polyfill-ie11";

/**
 * If your field customizer uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IFavoriteDocFieldCustomizerProperties {
  // This is an example; replace with your own property
  sampleText?: string;
  id: string;
}

const LOG_SOURCE: string = 'FavoriteDocFieldCustomizer';

export default class FavoriteDocFieldCustomizer
  extends BaseFieldCustomizer<IFavoriteDocFieldCustomizerProperties> {

  @override
  public onInit(): Promise<void> {
    // Add your custom initialization to this method.  The framework will wait
    // for the returned promise to resolve before firing any BaseFieldCustomizer events.
    Log.info(LOG_SOURCE, 'Activated FavoriteDocFieldCustomizer with properties:');
    Log.info(LOG_SOURCE, JSON.stringify(this.properties, undefined, 2));
    Log.info(LOG_SOURCE, `The following string should be equal: "FavoriteDocFieldCustomizer" and "${strings.Title}"`);
    return super.onInit().then(_ => {
      sp.setup({
        spfxContext: this.context
      });
    })
  }

  @override
  public onRenderCell(event: IFieldCustomizerCellEventParameters): void {
    // Use this method to perform your custom cell rendering.
    const text: string = `${this.properties.sampleText}: ${event.fieldValue}`;
    const id: string = event.listItem.getValueByName('ID').toString();

    const favoriteDoc: React.ReactElement<{}> =
      React.createElement(FavoriteDoc, { text: text, id: id, onClick: this.onAddFavoriteClicked.bind(this)  } as IFavoriteDocProps);

    ReactDOM.render(favoriteDoc, event.domElement);
  }

  @override
  public onDisposeCell(event: IFieldCustomizerCellEventParameters): void {
    // This method should be used to free any resources that were allocated during rendering.
    // For example, if your onRenderCell() called ReactDOM.render(), then you should
    // call ReactDOM.unmountComponentAtNode() here.
    ReactDOM.unmountComponentAtNode(event.domElement);
    super.onDisposeCell(event);
  }

  private onAddFavoriteClicked(id: string): void {
    let userEmail: string;
    let userID: number;
    
    /* sp.utility.getCurrentUserEmailAddresses().then((addressString: string): Promise<any> => {
      return Promise.resolve((addressString as any) as any);
    })
    .then((addressString: any): Promise<Item> => {
       userEmail = addressString; 
       return sp.web.lists.getByTitle(this.context.pageContext.list.title).items.getById(1).expand('File').get();
    }) */
    sp.web.currentUser.get().then((r:CurrentUser) => {
      //console.log(r);
      userID = r['Id'];
      userEmail = r['Email'];
      //return sp.web.lists.getByTitle(this.context.pageContext.list.title).items.getById(parseInt(id)).expand('File').get();
      return sp.web.lists.getByTitle(this.context.pageContext.list.title).items.getById(parseInt(id)).get();
    })
    .then((item: Item): Promise<ItemAddResult> => {
      console.log(item);
      console.log(userEmail);
      //console.log(item['File'].Name + ' ' + item['File'].LinkingUrl);
      return sp.web.lists.getByTitle("My Master Library").items.add({
        Report_x0020_Name: {
          "__metadata": { type: "SP.FieldUrlValue" },
          Url: item['Report_x0020_Name0'].Url,
          Description: item['Report_x0020_Name0'].Description
        },
        Report_x0020_Number: item['Report_x0020_Number'],
        Report_x0020_Type: item['Report_x0020_Type'],
        FavoritedById: userID,
        Subject_x0020_Area: item["Subject_x0020_Area"],
        Description: item["Description"]
      })
    })
    .then((result: ItemAddResult): void => {
      //console.log("Add item has " + result.data)
      alert('Document added to favorites');
    });

    
  }
  
}

