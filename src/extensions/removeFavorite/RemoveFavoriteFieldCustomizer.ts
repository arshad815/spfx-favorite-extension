import * as React from 'react';
import * as ReactDOM from 'react-dom';

import { Log } from '@microsoft/sp-core-library';
import { override } from '@microsoft/decorators';
import {
  BaseFieldCustomizer,
  IFieldCustomizerCellEventParameters
} from '@microsoft/sp-listview-extensibility';

import * as strings from 'RemoveFavoriteFieldCustomizerStrings';
import RemoveFavorite from './components/RemoveFavorite';
import { IRemoveFavoriteProps } from './components/RemoveFavoriteProps'
import {sp, Item, ItemAddResult} from '@pnp/sp';

/**
 * If your field customizer uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IRemoveFavoriteFieldCustomizerProperties {
  // This is an example; replace with your own property
  sampleText?: string;
  id: string;
}

const LOG_SOURCE: string = 'RemoveFavoriteFieldCustomizer';

export default class RemoveFavoriteFieldCustomizer
  extends BaseFieldCustomizer<IRemoveFavoriteFieldCustomizerProperties> {

  @override
  public onInit(): Promise<void> {
    // Add your custom initialization to this method.  The framework will wait
    // for the returned promise to resolve before firing any BaseFieldCustomizer events.
    Log.info(LOG_SOURCE, 'Activated RemoveFavoriteFieldCustomizer with properties:');
    Log.info(LOG_SOURCE, JSON.stringify(this.properties, undefined, 2));
    Log.info(LOG_SOURCE, `The following string should be equal: "RemoveFavoriteFieldCustomizer" and "${strings.Title}"`);
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

    this.properties.id = id;

    const removeFavorite: React.ReactElement<{}> =
      React.createElement(RemoveFavorite, { text: text, id: id, onClick: this.onRemoveFavoriteClicked.bind(this)  } as IRemoveFavoriteProps);

    ReactDOM.render(removeFavorite, event.domElement);
  }

  @override
  public onDisposeCell(event: IFieldCustomizerCellEventParameters): void {
    // This method should be used to free any resources that were allocated during rendering.
    // For example, if your onRenderCell() called ReactDOM.render(), then you should
    // call ReactDOM.unmountComponentAtNode() here.
    ReactDOM.unmountComponentAtNode(event.domElement);
    super.onDisposeCell(event);
  }

  private onRemoveFavoriteClicked(id: string): void {
    alert('Remove clicked' + id);
    sp.web.lists.getByTitle(this.context.pageContext.list.title).items.getById(parseInt(id)).delete()
    .then(() => {
      Log.info(LOG_SOURCE,'Deleted item: '+ id);
    })
  }
}
