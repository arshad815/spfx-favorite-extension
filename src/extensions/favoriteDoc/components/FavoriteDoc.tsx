import { Log } from '@microsoft/sp-core-library';
import { override } from '@microsoft/decorators';
import * as React from 'react';
import {PrimaryButton, IButtonProps} from 'office-ui-fabric-react/lib/Button';
import {Label} from 'office-ui-fabric-react/lib/Label'

import styles from './FavoriteDoc.module.scss';
import { IFavoriteDocState } from './IFavoriteDocState';
import { IFavoriteDocProps } from './IFavoriteDocProps';



const LOG_SOURCE: string = 'FavoriteDoc';

export default class FavoriteDoc extends React.Component<IFavoriteDocProps, IFavoriteDocState> {
  @override
  public componentDidMount(): void {
    Log.info(LOG_SOURCE, 'React Element: FavoriteDoc mounted');
  }

  @override
  public componentWillUnmount(): void {
    Log.info(LOG_SOURCE, 'React Element: FavoriteDoc unmounted');
  }

  @override
  public render(): React.ReactElement<{}> {
    return (
      <div className={styles.cell}>
        <Label>Add To Favorites</Label>>
        <PrimaryButton
          text="Add to Favorites"
          onClick={this.onClick.bind(this)}
          />
      </div>
    );
  }

  private onClick(): void {
    if (this.props.onClick)
      this.props.onClick();
  }

 /*  private _alertClicked(): void {
    let etag: string = undefined;
    sp.web.lists.getByTitle(this.context.pageContext.list.title).items.getById(1).get().then((item: any) => {
      console.log(item);
    });

    sp.utility.getCurrentUserEmailAddresses().then((addressString: string) => {
      console.log(addressString);
    })

    alert('Clicked');
  } */
}
