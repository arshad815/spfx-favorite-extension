import { Log } from '@microsoft/sp-core-library';
import { override } from '@microsoft/decorators';
import * as React from 'react';
import {PrimaryButton, IButtonProps} from 'office-ui-fabric-react/lib/Button';
import {Label} from 'office-ui-fabric-react/lib/Label'

import styles from './RemoveFavorite.module.scss';

import {IRemoveFavoriteProps} from './RemoveFavoriteProps'
import {IRemoveFavoriteState} from './RemoveFavoriteState'

const LOG_SOURCE: string = 'RemoveFavorite';

export default class RemoveFavorite extends React.Component<IRemoveFavoriteProps, {}> {
  @override
  public componentDidMount(): void {
    Log.info(LOG_SOURCE, 'React Element: RemoveFavorite mounted');
  }

  @override
  public componentWillUnmount(): void {
    Log.info(LOG_SOURCE, 'React Element: RemoveFavorite unmounted');
  }

  @override
  public render(): React.ReactElement<{}> {
    return (
      <div className={styles.cell}>
        <Label>Remove From Favorites</Label>>
        <PrimaryButton
          text="Remove From Favorites"
          onClick={this.onClick.bind(this)}
          />
      </div>
    );
  }

  private onClick(): void {
    if (this.props.onClick) {
      this.props.onClick(this.props.id);
    }

  }
}
