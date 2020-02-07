import * as React from 'react';
import styles from './AssetBuilder.module.scss';
import { IAssetBuilderProps } from './IAssetBuilderProps';
import { escape } from '@microsoft/sp-lodash-subset';

export default class AssetBuilder extends React.Component<IAssetBuilderProps, {}> {

  

  public render(): React.ReactElement<IAssetBuilderProps> {
    console.log('render Props:', this.props );
    return (
      <div className={ styles.assetBuilder }>
        <div className={ styles.container }>
          <div className={ styles.row }>
            <div className={ styles.column }>
              <span className={ styles.title }>Welcome to SharePoint!</span>
              <p className={ styles.subTitle }>Customize SharePoint experiences using Web Parts.</p>
              <p className={ styles.description }>{escape(this.props.description)}</p>
              <a href="https://aka.ms/spfx" className={ styles.button }>
                <span className={ styles.label }>Learn more</span>
                { this.props.buildStatus.map( item => <div key={item} > { item } </div> ) }
              </a>
            </div>
          </div>
        </div>
      </div>
    );
  }
}
