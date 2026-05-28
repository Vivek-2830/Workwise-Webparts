import * as React from 'react';
import styles from './Anniversary.module.scss';
import { IAnniversaryProps } from './IAnniversaryProps';
import { escape } from '@microsoft/sp-lodash-subset';

export default class Anniversary extends React.Component<IAnniversaryProps, {}> {
  public render(): React.ReactElement<IAnniversaryProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName
    } = this.props;

    return (
      <section className="anniversary">
     
      </section>
    );
  }
}
