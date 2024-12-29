import * as React from 'react';
import styles from './UploadAttachment.module.scss';
import type { IUploadAttachmentProps } from './IUploadAttachmentProps';
import { escape } from '@microsoft/sp-lodash-subset';
import MyDialogPopup from './MyModalPopupWebPart';

export default class UploadAttachment extends React.Component<IUploadAttachmentProps> {
  public render(): React.ReactElement<IUploadAttachmentProps> {
    const {
      description,
     // isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName,
      absoluteURL
    } = this.props;

    return (
      <section className={`${styles.uploadAttachment} ${hasTeamsContext ? styles.teams : ''}`}>
        <div className={styles.welcome}>
               <h2>Well done, {escape(userDisplayName)}!</h2>
          <div>{environmentMessage}</div>
          <h1>{absoluteURL}</h1>
          <div>Web part property value: <strong>{escape(description)}</strong></div>
        </div>
        <div>

          <MyDialogPopup absoluteURL={absoluteURL} spHttpClient={this.context.spHttpClient}/>
        </div>
      </section>
    );
  }
}
