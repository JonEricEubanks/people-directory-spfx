import * as React from 'react';
import styles from './PersonaCard.module.scss';
import { IPersonaCardProps } from './IPersonaCardProps';
import { IPersonaCardState } from './IPersonaCardState';
import { Log } from '@microsoft/sp-core-library';
import { SPComponentLoader } from '@microsoft/sp-loader';

import {
  Avatar,
  Card,
} from '@fluentui/react-components';
import { Call16Filled, Location16Filled, Mail16Filled } from '@fluentui/react-icons';
import { LIVE_PERSONA_COMPONENT_ID, EXP_SOURCE } from '../../../../constants';

export class PersonaCard extends React.Component<
  IPersonaCardProps,
  IPersonaCardState
> {
  constructor(props: IPersonaCardProps) {
    super(props);

    this.state = { livePersonaCard: undefined, pictureUrl: undefined };
  }
  /**
   *
   *
   * @memberof PersonaCard
   */
  public async componentDidMount(): Promise<void> {
    const sharedLibrary = await this._loadSPComponentById(
      LIVE_PERSONA_COMPONENT_ID
    );
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const livePersonaCard: any = sharedLibrary.LivePersonaCard;
    this.setState({ livePersonaCard: livePersonaCard });
  }

  /**
   *
   *
   * @private
   * @returns
   * @memberof PersonaCard
   */
  private _LivePersonaCard(): JSX.Element {
    return React.createElement(
      this.state.livePersonaCard,
      {
        serviceScope: this.props.context.serviceScope,
        legacyUpn: this.props.profileProperties.Email,
        onCardOpen: () => {
          console.log('LivePersonaCard Open');
        },
        onCardClose: () => {
          console.log('LivePersonaCard Close');
        },
      },
      this._PersonaCard()
    );
  }

  /**
   *
   *
   * @private
   * @returns {JSX.Element}
   * @memberof PersonaCard
   */
  private _PersonaCard(): JSX.Element {
    const { DisplayName, Title, PictureUrl, Department, WorkPhone, Location, Email } =
      this.props.profileProperties;
    const hasContact = WorkPhone || Location || Email;
    return (
      <Card className={styles.documentCard} style={{ padding: 0 }}>
        <div className={styles.cardHeader} />
        <div className={styles.avatarWrapper}>
          <Avatar
            name={DisplayName}
            image={{ src: PictureUrl }}
            size={72}
          />
        </div>
        <div className={styles.personaDetails}>
          <span className={styles.displayName}>{DisplayName}</span>
          {Title && <span className={styles.jobTitle}>{Title}</span>}
          {Department && <span className={styles.departmentPill}>{Department}</span>}
          {hasContact && <div className={styles.divider} />}
          {hasContact && (
            <div className={styles.contactInfo}>
              {WorkPhone && (
                <div className={styles.others}>
                  <Call16Filled />
                  <span className={styles.textOverflow}>{WorkPhone}</span>
                </div>
              )}
              {Location && (
                <div className={styles.others}>
                  <Location16Filled />
                  <span className={styles.textOverflow}>{Location}</span>
                </div>
              )}
              {Email && (
                <div className={styles.others}>
                  <Mail16Filled />
                  <a href={`mailto:${Email}`} className={styles.emailLink}>{Email}</a>
                </div>
              )}
            </div>
          )}
        </div>
      </Card>
    );
  }
  /**
   * Load SPFx component by id, SPComponentLoader is used to load the SPFx components
   * @param componentId - componentId, guid of the component library
   */
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  private async _loadSPComponentById(componentId: string): Promise<any> {
    try {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const component: any = await SPComponentLoader.loadComponentById(
        componentId
      );
      return component;
    } catch (error) {
      const err = error instanceof Error ? error : new Error(String(error));
      Log.error(EXP_SOURCE, err);
      throw err;
    }
  }

  /**
   *
   *
   * @returns {React.ReactElement<IPersonaCardProps>}
   * @memberof PersonaCard
   */
  public render(): React.ReactElement<IPersonaCardProps> {
    const delay = `${(this.props.cardIndex || 0) * 0.06}s`;
    return (
      <div className={styles.personaContainer} style={{ '--card-delay': delay } as React.CSSProperties}>
        {this.state.livePersonaCard
          ? this._LivePersonaCard()
          : this._PersonaCard()}
      </div>
    );
  }
}
