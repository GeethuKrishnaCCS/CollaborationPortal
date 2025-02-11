import * as React from 'react';
//import styles from './BirthdayWp.module.scss';
import { IBirthdayWpProps } from './IBirthdayWpProps';
import { IBirthdayWpState } from './IBirthdayWpState';
import { BirthdayController } from '../../../controllers/BirthdayController';
import { BirthdayModel } from '../../../models/BirthdayModel';
import { BirthdayView } from '../../../views/BirthdayView';

export default class BirthdayWp extends React.Component<IBirthdayWpProps, IBirthdayWpState> {
  private controller: BirthdayController;

  constructor(props: IBirthdayWpProps) {
    super(props);
    this.state = {
      birthdays: []
    };

    const model = new BirthdayModel();
    this.controller = new BirthdayController(model);
  }

  public async componentDidMount() {
    // Fetch birthdays and update state
    const birthdays = await this.controller.model.getTodaysBirthdays();
    this.setState({ birthdays });
  }

  public render(): React.ReactElement<IBirthdayWpProps> {
    const {
      userDisplayName
    } = this.props;
    const renderBirthdays = () => (
      <BirthdayView birthdays={this.state.birthdays} userDisplayName={userDisplayName} layout={this.props.layout as 'one-column' | 'two-columns' | 'three-columns'} description={this.props.description} environmentMessage={this.props.environmentMessage}/>
    );
    return (
      renderBirthdays()
    );
  }
}