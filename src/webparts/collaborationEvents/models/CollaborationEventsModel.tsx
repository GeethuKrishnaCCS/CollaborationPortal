import * as React from 'react';
import { ICollaborationEventsModelProps, ICollaborationEventsProps, ICollaborationEventsModelState } from '../interfaces/ICollaborationEvents';
import CollaborationEvents from '../components/CollaborationEvents';
// import { CollaborationEventsService } from '../services/collaborationEventsService';
import { CollaborationEventsService } from "../services/CollaborationEventsService";



export default class CollaborationEventModel extends React.Component<ICollaborationEventsModelProps, ICollaborationEventsModelState> {
  private CollaborationEventsService: CollaborationEventsService;
  constructor(props: ICollaborationEventsModelProps) {
    super(props);
    this.CollaborationEventsService = new CollaborationEventsService(this.props.context);
    this.state = {
      userData: { mail: "", title: "" }
    }
  }
  public componentDidMount() {
    this.CollaborationEventsService.getCurrentUser()
      .then((userres: any) => {
        if (userres !== null && userres !== undefined)
          this.setState({ userData: { title: userres.Title, mail: userres.Email } });
      })
  }
  private fetchGroupData = async (): Promise<any> => {
    const groupres = await this.CollaborationEventsService.getCurrentUserGroups();
    return groupres;
  }

  public render(): React.ReactElement<ICollaborationEventsProps> {
    return (
      <>
        {this.state.userData.mail !== "" &&
          <CollaborationEvents
            description={this.props.description}
            isDarkTheme={this.props.isDarkTheme}
            environmentMessage={this.props.environmentMessage}
            hasTeamsContext={this.props.hasTeamsContext}
            userData={this.state.userData}
            fetchGroupData={this.fetchGroupData}
            context={this.props.context}
            Cardlayout={this.props.Cardlayout}
            dataSource={this.props.dataSource}
            StylesForCards={this.props.StylesForCards}
            StylesForImages={this.props.StylesForImages} Events={this.props.Events} SharePointLists={this.props.SharePointLists} location={this.props.location} headingfontcolor={this.props.headingfontcolor}
            Field={this.props.Field}        >
          </CollaborationEvents>}
      </>
    );
  }
}
