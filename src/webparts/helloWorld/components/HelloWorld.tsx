import * as React from 'react';
import { IUserProps } from './IUserProps';
import { IUsersState } from './IUsersState';
import { List } from 'office-ui-fabric-react/lib/List';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import { IPersonaSharedProps, Persona, PersonaInitialsColor } from 'office-ui-fabric-react/lib/Persona';

export default class HelloWorld extends React.Component<IUserProps, IUsersState> {

  constructor(props: IUserProps) {
    super(props);

    this.state = {
      users: []
    };
  }

  public componentDidMount(): void {
    this.props.graphClient
    .api('/users')
    .get((error:any, usersResponse: any, rawResponse?: any) => {
      console.log('Users', usersResponse);
      console.log('Error', error);
      const userList:MicrosoftGraph.User[] = usersResponse.value;
      this.setState({users: userList});
    });
  }

  private _onRenderEventCell(user: MicrosoftGraph.User, index: number | undefined): JSX.Element{
    return (
      <Persona {...user} text={user.displayName} secondaryText= {user.mail} />
    );
  }

  public render(): React.ReactElement<IUserProps> {
    return (
      <List items ={this.state.users}
        onRenderCell={this._onRenderEventCell} />
    );
  }
}
