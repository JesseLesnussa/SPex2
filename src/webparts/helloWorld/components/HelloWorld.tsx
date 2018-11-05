import * as React from 'react';
import { IUserProps } from './IUserProps';
import { IUsersState } from './IUsersState';
import { List } from 'office-ui-fabric-react/lib/List';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import { IPersonaSharedProps, Persona, PersonaInitialsColor } from 'office-ui-fabric-react/lib/Persona';
import { Icon } from 'office-ui-fabric-react/lib/Icon';
import { HoverCard, IExpandingCardProps } from 'office-ui-fabric-react/lib/HoverCard';
import { KeyCodes } from '@uifabric/utilities';
import { DetailsList, buildColumns, IColumn } from 'office-ui-fabric-react/lib/DetailsList';
import { DirectionalHint } from 'office-ui-fabric-react/lib/common/DirectionalHint';

let _items: any[];

export interface IHoverCardExampleState {
  items?: any[];
  columns?: IColumn[];
}

interface IHoverCardFieldProps {
  componentRef?: any;
  content: HTMLDivElement;
  expandingCardProps: IExpandingCardProps;
}

interface IHoverCardFieldState {
  contentRendered?: HTMLDivElement;
}





export default class HelloWorld extends React.Component<IUserProps, IUsersState> {

  constructor(props: IUserProps) {
    super(props);

    this.state = {
      users: [],
      contentRendered: undefined
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
      <div>
        <Persona {...user} text={user.displayName} secondaryText= {user.mail} imageUrl={'/_layouts/15/userphoto.aspx?accountname='+ user.mail +'&size=S'} />
      </div>
      
      
    );
  }

  public render(): React.ReactElement<IUserProps> {
    return (
      <List items ={this.state.users}
        onRenderCell={this._onRenderEventCell} />
        
    );
  }
}
