import * as React from 'react';
import styles from './Dashboard.module.scss';
import { IDashboardProps } from './IDashboardProps';
import { escape } from '@microsoft/sp-lodash-subset';

import { HashRouter, Route } from "react-router-dom";
import { Nav, INavStyles, INavLinkGroup } from 'office-ui-fabric-react/lib/Nav';
import { Stack, IStackTokens } from 'office-ui-fabric-react/lib/Stack';

import { Link } from 'react-router-dom';

import Home from './CustomComponent/Home';
import LandingList from './CustomComponent/LandingList';
import LandingLibrary from './CustomComponent/LandingLibrary';
import ListAddEdit from './CustomComponent/ListAddEdit';
import LibraryAddEdit from './CustomComponent/LibraryAddEdit';

import {RequestServices} from '../Services/ServiceRequest';

const stackTokens: IStackTokens = { childrenGap: 40 };

export default class Dashboard extends React.Component<IDashboardProps, {}> {

  public render(): React.ReactElement<IDashboardProps> {
    return (
      <div className={ styles.dashboard }>
      
        <Stack horizontal tokens={stackTokens}>

          <HashRouter>      
            <Route path="/" exact render={(props) => <Home webURL={this.props.webURL} context={this.props.context} {...props} />} />
            <Route path="/list" exact render={(props) => <LandingList webURL={this.props.webURL} context={this.props.context} {...props} />} />
            <Route path="/list/addedit" render={(props) => <ListAddEdit webURL={this.props.webURL} context={this.props.context} {...props} />} />
            <Route path="/library" exact render={(props) => <LandingLibrary webURL={this.props.webURL} context={this.props.context} {...props} />} />
            <Route path="/library/addedit" render={(props) => <LibraryAddEdit webURL={this.props.webURL} context={this.props.context} {...props} />} />
          </HashRouter>
        </Stack>
      </div>
    );
  }
}
