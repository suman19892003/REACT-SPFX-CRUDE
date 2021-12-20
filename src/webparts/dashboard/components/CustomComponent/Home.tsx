import * as React from 'react';
import { escape, find } from "@microsoft/sp-lodash-subset";

import { Link } from 'react-router-dom';

export default class Landing extends React.Component<any, any> {
 
  public render(): React.ReactElement<any> {
    
    return (
        <div className="w100">
        <h4>SharePoint Crude Operation</h4>
        <ul>
          <li><Link to={`/list`}>List Operation</Link></li>
          <li><Link to={`/library`}>Library Operation</Link> </li>
        </ul>      
        </div>
      );
    }
  }
  