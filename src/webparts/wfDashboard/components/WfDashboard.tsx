import * as React from 'react';
// import styles from './WfDashboard.module.scss';
import { IWfDashboardProps } from './IWfDashboardProps';
import { DashBoardAdmin } from '../components/RequestDashboard';
import { escape } from '@microsoft/sp-lodash-subset';
import RequestDashboard from "./RequestDashboard";
import { sp } from "@pnp/sp/presets/all";


export default class WfDashboard extends React.Component<IWfDashboardProps, {}> {

  constructor(prop: IWfDashboardProps, state: {}) {
    super(prop);
    sp.setup({
      spfxContext: this.props.context
    });
  }

  public render(): React.ReactElement<IWfDashboardProps> {
    return (
      <DashBoardAdmin description={this.props.description} context={this.props.context} spcontext={sp.web} />
    );
  }
}


