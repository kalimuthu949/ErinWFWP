import * as React from 'react';
import styles from './QuoteRequest.module.scss';
import{IQuoteRequestProps} from './IQuoteRequestProps'
import { RequestNewQuoteAdmin } from '../components/QuoteRequestDB';
import { escape } from '@microsoft/sp-lodash-subset';
import { sp } from "@pnp/sp/presets/all";
import '../../../ExternalRef/css/style.css'

export default class QuoteRequest extends React.Component<IQuoteRequestProps, {}> {

  constructor(prop: IQuoteRequestProps, state: {}) {
    super(prop);
    sp.setup({
      spfxContext: this.props.context
    });
  }

  public render(): React.ReactElement<IQuoteRequestProps> {
    return (
      <RequestNewQuoteAdmin description={this.props.description} context={this.props.context} spcontext={sp.web} />
    );
  }
}
