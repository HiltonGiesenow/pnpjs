import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import PnpJsTestingWebpart from './components/PnpJsTestingWebpart';
import { IPnpJsTestingWebpartProps } from './components/IPnpJsTestingWebpartProps';

export interface IPnpJsTestingWebpartWebPartProps { }

export default class PnpJsTestingWebpartWebPart extends BaseClientSideWebPart<IPnpJsTestingWebpartWebPartProps> {

	public render(): void {
		const element: React.ReactElement<IPnpJsTestingWebpartProps> = React.createElement(
			PnpJsTestingWebpart,
			{
				context: this.context,
			}
		);

		ReactDom.render(element, this.domElement);
	}

	protected onDispose(): void {
		ReactDom.unmountComponentAtNode(this.domElement);
	}

	protected get dataVersion(): Version {
		return Version.parse('1.0');
	}

}
