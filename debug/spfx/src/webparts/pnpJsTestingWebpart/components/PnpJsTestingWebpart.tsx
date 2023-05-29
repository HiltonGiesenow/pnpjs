import * as React from 'react';

import { spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/views";
import "@pnp/sp/fields";
import "@pnp/sp/items";
import "@pnp/sp/site-users/web";
import "@pnp/sp/sputilities";

import styles from './PnpJsTestingWebpart.module.scss';
import { IPnpJsTestingWebpartProps } from './IPnpJsTestingWebpartProps';
import { IPnpJsTestingWebpartState } from './IPnpJsTestingWebpartState';

export default class PnpJsTestingWebpart extends React.Component<IPnpJsTestingWebpartProps, IPnpJsTestingWebpartState> {
	private sp = spfi().using(SPFx({ pageContext: this.props.context.pageContext }));

	private testInfo = {
		testWebName: "TestWeb",
	}

	public constructor(props: IPnpJsTestingWebpartProps) {
		super(props);

		this.state = {
			testInProgress: false,
		};
	}

	public async componentDidMount(): Promise<void> {
		// anything to load? Maybe later with certain tests?

		document.getElementById('workbenchPageContent').style.maxWidth = "90%";
		(document.getElementsByClassName('CanvasZone')[0] as HTMLDivElement).style.maxWidth = "100%";
	}

	public addLogEntry(message: string, color: string = "darkgray"): void {
		const logContainer = document.getElementById("logContainer");
		if (logContainer) {
			logContainer.innerHTML += `<div style="color: ${color}">${message}</div>`;
		}
	}

	public async addWebTest(): Promise<void> {
		this.setState({ testInProgress: true });

		this.addLogEntry("&nbsp;");

		try {
			this.addLogEntry("Adding new web...");

			const newWeb = await this.sp.web.webs.add(this.testInfo.testWebName, this.testInfo.testWebName, "Testing PnP JS", "STS#0", 1033, true);

			this.addLogEntry(`New web created at <a target="_blank" rel="noreferrer" href="${newWeb.data.ServerRelativeUrl}" style="color: green">${newWeb.data.ServerRelativeUrl}</a>`, "green");
		} catch (e) {
			this.addLogEntry("Web creation failed: " + e.message, "red");
		}

		this.setState({ testInProgress: false });
	}

	public async deleteWebTest(): Promise<void> {
		this.setState({ testInProgress: true });

		this.addLogEntry("&nbsp;");
		
		try {
			this.addLogEntry("Verifying preconditions...");
			
			const subWeb = await this.sp.web.getSubwebsFilteredForCurrentUser().filter("Title eq '" + this.testInfo.testWebName + "'")();

			if (subWeb.length === 0) {
				this.addLogEntry("Subweb for test delete does not exist. Run the 'Add new web' test first.", "red");
			} else {
				this.addLogEntry("Deleting sub web...");

				const sp2 = spfi().using(SPFx({ pageContext: { web: { absoluteUrl: this.props.context.pageContext.web.absoluteUrl + "/" + this.testInfo.testWebName }, legacyPageContext: null } }));
	
				await sp2.web.delete();
	
				this.addLogEntry("Sub web deleted", "green");
			}
		} catch (e) {
			this.addLogEntry("Web deletion failed: " + e.message, "red");
		}

		this.setState({ testInProgress: false });
	}

	public render(): React.ReactElement<IPnpJsTestingWebpartProps> {
		return (
			<section className={styles.pnpJsTestingWebpart}>
				<div className={styles.controlsContainer}>
					<h1>Webs</h1>
					<div>
						<h2>Add new web</h2>
						<p>Adds a new sub web to the existing web, called {this.testInfo.testWebName}</p>
						<button disabled={this.state.testInProgress} onClick={this.addWebTest.bind(this)}>Add web</button>
					</div>
					<div>
						<h2>Delete web</h2>
						<p>Deletes an existing sub web called {this.testInfo.testWebName}</p>
						<button disabled={this.state.testInProgress} onClick={this.deleteWebTest.bind(this)}>Delete web</button>
					</div>
				</div>
				<div className={styles.logContainer} id="logContainer">
					Logging started...
				</div>
			</section>
		);
	}
}