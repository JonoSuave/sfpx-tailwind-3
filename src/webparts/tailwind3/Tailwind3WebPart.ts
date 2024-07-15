import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import {
	type IPropertyPaneConfiguration,
	PropertyPaneTextField,
} from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import { IReadonlyTheme } from "@microsoft/sp-component-base";

import * as strings from "Tailwind3WebPartStrings";
import Tailwind3 from "./components/Tailwind3";
import { ITailwind3Props } from "./components/ITailwind3Props";
import "../../styles/dist/tailwind.css";

export interface ITailwind3WebPartProps {
	description: string;
}

export enum ThemeCSSVariables {
	fontFamilyPrimary = "--myWebPart-fontPrimary",
	colorPrimary = "--myWebPart-primary",
	background = "--myWebPart-background",
	primaryBackgroundColorDark = "--myWebPart-colorBackgroundDarkPrimary",
	bodyText = "--myWebPart-bodyText",
	link = "--myWebPart-link",
	linkHover = "--myWebPart-linkHover",
}

export default class Tailwind3WebPart extends BaseClientSideWebPart<ITailwind3WebPartProps> {
	private _isDarkTheme: boolean = false;
	private _environmentMessage: string = "";

	public render(): void {
		const element: React.ReactElement<ITailwind3Props> = React.createElement(Tailwind3, {
			description: this.properties.description,
			isDarkTheme: this._isDarkTheme,
			environmentMessage: this._environmentMessage,
			hasTeamsContext: !!this.context.sdks.microsoftTeams,
			userDisplayName: this.context.pageContext.user.displayName,
		});

		ReactDom.render(element, this.domElement);
	}

	protected onInit(): Promise<void> {
		return this._getEnvironmentMessage().then((message) => {
			this._environmentMessage = message;
		});
	}

	private _getEnvironmentMessage(): Promise<string> {
		if (!!this.context.sdks.microsoftTeams) {
			// running in Teams, office.com or Outlook
			return this.context.sdks.microsoftTeams.teamsJs.app.getContext().then((context) => {
				let environmentMessage: string = "";
				switch (context.app.host.name) {
					case "Office": // running in Office
						environmentMessage = this.context.isServedFromLocalhost
							? strings.AppLocalEnvironmentOffice
							: strings.AppOfficeEnvironment;
						break;
					case "Outlook": // running in Outlook
						environmentMessage = this.context.isServedFromLocalhost
							? strings.AppLocalEnvironmentOutlook
							: strings.AppOutlookEnvironment;
						break;
					case "Teams": // running in Teams
					case "TeamsModern":
						environmentMessage = this.context.isServedFromLocalhost
							? strings.AppLocalEnvironmentTeams
							: strings.AppTeamsTabEnvironment;
						break;
					default:
						environmentMessage = strings.UnknownEnvironment;
				}

				return environmentMessage;
			});
		}

		return Promise.resolve(
			this.context.isServedFromLocalhost
				? strings.AppLocalEnvironmentSharePoint
				: strings.AppSharePointEnvironment
		);
	}

	protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
		if (!currentTheme) {
			return;
		}

		this._isDarkTheme = !!currentTheme.isInverted;
		const { semanticColors } = currentTheme;

		if (semanticColors) {
			console.log(semanticColors);

			// Map colors from the theme to CSS variables and therefore to TailwindCSS custom colors so they can be used in classes
			this.domElement.style.setProperty(
				ThemeCSSVariables.colorPrimary,
				currentTheme?.palette?.themePrimary || null
			);
			this.domElement.style.setProperty(
				ThemeCSSVariables.fontFamilyPrimary,
				currentTheme?.fonts?.medium?.fontFamily || null
			);
			this.domElement.style.setProperty(
				ThemeCSSVariables.bodyText,
				semanticColors.bodyText || null
			);
			this.domElement.style.setProperty(ThemeCSSVariables.link, semanticColors.link || null);
			this.domElement.style.setProperty(
				ThemeCSSVariables.linkHover,
				semanticColors.linkHovered || null
			);
			this.domElement.style.setProperty(
				ThemeCSSVariables.background,
				semanticColors.bodyBackground || null
			);
		}
	}

	protected onDispose(): void {
		ReactDom.unmountComponentAtNode(this.domElement);
	}

	protected get dataVersion(): Version {
		return Version.parse("1.0");
	}

	protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
		return {
			pages: [
				{
					header: {
						description: strings.PropertyPaneDescription,
					},
					groups: [
						{
							groupName: strings.BasicGroupName,
							groupFields: [
								PropertyPaneTextField("description", {
									label: strings.DescriptionFieldLabel,
								}),
							],
						},
					],
				},
			],
		};
	}
}
