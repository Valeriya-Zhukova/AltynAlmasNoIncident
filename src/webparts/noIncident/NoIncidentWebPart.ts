import { Version } from '@microsoft/sp-core-library';
import {
	IPropertyPaneConfiguration,
	PropertyPaneTextField,
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';
import { getIconClassName } from '@uifabric/styling';

import styles from './NoIncidentWebPart.module.scss';
import * as strings from 'NoIncidentWebPartStrings';

export interface INoIncidentWebPartProps {
	description: string;
}

import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

export interface ISPLists {
	value: ISPList[];
}
export interface ISPList {
	Title: string;
	Project: string;
	iDate: Date;
	Event: string;
	Plan: number;
	Fact: number;
}

export default class NoIncidentWebPart extends BaseClientSideWebPart<INoIncidentWebPartProps> {
	private _getListData(): Promise<ISPLists> {
		return this.context.spHttpClient
			.get(
				this.context.pageContext.web.absoluteUrl +
					`/_api/web/lists/GetByTitle('NoIncident')/Items`,
				SPHttpClient.configurations.v1
			)
			.then((response: SPHttpClientResponse) => {
				debugger;
				return response.json();
			});
	}

	private _renderListAsync(): void {
		this._getListData().then((response) => {
			this._renderList(response.value);
		});
	}

	private _renderList(items: ISPList[]): void {
		let projDay = 0;
		let one_day = 1000 * 60 * 60 * 24;
		const newLocal = new Date().getTime() + 6 * 60 * 60 * 1000;

		let html: string = ``;

		items.forEach((item: ISPList) => {
			let d = newLocal - new Date(item.iDate.toString()).getTime();
			projDay = Math.round(Math.abs(d / one_day) - 1);

			const plan = item.Plan;
			const fact = item.Fact;

			html += `
			<div class="${styles.project} ${styles.back} ${styles.round}">
				<div class="${styles.subTitle}">
					<div>Проект</div>
					<div class="${styles.projectTitle}">${item.Project}</div>
				</div>
				<div class="${styles.days}">
					<div class="${styles.daysWrapper}">
						<ul class="${styles.daysNum}">
							<li class="${styles.number}">
								<span class="${styles.opas}">${projDay - 1}</span>
							</li>
							<li class="${styles.number}">
								<span class="${styles.numwhite}">${projDay}</span>
							</li>
							<li class="${styles.number}">
								<span class="${styles.opas}">${projDay + 1}</span>
							</li> 
						</ul>
						<div class="${styles.text}">дн.</div> 
					</div>	
					<div class="${styles.info}">
						<i class="${getIconClassName('InfoSolid')}"></i>
						<div class="${styles.tooltiptext}">
							<span class="${styles.text}">${item.Event}</span>
							<div class="${styles.table}">
								<div class="${styles.tr}">
									<span class="${styles.th}"> </span>
									<span class="${styles.th}"> План </span>
									<span class="${styles.th}"> Факт </span>
								</div>
								<div class="${styles.tr}">  
									<span class="${styles.td}"> LTIFR </span>
									<span class="${styles.td}"> ${plan.toFixed(2)} </span>
									<span class="${styles.td} 
													${plan > fact ? styles.factGreen : styles.factRed}
									"> 
												${fact.toFixed(2)}    
									</span>                   
								</div>
							</div>
						</div>
					</div>			
				</div> 				
			</div>`;
		});

		const listContainer: Element =
			this.domElement.querySelector('#spListContainer');

		listContainer.innerHTML = html;
	}

	public render(): void {
		var test = new Date().toLocaleDateString();
		this.domElement.innerHTML = `
			<div class="${styles.noIncident}">
				<div class="${styles.container}">
					<div class="${styles.color}"> 
					<div class="${styles.square} ${styles.title}">
						<div class="${styles.back} ${styles.title1} ${styles.round}">
							Дни без происшествий
							<br/>с потерей рабочих дней
						</div>
						<div id="spListContainer"> </div>    
					</div>
					</div>
				</div>
			</div>`;
		this._renderListAsync();
	}

	protected get dataVersion(): Version {
		return Version.parse('1.0');
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
								PropertyPaneTextField('description', {
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
