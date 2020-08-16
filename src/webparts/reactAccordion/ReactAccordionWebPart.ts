import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version, DisplayMode } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import { IPropertyPaneConfiguration, PropertyPaneSlider, PropertyPaneTextField } from "@microsoft/sp-property-pane";
import { PropertyFieldListPicker, PropertyFieldListPickerOrderBy } from '@pnp/spfx-property-controls/lib/PropertyFieldListPicker';
import * as strings from 'ReactAccordionWebPartStrings';
import ReactAccordion, { ReactAccordion1 } from './components/ReactAccordion';
import { IReactAccordionProps } from './components/IReactAccordionProps';
import { sp } from "@pnp/sp";

export interface IReactAccordionWebPartProps {
    listid: string;
    title: string;
    displayMode: DisplayMode;
    maxItemsPerPage: number;
    updateProperty: (value: string) => void;
}

export default class ReactAccordionWebPart extends BaseClientSideWebPart<IReactAccordionWebPartProps> {

    public onInit(): Promise<void> {
        return super.onInit().then(_ => {
            sp.setup({
                spfxContext: this.context
            });
        });
    }

    public render(): void {
        const element: React.ReactElement<IReactAccordionProps> = React.createElement(
            ReactAccordion1,
            {
                listId: this.properties.listid,
                title: this.properties.title,
                displayMode: this.displayMode,
                maxItemsPerPage: this.properties.maxItemsPerPage,
                updateProperty: (value: string) => {
                    this.properties.title = value;
                },
                configurePropertyPane: this.onConfigurePropertyPane
            }
        );

        ReactDom.render(element, this.domElement);
    }

    private onConfigurePropertyPane = () => {
        this.context.propertyPane.open();
    }

    protected onDispose(): void {
        ReactDom.unmountComponentAtNode(this.domElement);
    }

    protected get dataVersion(): Version {
        return Version.parse('1.0');
    }

    protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
        return {
            pages: [
                {
                    header: {
                        description: strings.PropertyPaneDescription
                    },
                    groups: [
                        {
                            groupName: strings.BasicGroupName,
                            groupFields: [
                                PropertyFieldListPicker('listid', {
                                    label: 'Select a list',
                                    selectedList: this.properties.listid,
                                    includeHidden: false,
                                    orderBy: PropertyFieldListPickerOrderBy.Title,
                                    disabled: false,
                                    onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                                    properties: this.properties,
                                    context: this.context,
                                    onGetErrorMessage: null,
                                    deferredValidationTime: 0,
                                    key: 'listidFieldId'
                                }),
                                PropertyPaneSlider('maxItemsPerPage', {
                                    label: strings.MaxItemsPerPageLabel,
                                    ariaLabel: strings.MaxItemsPerPageLabel,
                                    min: 3,
                                    max: 20,
                                    value: 5,
                                    showValue: true,
                                    step: 1
                                }),
                            ]
                        }
                    ]
                }
            ]
        };
    }
}
