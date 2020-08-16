import * as React from 'react';
import { useEffect, useState } from 'react';
import styles from './ReactAccordion.module.scss';
import * as strings from 'ReactAccordionWebPartStrings';
import { IReactAccordionProps } from './IReactAccordionProps';
import { PrimaryButton } from 'office-ui-fabric-react/lib/Button';
import { SearchBox } from 'office-ui-fabric-react/lib/SearchBox';
import { Spinner, SpinnerSize } from 'office-ui-fabric-react/lib/Spinner';
import { ProgressIndicator } from 'office-ui-fabric-react/lib/ProgressIndicator';
import { MessageBar, MessageBarType } from 'office-ui-fabric-react/lib/MessageBar';
import { Accordion, AccordionItem, AccordionItemTitle, AccordionItemBody } from 'react-accessible-accordion';
import 'react-accessible-accordion/dist/react-accessible-accordion.css';
import { IReactAccordionState } from "./IReactAccordionState";
import IAccordionListItem from "../models/IAccordionListItem";
import { WebPartTitle } from "@pnp/spfx-controls-react/lib/WebPartTitle";
import { Placeholder } from "@pnp/spfx-controls-react/lib/Placeholder";
import { sp } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items/list";
import './accordion.css';
import { DisplayMode } from '@microsoft/sp-core-library';

export const ReactAccordion1: React.FC<IReactAccordionProps> = (props) => {
    const [state, setState] = useState<IReactAccordionState>({
        status: '',
        items: [],
        listItems: [],
        isLoading: true,
        loaderMessage: ''
    });

    const _readItems = async () => {
        try {
            let listItemsCollection: any[] = await sp.web.lists.getById(props.listId).items
                .select('ID', 'Title', 'Description')
                .get();
            setState({
                ...state,
                items: listItemsCollection.length > 0 ? listItemsCollection.splice(0, props.maxItemsPerPage) : [],
                listItems: listItemsCollection,
                isLoading: false
            });
        } catch (err) {
            setState({
                ...state,
                items: [],
                listItems: [],
                isLoading: false,
                loaderMessage: err.toString()
            });
        }
    };

    const items: JSX.Element[] = state.items.map((item: IAccordionListItem, i: number): JSX.Element => {
        return (
            <AccordionItem>
                <AccordionItemTitle className="accordion__title">
                    <h3 className="u-position-relative ms-fontColor-white">{item.Title}</h3>
                    <div className="accordion__arrow ms-fontColor-white" role="presentation" />
                </AccordionItemTitle>
                <AccordionItemBody className="accordion__body">
                    <div className="" dangerouslySetInnerHTML={{ __html: item.Description }}>
                    </div>
                </AccordionItemBody>
            </AccordionItem>
        );
    });

    const _searchData = (newValue: string) => {
        let items: IAccordionListItem[] = state.listItems;
        let searchItems: IAccordionListItem[] = items.filter(o => {
            return o.Title.toLowerCase().indexOf(newValue.toLowerCase()) >= 0 ||
                o.Description.toLowerCase().indexOf(newValue.toLowerCase()) >= 0;
        });
        setState({
            ...state,
            items: searchItems.splice(0, props.maxItemsPerPage)
        });
    };

    useEffect(() => {
        _readItems();
    }, [props]);

    return (
        <div className={styles.reactAccordion}>
            <div className={styles.container}>
                <WebPartTitle displayMode={props.displayMode} title={props.title} updateProperty={props.updateProperty} />
                {!props.listId ? (
                    <>
                        {props.displayMode == DisplayMode.Edit ? (
                            <Placeholder iconName='Edit'
                                iconText={strings.ConfigLabel}
                                description={strings.ConfigDesc}
                                buttonLabel={strings.ConfigButton}
                                onConfigure={props.configurePropertyPane} />
                        ) : (
                                <Placeholder iconName='Edit'
                                    iconText={strings.ConfigLabel}
                                    description={strings.ConfigDesc} />
                            )}
                    </>
                ) : (
                        <>
                            {state.isLoading ? (
                                <ProgressIndicator label={strings.LoadingLabel} description={strings.LoadingDesc} />
                            ) : (
                                    <>
                                        <div className='ms-Grid-row'>
                                            <div className='ms-Grid-col ms-u-lg12'>
                                                <SearchBox onChange={_searchData} />
                                            </div>
                                        </div>
                                        {state.items.length <= 0 ? (
                                            <MessageBar messageBarType={MessageBarType.error}>{strings.NoData}</MessageBar>
                                        ) : (
                                                <>
                                                    <div className={`ms-Grid-row`}>
                                                        <div className='ms-Grid-col ms-u-lg12'>
                                                            <Accordion accordion={false}>
                                                                {items}
                                                            </Accordion>
                                                        </div>
                                                    </div>
                                                </>
                                            )}
                                    </>
                                )}
                        </>
                    )}
            </div>
        </div >
    );
};

export default class ReactAccordion extends React.Component<IReactAccordionProps, IReactAccordionState> {

    constructor(props: IReactAccordionProps, state: IReactAccordionState) {
        super(props);
        this.state = {
            status: this.listNotConfigured(this.props) ? 'Please configure list in Web Part properties' : 'Ready',
            items: [],
            listItems: [],
            isLoading: false,
            loaderMessage: ''
        };

        if (!this.listNotConfigured(this.props)) {
            this.readItems();
        }

        this.searchTextChange = this.searchTextChange.bind(this);

    }

    private listNotConfigured(props: IReactAccordionProps): boolean {
        return props.listId === undefined || props.listId === null || props.listId.length === 0;
    }

    private searchTextChange(event) {

        if (event === undefined ||
            event === null ||
            event === "") {
            let listItemsCollection = [...this.state.listItems];
            this.setState({ items: listItemsCollection.splice(0, this.props.maxItemsPerPage) });
        }
        else {
            var updatedList = [...this.state.listItems];
            updatedList = updatedList.filter((item) => {
                return item.Title.toLowerCase().search(
                    event.toLowerCase()) !== -1 || item.Description.toLowerCase().search(
                        event.toLowerCase()) !== -1;
            });
            this.setState({ items: updatedList });
        }
    }

    private readItems = async () => {
        let listItemsCollection: IAccordionListItem[] = await sp.web.lists.getById(this.props.listId).items
            .select('ID', 'Title', 'Description')
            .get();
        try {
            this.setState({
                status: "",
                items: listItemsCollection.splice(0, this.props.maxItemsPerPage),
                listItems: listItemsCollection,
                isLoading: false,
                loaderMessage: ""
            });
        } catch (err) {
            this.setState({
                status: 'Loading all items failed with error: ' + err,
                items: [],
                isLoading: false,
                loaderMessage: ""
            });
        }
    }

    public render(): React.ReactElement<IReactAccordionProps> {
        let displayLoader;
        let faqTitle;
        let { listItems } = this.state;
        let pageCountDivisor: number = this.props.maxItemsPerPage;
        let pageCount: number;
        let pageButtons = [];

        let _pagedButtonClick = (pageNumber: number, listData: any) => {
            let startIndex: number = (pageNumber - 1) * pageCountDivisor;
            let listItemsCollection = [...listData];
            this.setState({ items: listItemsCollection.splice(startIndex, pageCountDivisor) });
        };

        const items: JSX.Element[] = this.state.items.map((item: IAccordionListItem, i: number): JSX.Element => {
            return (
                <AccordionItem>
                    <AccordionItemTitle className="accordion__title">
                        <h3 className="u-position-relative ms-fontColor-white">{item.Title}</h3>
                        <div className="accordion__arrow ms-fontColor-white" role="presentation" />
                    </AccordionItemTitle>
                    <AccordionItemBody className="accordion__body">
                        <div className="" dangerouslySetInnerHTML={{ __html: item.Description }}>
                        </div>
                    </AccordionItemBody>
                </AccordionItem>
            );
        });

        if (this.state.isLoading) {
            displayLoader = (<div className={`ms-Grid-row ms-bgColor-themeDark ms-fontColor-white ${styles.row}`}>
                <div className='ms-Grid-col ms-u-lg12'>
                    <Spinner size={SpinnerSize.large} label={this.state.loaderMessage} />
                </div>
            </div>);
        }
        else {
            displayLoader = (null);
        }

        if (this.state.listItems.length > 0) {
            pageCount = Math.ceil(this.state.listItems.length / pageCountDivisor);
        }
        for (let i = 0; i < pageCount; i++) {
            pageButtons.push(<PrimaryButton onClick={() => { _pagedButtonClick(i + 1, listItems); }}> {i + 1} </PrimaryButton>);
        }
        return (
            <div className={styles.reactAccordion}>
                <div className={styles.container}>
                    <WebPartTitle displayMode={this.props.displayMode}
                        title={this.props.title}
                        updateProperty={this.props.updateProperty} />
                    {!this.props.listId ? (
                        <>
                            {this.props.displayMode == DisplayMode.Edit ? (
                                <Placeholder iconName='Edit'
                                    iconText={strings.ConfigLabel}
                                    description={strings.ConfigDesc}
                                    buttonLabel={strings.ConfigButton}
                                    onConfigure={this.props.configurePropertyPane} />
                            ) : (
                                    <Placeholder iconName='Edit'
                                        iconText={strings.ConfigLabel}
                                        description={strings.ConfigDesc} />
                                )}
                        </>
                    ) : (
                            <>
                                {/* {faqTitle}
                                {displayLoader} */}
                                <div className='ms-Grid-row'>
                                    <div className='ms-Grid-col ms-u-lg12'>
                                        <SearchBox onChange={this.searchTextChange} />
                                    </div>
                                </div>
                                <div className={`ms-Grid-row`}>
                                    <div className='ms-Grid-col ms-u-lg12'>
                                        {this.state.status}
                                        <Accordion accordion={false}>
                                            {items}
                                        </Accordion>
                                    </div>
                                </div>
                                <div className='ms-Grid-row'>
                                    <div className='ms-Grid-col ms-u-lg12'>
                                        {pageButtons}
                                    </div>
                                </div>
                            </>
                        )}
                </div>
            </div >
        );
    }
}
