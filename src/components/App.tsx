import * as React from 'react';
import { Button, ButtonType } from 'office-ui-fabric-react';
import Header from './Header';
import HeroList, { HeroListItem } from './HeroList';
import Progress from './Progress';
import * as OfficeHelpers from '@microsoft/office-js-helpers';


export interface AppProps {
    title: string;
    isOfficeInitialized: boolean;
}

export interface AppState {
    listItems: HeroListItem[];
}

export default class App extends React.Component<AppProps, AppState> {
    constructor(props, context) {
        super(props, context);
        this.state = {
            listItems: []
        };

        Office.select("binding#1").addHandlerAsync(Office.EventType.BindingSelectionChanged,this._controlSelected);
            
        Office.select("binding#1").addHandlerAsync(Office.EventType.BindingDataChanged, this._controlUpdated);
    }

    componentDidMount() {
        this.setState({
            listItems: [
                {
                    icon: 'Ribbon',
                    primaryText: 'Achieve more with Office integration'
                },
                {
                    icon: 'Unlock',
                    primaryText: 'Unlock features and functionality'
                },
                {
                    icon: 'Design',
                    primaryText: 'Create and visualize like a pro'
                }
            ]
        });
    }

    click = () => {
        Office.context.document.bindings.addFromSelectionAsync(Office.BindingType.Text, { id: '1' },
            (asyncResult: Office.AsyncResult) => {
                if (asyncResult.status == Office.AsyncResultStatus.Failed) {
                    OfficeHelpers.UI.notify(asyncResult.error);
                }
            });
            
            Office.select("binding#1").setDataAsync("hello from TaskPane")

            Office.select("binding#1").addHandlerAsync(Office.EventType.BindingSelectionChanged,this._controlSelected);
            
            Office.select("binding#1").addHandlerAsync(Office.EventType.BindingDataChanged, this._controlUpdated);
    }

    _controlSelected = () => {
        OfficeHelpers.UI.notify("Binding Selected");
    }

    _controlUpdated = () => {
        OfficeHelpers.UI.notify("Binding Value Updated");
    }

    render() {
        const {
            title,
            isOfficeInitialized,
        } = this.props;

        if (!isOfficeInitialized) {
            return (
                <Progress
                    title={title}
                    logo='assets/logo-filled.png'
                    message='Please sideload your addin to see app body.'
                />
            );
        }

        return (
            <div className='ms-welcome'>
                <Header logo='assets/logo-filled.png' title={this.props.title} message='Welcome' />
                <HeroList message='Discover what test_addin can do for you today!' items={this.state.listItems}>
                    <p className='ms-font-l'>Modify the source files, then click <b>Run</b>.</p>
                    <Button className='ms-welcome__action' buttonType={ButtonType.hero} iconProps={{ iconName: 'ChevronRight' }} onClick={this.click}>Run</Button>
                </HeroList>
            </div>
        );
    }
}
