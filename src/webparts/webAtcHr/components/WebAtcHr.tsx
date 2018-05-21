import * as React from 'react';
import styles from './WebAtcHr.module.scss';
import { IWebAtcHrProps } from './IWebAtcHrProps';
import { escape } from '@microsoft/sp-lodash-subset';
import Animate from 'react-simple-animate';
import { SPComponentLoader } from '@microsoft/sp-loader';
import { CommandButton, IButtonProps } from 'office-ui-fabric-react/lib/Button';
import { GridForm, Fieldset, Row, Field } from 'react-gridforms'
import * as Datetime from 'react-datetime';
import 'react-datetime/css/react-datetime.css';
import pnp from "sp-pnp-js";

import {
    CompactPeoplePicker,
    IBasePickerSuggestionsProps,
    NormalPeoplePicker
} from 'office-ui-fabric-react/lib/Pickers';
import { IPersonaProps } from 'office-ui-fabric-react/lib/Persona';
import {
    assign,
    autobind
} from 'office-ui-fabric-react/lib/Utilities';
import { IContextualMenuItem } from 'office-ui-fabric-react/lib/ContextualMenu';
import {
    SPHttpClient,
    SPHttpClientBatch,
    SPHttpClientResponse, SPHttpClientConfiguration
} from '@microsoft/sp-http';
import { IDigestCache, DigestCache } from '@microsoft/sp-http';
import {
    Environment,
    EnvironmentType
} from '@microsoft/sp-core-library';
import { Promise } from 'es6-promise';
import * as lodash from 'lodash';
import * as jquery from 'jquery';
import { IPersonaWithMenu } from 'office-ui-fabric-react/lib/components/pickers/PeoplePicker/PeoplePickerItems/PeoplePickerItem.Props';
import { IOfficeUiFabricPeoplePickerProps } from './IOfficeUiFabricPeoplePickerProps';
import { people } from './PeoplePickerExampleData';
import {
    IClientPeoplePickerSearchUser,
    IEnsurableSharePointUser,
    IEnsureUser,
    IOfficeUiFabricPeoplePickerState,
    SharePointUserPersona
} from '../models/OfficeUiFabricPeoplePicker';
export interface ISPlists {
    value: ISPList[];
}

export interface ISPList {
    Title: string;
    Id: string;
}


export interface RFIitems {
    Title: string;
    Id: string;
}

const suggestionProps: IBasePickerSuggestionsProps = {
    suggestionsHeaderText: 'Suggested People',
    noResultsFoundText: 'No results found',
    loadingText: 'Loading'
};


export default class WebAtcHr extends React.Component<IWebAtcHrProps, {}> {
    private _peopleList;
    private myhtpclient: SPHttpClient;

    public state: IWebAtcHrProps;
    constructor(props, context) {
        super(props);
        this.state = {
            description: "",
            PassportRequest: 0,
            LeaveRequest: 0,
            FormIsEnabled: 0,
            RequestTypeString: "",
            spHttpClient: this.props.spHttpClient,
            siteUrl: "https://arabtec.sharepoint.com/sites/dev/LMS",
            currentPicker: 1,
            delayResults: true,
            selectedItems: [],
            descriptionpicker: "",
            siteUrlpicker: "https://arabtec.sharepoint.com/sites/dev/LMS",
            typePicker: "",
            principalTypeUser: true,
            principalTypeSharePointGroup: true,
            principalTypeSecurityGroup: true,
            principalTypeDistributionList: true,
            numberOfItems: 5,
            EmployeeName: "",
            EmployeeNumber: "",
            EmployeeManager: "",
            EmployeeEmail: "",
            EmpFirstName: "",
            EmpLastName: "",
            EmpNumber: "",
            Description: "",
            FromDate: "",
            ToDate: "",
            FromCity: "",
            ToCity: "",
            StorageCapaity: "",
            LineManager: "",
            ManagerHead: "",
            Status: "",
            Stage: "",
            EmpEmirates: "",
            EmpPassportNumber: "",

        }
        SPComponentLoader.loadCss('https://cdnjs.cloudflare.com/ajax/libs/twitter-bootstrap/3.3.7/css/bootstrap.min.css');
    }

    public onSelectDateFrom(event: any): void {
        //this.setState({ fromdt: event._d });
    }
    public onSelectDateTo(event: any): void {
        // this.setState({ todt: event._d });
    }

    public CreateNewItem(event: any): void {
        
        var dateFormat = require('dateformat');
        //var FinalDate = dateFormat(this.state.RequiredDate, "m/dd/yyyy");
        var NewISiteUrl = this.props.siteUrl;
        var NewSiteUrl = NewISiteUrl.replace("/SitePages", "");

        pnp.sp.web.lists.getByTitle("Human Resource").items.add({
            Title: "My Item Title"
        }).then(r => {
            console.log(r);
            //r.item.attachmentFiles.add("file.txt", "Here is some file content.");
        });
        
       // let webx = new Web(NewSiteUrl);
       // webx.lists.getByTitle("RFI").items.add({
         // Title: this.state.ItemGuid,
         // Main_x0020_Comments: this.refs.RefComments["innerHTML"],
         // RFI_x0020_Subject: this.state.Subject,
         // Building: this.state.Building,
         // Date: FinalDate,
       // }).then((iar: ItemAddResult) => {
         // let CurUser = iar.data["AuthorId"];
         // this.UpdateSeriesItemIRFI(iar.data["Id"]);
       // });
    }

    public CloseGrid(event: any): void {
        this.setState({
            FormIsEnabled: 0
        });
    }

    componentDidMount() {
        this.GetUSerDetails();
    }

    private GetUSerDetails() {
        var reactHandler = this;
        var NewISiteUrl = this.props.siteUrl;
        var NewSiteUrl = NewISiteUrl.replace("/SitePages", "");
        var reqUrl = NewSiteUrl + "/_api/sp.userprofiles.peoplemanager/GetMyProperties";
        jquery.ajax(
            {
                url: reqUrl, type: "GET", headers:
                    {
                        "accept": "application/json;odata=verbose"
                    }
            }).then((response) => {
                //console.log(response.data);
                var Name = response.d.DisplayName;
                var email = response.d.Email;
                var oneUrl = response.d.PersonalUrl;
                var imgUrl = response.d.PictureUrl;
                var jobTitle = response.d.Title;
                var profUrl = response.d.UserUrl;
                var MBNumber = response.d.AccountName;
                var Tmpe = MBNumber.toString().split('|');
                var Tmp2 = Tmpe[2].toString().split('@');
                MBNumber = Tmp2[0];
                reactHandler.setState({
                    EmployeeName: response.d.DisplayName,
                    EmployeeNumber: MBNumber
                });
            });
    }




    public DivSelected(reqteype) {
        switch (reqteype) {
            case 1:
                this.setState({
                    FormIsEnabled: 1,
                    RequestTypeString: "Leave Request",
                    PassportRequest: 0,
                    LeaveRequest: 1
                })
                break;

            case 2:
                this.setState({
                    FormIsEnabled: 1,
                    RequestTypeString: "Passport Request",
                    PassportRequest: 1,
                    LeaveRequest: 0,

                })
                break;


            case 3:
                this.setState({
                    FormIsEnabled: 1,
                    RequestTypeString: "Server Request",
                })
                break;


            case 4:
                this.setState({
                    FormIsEnabled: 1,
                    RequestTypeString: "Sick Request",
                })
                break;

            case 5:
                this.setState({
                    FormIsEnabled: 1,
                    RequestTypeString: "Allowance Request",
                })
                break;

            case 6:
                this.setState({
                    FormIsEnabled: 1,
                    RequestTypeString: "User Request",
                })
                break;


        }
    }


    public render(): React.ReactElement<IWebAtcHrProps> {
        return (
            <div className={styles.webAtcHr} >
                {
                    this.state.FormIsEnabled == 0 &&
                    <div>
                        <div className={styles.containerLeave} onClick={this.DivSelected.bind(this, 1)}>

                            <img src="https://www.healthline.com/hlcmsresource/images/News/6109-Man_Office-Sick-1296x728-Header.jpg" className={styles.imageLeave} />
                            <div className={styles.overlayLeave}>
                                <div className={styles.pragraphLeave}>Leave Request</div>
                            </div>
                        </div>
                        <div className={styles.containerPassport} onClick={this.DivSelected.bind(this, 2)}>
                            <img src="https://robinpowered.com/blog/wp-content/uploads/2018/01/Be-a-Better-Office-during-the-Season-of-Sick-Days.jpg" className={styles.imagePassport} />
                            <div className={styles.overlayPassport}>
                                <div className={styles.pragraphPassport}>Passport Request</div>
                            </div>
                        </div>
                        <div className={styles.containerServer} onClick={this.DivSelected.bind(this, 3)}>
                            <img src="https://trott.house.gov/sites/trott.house.gov/files/styles/congress_featured_image/public/featured_image/office_location/Office-Door-1Small.jpg?itok=VkWcZOnr" className={styles.imageServer} />
                            <div className={styles.overlayServer}>
                                <div className={styles.pragraphServer}>Server Request</div>
                            </div>
                        </div>
                        <div className={styles.containerSick} onClick={this.DivSelected.bind(this, 4)}>
                            <img src="https://www.myledlightingguide.com/images/installations/Office.jpg" className={styles.imageSick} />
                            <div className={styles.overlaySick}>
                                <div className={styles.pragraphSick}>Sick Request</div>
                            </div>
                        </div>
                        <div className={styles.containerAllowance} onClick={this.DivSelected.bind(this, 5)}>
                            <img src="http://under30ceo.com/wp-content/uploads/2010/10/home-office-2-582x379.jpg" className={styles.imageAllowance} />
                            <div className={styles.overlayAllowance}>
                                <div className={styles.pragraphAllowance}>Allowance Request</div>
                            </div>
                        </div>
                        <div className={styles.containerUser} onClick={this.DivSelected.bind(this, 6)}>
                            <img src="http://msofficeuser.com/pages/wp-content/uploads/2011/01/HelpHeader.png" className={styles.imageUser} />
                            <div className={styles.overlayUser}>
                                <div className={styles.pragraphUser}>User Request</div>
                            </div>
                        </div>
                    </div>
                }
                {this.state.FormIsEnabled == 1 &&
                    <div className={styles.HeaderGrid}>

                        <GridForm>

                            <Fieldset legend={this.state.RequestTypeString}>
                                <Row>
                                    <Field span={3}>
                                        <label>Employee Name</label>
                                        {this.state.EmployeeName}
                                    </Field>
                                    <Field>
                                        <label>MB Number</label>
                                        {this.state.EmployeeNumber.toLocaleUpperCase()}
                                    </Field>
                                </Row>
                            </Fieldset>
                        </GridForm>
                    </div>
                }

                {this.state.LeaveRequest == 1 && this.state.FormIsEnabled == 1 &&
                    <div className={styles.HeaderGrid}>
                        <GridForm>
                            <Row>
                                <Field span={2}>
                                    <label>From Date</label>
                                    <Datetime onChange={this.onSelectDateFrom.bind(this)} />
                                </Field>
                                <Field span={2}>
                                    <label>To Date</label>
                                    <Datetime onChange={this.onSelectDateTo.bind(this)} />
                                </Field>
                            </Row>
                            <Row>
                                <Field span={4}>
                                    <label>Description</label>
                                    <input type="text" className={styles.myinput} />
                                </Field>

                            </Row>
                        </GridForm>
                    </div>
                }

                {this.state.PassportRequest == 1 && this.state.FormIsEnabled == 1 &&
                    <div className={styles.HeaderGrid}>
                        <GridForm>
                            <Row>
                                <Field span={4}>
                                    <label>Description</label>
                                    <input type="text" className={styles.myinput} />
                                </Field>

                            </Row>
                        </GridForm>
                    </div>
                }


                {this.state.FormIsEnabled == 1 &&
                    <div className={styles.HeaderGrid}>
                            <Row>
                                <Field span={3}>
                                    <label>Manager-To-Approval</label>
                                    <NormalPeoplePicker
                                        onChange={this._onChange.bind(this)}
                                        onResolveSuggestions={this._onFilterChanged}
                                        getTextFromItem={(persona: IPersonaProps) => persona.primaryText}
                                        pickerSuggestionsProps={suggestionProps}
                                        className={'ms-PeoplePicker'}
                                        key={'normal'}
                                    />
                                </Field>
                            </Row>
                            <Row>
                                <Field span={3}>
                                <div className={styles.FooterButtonDiv}>
                                    <button id="btn_add" className={styles.MainButton} onClick={this.CreateNewItem.bind(this)}>Create Request </button>
                                    <button id="btn_add" className={styles.MyButton} onClick={this.CloseGrid.bind(this)}>Close </button>
                                    </div>
                                </Field>
                            </Row>
                    </div>
                }



            </div>
        );
    }

    private _onChange(items: any[]) {
        this.setState({
            selectedItems: items
        });
        if (this.props.onChange) {
            this.props.onChange(items);
        }
    }
    @autobind
    private _onFilterChanged(filterText: string, currentPersonas: IPersonaProps[], limitResults?: number) {
        if (filterText) {
            if (filterText.length > 2) {
                return this._searchPeople(filterText, this._peopleList);
            }
        } else {
            return [];
        }
    }


    private searchPeopleFromMock(): IPersonaProps[] {
        return this._peopleList = [
            {
                imageUrl: './images/persona-female.png',
                imageInitials: 'PV',
                primaryText: 'Annie Lindqvist',
                secondaryText: 'Designer',
                tertiaryText: 'In a meeting',
                optionalText: 'Available at 4:00pm'
            },
            {
                imageUrl: './images/persona-male.png',
                imageInitials: 'AR',
                primaryText: 'Aaron Reid',
                secondaryText: 'Designer',
                tertiaryText: 'In a meeting',
                optionalText: 'Available at 4:00pm'
            },
            {
                imageUrl: './images/persona-male.png',
                imageInitials: 'AL',
                primaryText: 'Alex Lundberg',
                secondaryText: 'Software Developer',
                tertiaryText: 'In a meeting',
                optionalText: 'Available at 4:00pm'
            },
            {
                imageUrl: './images/persona-male.png',
                imageInitials: 'RK',
                primaryText: 'Roko Kolar',
                secondaryText: 'Financial Analyst',
                tertiaryText: 'In a meeting',
                optionalText: 'Available at 4:00pm'
            },
        ];
    }
    private _searchPeople(terms: string, results: IPersonaProps[]): IPersonaProps[] | Promise<IPersonaProps[]> {

        if (DEBUG && Environment.type === EnvironmentType.Local) {
            // If the running environment is local, load the data from the mock
            return this.searchPeopleFromMock();
        } else {
            const userRequestUrl: string = `https://arabtec.sharepoint.com/sites/dev/LMS/_api/SP.UI.ApplicationPages.ClientPeoplePickerWebServiceInterface.clientPeoplePickerSearchUser`;
            let principalType: number = 0;
            if (this.props.principalTypeUser === true) {
                principalType += 1;
            }
            if (this.props.principalTypeSharePointGroup === true) {
                principalType += 8;
            }
            if (this.props.principalTypeSecurityGroup === true) {
                principalType += 4;
            }
            if (this.props.principalTypeDistributionList === true) {
                principalType += 2;
            }
            const userQueryParams = {
                'queryParams': {
                    'AllowEmailAddresses': true,
                    'AllowMultipleEntities': false,
                    'AllUrlZones': false,
                    'MaximumEntitySuggestions': 5,
                    'PrincipalSource': 15,
                    // PrincipalType controls the type of entities that are returned in the results.
                    // Choices are All - 15, Distribution List - 2 , Security Groups - 4, SharePoint Groups - 8, User - 1.
                    // These values can be combined (example: 13 is security + SP groups + users)
                    'PrincipalType': 1,
                    'QueryString': terms
                }
            };

            return new Promise<SharePointUserPersona[]>((resolve, reject) =>
                this.props.spHttpClient.post(userRequestUrl,
                    SPHttpClient.configurations.v1, { body: JSON.stringify(userQueryParams) })
                    .then((response: SPHttpClientResponse) => {
                        return response.json();
                    })
                    .then((response: { value: string }) => {
                        let userQueryResults: IClientPeoplePickerSearchUser[] = JSON.parse(response.value);
                        let persons = userQueryResults.map(p => new SharePointUserPersona(p as IEnsurableSharePointUser));
                        return persons;
                    })
                    .then((persons) => {
                        const batch = this.props.spHttpClient.beginBatch();
                        const ensureUserUrl = `https://arabtec.sharepoint.com/sites/dev/LMS/_api/web/ensureUser`;
                        const batchPromises: Promise<IEnsureUser>[] = persons.map(p => {
                            var userQuery = JSON.stringify({ logonName: p.User.Key });
                            return batch.post(ensureUserUrl, SPHttpClientBatch.configurations.v1, {
                                body: userQuery
                            })
                                .then((response: SPHttpClientResponse) => response.json())
                                .then((json: IEnsureUser) => json);
                        });

                        var User: string = "";
                        var users = batch.execute().then(() => Promise.all(batchPromises).then(values => {
                            values.forEach(v => {
                                let userPersona = lodash.find(persons, o => o["User"].Key == v.LoginName);
                                if (userPersona && userPersona["User"]) {
                                    let user = userPersona["User"];
                                    lodash.assign(user, v);
                                    userPersona["User"] = user;
                                }
                            });

                            resolve(persons);
                        }));
                    }, (error: any): void => {
                        reject(this._peopleList = []);
                    }));
        }
    }
    private _filterPersonasByText(filterText: string): IPersonaProps[] {
        return this._peopleList.filter(item => this._doesTextStartWith(item.primaryText, filterText));
    }
    private _removeDuplicates(personas: IPersonaProps[], possibleDupes: IPersonaProps[]) {
        return personas.filter(persona => !this._listContainsPersona(persona, possibleDupes));
    }
    private _listContainsPersona(persona: IPersonaProps, personas: IPersonaProps[]) {
        if (!personas || !personas.length || personas.length === 0) {
            return false;
        }
        return personas.filter(item => item.primaryText === persona.primaryText).length > 0;
    }
    private _filterPromise(personasToReturn: IPersonaProps[]): IPersonaProps[] | Promise<IPersonaProps[]> {
        if (this.state.delayResults) {
            return this._convertResultsToPromise(personasToReturn);
        } else {
            return personasToReturn;
        }
    }
    private _convertResultsToPromise(results: IPersonaProps[]): Promise<IPersonaProps[]> {
        return new Promise<IPersonaProps[]>((resolve, reject) => setTimeout(() => resolve(results), 2000));
    }
    private _doesTextStartWith(text: string, filterText: string): boolean {
        return text.toLowerCase().indexOf(filterText.toLowerCase()) === 0;
    }
    //picker function ends
}