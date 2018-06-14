import * as React from 'react';
import styles from './WebAtcHr.module.scss';
import { IWebAtcHrProps } from './IWebAtcHrProps';
import { escape } from '@microsoft/sp-lodash-subset';
import Animate from 'react-simple-animate';
import { SPComponentLoader } from '@microsoft/sp-loader';
import { CommandButton, IButtonProps } from 'office-ui-fabric-react/lib/Button';
import { GridForm, Fieldset, Row, Field } from 'react-gridforms'
import { Button, Modal } from 'react-bootstrap';
import { default as pnp, ItemAddResult } from "sp-pnp-js";
import * as $ from 'jquery';
var moment = require('moment');

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
import * as Datetime from 'react-datetime';
import 'react-datetime/css/react-datetime.css';
import Moment from 'react-moment';

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
    protected onInit(): Promise<void> {
        return new Promise<void>((resolve: () => void, reject: (error?: any) => void): void => {
            pnp.setup({
                sp: {
                    headers: {
                        "Accept": "application/json; odata=nometadata"
                    }
                }
            });
            resolve();
        });
    }
    private _peopleList;
    private myhtpclient: SPHttpClient;

    public state: IWebAtcHrProps;
    constructor(props, context) {
        super(props);
        this.state = {
            DetailComments: "",
            description: "",
            PassportRequest: 0,
            LeaveRequest: 0,
            AirTicketRequest: 0,
            FormIsEnabled: 0,
            RequestTypeString: "",
            spHttpClient: this.props.spHttpClient,
            siteUrl: "https://arabtec.sharepoint.com",
            currentPicker: 1,
            delayResults: true,
            selectedItems: [],
            descriptionpicker: "",
            siteUrlpicker: "https://arabtec.sharepoint.com",
            typePicker: "",
            principalTypeUser: true,
            principalTypeSharePointGroup: true,
            principalTypeSecurityGroup: true,
            principalTypeDistributionList: true,
            numberOfItems: 5,
            EmployeeName: "NA",
            EmployeeNumber: "NA",
            EmployeeManager: "NA",
            EmployeeEmail: "NA",
            EmpFirstName: "NA",
            EmpLastName: "NA",
            EmpNumber: "NA",
            Description: "NA",
            FromDate: "NA",
            ToDate: "NA",
            FromCity: "NA",
            ToCity: "NA",
            StorageCapaity: "NA",
            LineManager: "",
            ManagerHead: "",
            Status: "Pending",
            Stage: "",
            EmpEmirates: "NA",
            EmpPassportNumber: "NA",
            IsFormReadOnly: false,
            RequestType: "NA",
            SucessFullModal: false,
            ErrorModal: false,
            ItemId: "",
            ItemStatus: "",
            ManagerApprovalComments: "",
            TotalDays:"Total Days  - 0",
            ApprovalStatus:"Pending",

        }
        SPComponentLoader.loadCss('https://cdnjs.cloudflare.com/ajax/libs/twitter-bootstrap/3.3.7/css/bootstrap.min.css');
    }

    public onSelectDateFrom(event: any): void {

        this.setState({ FromDate: event._d });
    }
    public onSelectDateTo(event: any): void {
        this.setState({ ToDate: event._d });
    }

    public CreateNewItem(event: any): void {

        var dateFormat = require('dateformat');
        var NewISiteUrl = this.props.siteUrl;
        var NewSiteUrl = NewISiteUrl.replace("/SitePages", "");
        pnp.sp.web.lists.getByTitle("Human%20Resource").items.add({
            Title: "My Item Title",
            FromDate: this.state.FromDate == 'NA' ? null : this.state.FromDate,
            ToDate: this.state.ToDate == 'NA' ? null : this.state.ToDate,
            FromCity: this.state.FromCity,
            ToCity: this.state.ToCity,
            EmpFirstName: this.state.EmployeeName,
            EmpLastName: this.state.EmployeeName,
            EmpNumber: this.state.EmployeeNumber,
            ItemsComments: this.state.DetailComments,
            StorageCapaity: this.state.StorageCapaity,
            Status: "Pending",
            Stage: "User",
            EmpPassportNumber: this.state.EmpPassportNumber,
            LineManagerId: this.state.selectedItems[0]["_user"]["Id"].toString(),
            LineManagerEmail: this.state.selectedItems[0]["_user"]["Email"],
            RequestType: this.state.RequestType,
        }).then((response) => {
            this.setState({ SucessFullModal: true });
            this.setState({ EmployeeManager: "NA" });
            this.setState({ EmployeeEmail: "NA" });
            this.setState({ EmpFirstName: "NA" });
            this.setState({ EmpLastName: "NA" });
            this.setState({ EmpNumber: "NA" });
            this.setState({ Description: "NA" });
            this.setState({ FromDate: "NA" });
            this.setState({ ToDate: "NA" });
            this.setState({ FromCity: "NA" });
            this.setState({ ToCity: "NA" });
            this.setState({ StorageCapaity: "NA" });
            this.setState({ LineManager: "" });
            this.setState({ EmpEmirates: "NA" });
            this.setState({ EmpPassportNumber: "NA" });
            this.setState({ selectedItems: [] });
            this.setState({ DetailComments: '' });

        }).catch(function (data) {
            this.setState({ ErrorModal: true });
            // Handle error here
        });
    }

    CloseGrid(event: any): void {
        this.setState({ EmployeeManager: "NA" });
        this.setState({ EmployeeEmail: "NA" });
        this.setState({ EmpFirstName: "NA" });
        this.setState({ EmpLastName: "NA" });
        this.setState({ EmpNumber: "NA" });
        this.setState({ Description: "NA" });
        this.setState({ FromDate: "NA" });
        this.setState({ ToDate: "NA" });
        this.setState({ FromCity: "NA" });
        this.setState({ ToCity: "NA" });
        this.setState({ StorageCapaity: "NA" });
        this.setState({ LineManager: "" });
        this.setState({ EmpEmirates: "NA" });
        this.setState({ EmpPassportNumber: "NA" });
        this.setState({ selectedItems: [] });
        this.setState({ DetailComments: '' });

        this.setState({
            FormIsEnabled: 0
        });
    }

    public OnChangeDescription(event: any): void {
        this.setState({
            DetailComments: event.target.value
        });
    }

    public OnChangeApprovalComment(event: any): void {
        this.setState({
            ManagerApprovalComments: event.target.value
        });
    }

    componentDidMount() {
        this.GetUSerDetails();
        this.readUrl();


    }
    /* Edit Page Dtails ----------------------------------------------- START     */
    public readUrl() {
        var url = window.location.href;
        if (url.lastIndexOf('=') > -1) {
            var id = url.substring(url.lastIndexOf('=') + 1);
            this.setState({
                IsFormReadOnly: true,
                FormIsEnabled: 1,
                ItemId: id,
            });
            this.fetchitem(id);
        }
    }

    public fetchitem(id) {
        var NewISiteUrl = this.props.siteUrl;
        var NewSiteUrl = NewISiteUrl.replace("/SitePages", "");
        jquery.ajax({
            url: `${NewSiteUrl}/_api/web/lists/getbytitle('Human%20Resource')/items?$filter=ID%20eq%20${id}%20&$select=Title,ID,EmpFirstName,EmpNumber,FromDate,ToDate,FromCity,ToCity,StorageCapaity,Status,Stage,EmpEmirates,EmpPassportNumber,ItemsComments,RequestType`,
            type: "GET",
            headers: { 'Accept': 'application/json; odata=verbose;' },
            success: function (resultData) {
                var myObject = JSON.stringify(resultData.d.results);
                this.DivSelected(resultData.d.results[0]["RequestType"]);
                var TempFromDate=resultData.d.results[0]["FromDate"];
                var dateFormat = require('dateformat');
                var FinalDate = dateFormat(TempFromDate, "dd-mm-yyyy");
                var TempToDate= resultData.d.results[0]["ToDate"];
                var FinalDate2 = dateFormat(TempToDate, "dd-mm-yyyy");              
                this.GettheCount(FinalDate,FinalDate2);
                this.setState({
                    EmployeeName: resultData.d.results[0]["EmpFirstName"],
                    EmployeeNumber: resultData.d.results[0]["EmpNumber"],
                    IsFormReadOnly: false,
                    FromDate:FinalDate,
                    ToDate:FinalDate2,                    
                    DetailComments: resultData.d.results[0]["ItemsComments"],
                    Status: resultData.d.results[0]["Status"],
                });

                this.setState({
                    IsFormReadOnly: true,
                });

            }.bind(this),
            error: function (jqXHR, textStatus, errorThrown) {
            }
        });
    }
    public updateitems() {
        var id=this.state.ItemId;
        var Id =0;
        Id=parseInt(this.state.ItemId);
        
        if (this.state.ApprovalStatus!="Approve" && this.state.ApprovalStatus!="Reject"){
            this.setState({ApprovalStatus:"Pending"});
        }

        var TempManagerID="";
        var TempManagerEmail="";
        if (this.state.selectedItems.length>0)
        {
            TempManagerID=this.state.selectedItems[0]["_user"]["Id"].toString();
            TempManagerEmail=this.state.selectedItems[0]["_user"]["Email"].toString();
        }
        let list = pnp.sp.web.lists.getByTitle("Human%20Resource");
        list.items.getById(Id).update({
            Title: "Item has beeen Approved",
            Description: "Here is a new description",
            Status:this.state.ApprovalStatus,
            Stage:"Line Manager",
            ManagerHeadId:TempManagerID  ,
            ManagerHeadEmail:TempManagerEmail ,
            ApprovaqlComments:this.state.ManagerApprovalComments,
        }).then(i => {
            this.setState({ SucessFullModal: true });
        });
        
        
    }

    public GettheCount(stringfrmdt,stringtodt)
    {
        var Temp=stringfrmdt.split('-');
        var FinalFromdt=Temp[0]+"."+Temp[1]+"."+Temp[2];
        Temp=stringtodt.split('-');
        var FinalTodt2=Temp[0]+"."+Temp[1]+"."+Temp[2];
        var tmpstartDate = moment(FinalFromdt, "DD.MM.YYYY");
        var tmpendDate = moment( FinalTodt2, "DD.MM.YYYY");

        var result = "Total Days - "+tmpendDate.diff(tmpstartDate, 'days');
        this.setState({
            TotalDays:result
        });
        
    }


    /* Edit Page Dtails ----------------------------------------------- END     */
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
                    LeaveRequest: 1,
                    RequestType: "Leave Request",
                })
                break;

            case "Leave Request":
                this.setState({
                    FormIsEnabled: 1,
                    RequestTypeString: "Leave Request",
                    PassportRequest: 0,
                    LeaveRequest: 1,
                    RequestType: "Leave Request",
                })
                break;

            case 2:
                this.setState({
                    FormIsEnabled: 1,
                    RequestTypeString: "Passport Request",
                    PassportRequest: 1,
                    LeaveRequest: 0,
                    RequestType: "Passport Request",

                })
                break;

            case "Passport Request":
                this.setState({
                    FormIsEnabled: 1,
                    RequestTypeString: "Passport Request",
                    PassportRequest: 1,
                    LeaveRequest: 0,
                    AirTicketRequest: 0,
                    RequestType: "Passport Request",

                })
                break;

            case 3:
                this.setState({
                    FormIsEnabled: 1,
                    RequestTypeString: "Server Request",
                    RequestType: "Server Request",
                })
                break;
            case "Server Request":
                this.setState({
                    FormIsEnabled: 1,
                    RequestTypeString: "Server Request",
                    RequestType: "Server Request",
                })
                break;


            case 4:
                this.setState({
                    FormIsEnabled: 1,
                    RequestTypeString: "Sick Request",
                    RequestType: "Sick Request",
                })
                break;

            case "Sick Request":
                this.setState({
                    FormIsEnabled: 1,
                    RequestTypeString: "Sick Request",
                    RequestType: "Sick Request",
                })
                break;

            case 5:
                this.setState({
                    FormIsEnabled: 1,
                    RequestTypeString: "Allowance Request",
                    RequestType: "Allowance Request",
                })
                break;

            case "Allowance Request":
                this.setState({
                    FormIsEnabled: 1,
                    RequestTypeString: "Allowance Request",
                    RequestType: "Allowance Request",
                })
                break;

            case 6:
                this.setState({
                    FormIsEnabled: 1,
                    RequestTypeString: "User Request",
                    RequestType: "User Request",
                })
                break;

            case "User Request":
                this.setState({
                    FormIsEnabled: 1,
                    RequestTypeString: "User Request",
                    RequestType: "User Request",
                })
                break;

            case "Air Ticket":
                this.setState({
                    FormIsEnabled: 1,
                    PassportRequest: 0,
                    LeaveRequest: 0,
                    AirTicketRequest: 1,
                    RequestTypeString: "Air Ticket",
                    RequestType: "Air Ticket",
                })
                break;

            case 7:
                this.setState({
                    FormIsEnabled: 1,
                    RequestTypeString: "Air Ticket",
                    RequestType: "Air Ticket",
                })
                break;

        }
    }

    public className() {
        if (this.state.IsFormReadOnly == true) {
            return styles.ApprovalGrid;
        } else {
            return styles.HeaderGrid;
        }
    }

    CloseSucessFullModal(e) {
        this.setState({ SucessFullModal: false });
        return false;
    }
    CloseErrorModal(e) {
        this.setState({ CloseErrorModal: false });
        return false;
    }

    /* Radio Button */
    handleChange(event) {
        this.setState({
            ApprovalStatus: event.target.value
        });
      }
    /*Radio Button End */

    public render(): React.ReactElement<IWebAtcHrProps> {
        return (
            <div className={styles.webAtcHr} >
                {
                    this.state.FormIsEnabled == 0 && this.state.IsFormReadOnly == false &&
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
                        <div className={styles.containerAirTicket} onClick={this.DivSelected.bind(this, 7)}>
                            <img src="http://www.bokapowell.com/wp-content/uploads/2015/01/SWA-008.jpg" className={styles.imageUser} />
                            <div className={styles.overlayAirTicket}>
                                <div className={styles.pragraphAirTicket}>Air Ticket</div>
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
                            <Row >
                                <Field span={2} >
                                    <label>From Date</label>
                                    <Datetime dateFormat="DD-MM-YYYY" timeFormat={false}  value={this.state.FromDate} inputProps={{ disabled: this.state.IsFormReadOnly }} onChange={this.onSelectDateFrom.bind(this)} />
                                </Field>
                                <Field span={2} >
                                    <label>To Date</label>
                                    <Datetime dateFormat="DD-MM-YYYY" timeFormat={false}  value={this.state.ToDate} inputProps={{ disabled: this.state.IsFormReadOnly }} onChange={this.onSelectDateTo.bind(this)} />
                                </Field>
                            </Row>
                            <Row>
                                <Field span={4}>
                                    <label>Description</label>
                                    <input type="text" disabled={this.state.IsFormReadOnly} className={styles.myinput} value={this.state.DetailComments} onChange={this.OnChangeDescription.bind(this)} />
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
                                    <input type="text" className={styles.myinput} disabled={this.state.IsFormReadOnly} value={this.state.DetailComments} onChange={this.OnChangeDescription.bind(this)} />
                                </Field>

                            </Row>
                        </GridForm>
                    </div>
                }


                {this.state.AirTicketRequest == 1 && this.state.FormIsEnabled == 1 &&
                    <div className={styles.HeaderGrid}>
                        <GridForm>
                            <Row>
                                <Field span={4}>
                                    <label>Description</label>
                                    <input type="text" className={styles.myinput} disabled={this.state.IsFormReadOnly} value={this.state.DetailComments} onChange={this.OnChangeDescription.bind(this)} />
                                </Field>

                            </Row>
                        </GridForm>
                    </div>
                }


              
                {this.state.IsFormReadOnly == true && this.state.Status == "Pending" &&
                    <div className={styles.ApprovalGrid}>
                        <GridForm>
                            <Row>
                                <Field span={4}>
                                    <label>Action {this.state.TotalDays}</label>
                                    
                                </Field>
                            </Row>
                            <Row >
                                <Field span={2} >
                                    <label>Apprvoe</label>
                                    <input type="radio" value="Approve" name="gender"  onChange={this.handleChange.bind(this)} />
                                </Field>
                                <Field span={2} >
                                    <label>Reject</label>
                                    <input type="radio" value="Reject" name="gender"  onChange={this.handleChange.bind(this)} />
                                </Field>
                            </Row>
                        </GridForm>
                    </div>

                }
                {this.state.FormIsEnabled == 1 &&  this.state.Status == "Pending" &&
                    <div className={this.className()}>
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
                    </div>
                }

                {this.state.IsFormReadOnly == true && this.state.Status == "Pending" &&
                    <div className={styles.ApprovalGrid}>
                        <GridForm>
                            <Row>
                                <Field span={3}>
                                    <label>Approval Comment</label>
                                    <input type="text" className={styles.myinput}  value={this.state.ManagerApprovalComments} onChange={this.OnChangeApprovalComment.bind(this)} />
                                </Field>
                            </Row>
                            
                        </GridForm>
                        <Row>
                                <Field span={3}>
                                    <div className={styles.ApprovalGrid}>
                                    <br/>
                                        <button id="btn_add" className={'btn btn-primary'} onClick={this.updateitems.bind(this)}>Update </button>
                                    </div>
                                </Field>
                            </Row>
                    </div>
                }

                  {this.state.FormIsEnabled == 1 && this.state.IsFormReadOnly == false &&
                    <div className={styles.HeaderGrid}>
                        <Row>
                            <Field span={3}>
                                <div className={styles.FooterButtonDiv}>
                                    <button id="btn_add" className={'btn btn-primary'} onClick={this.CreateNewItem.bind(this)}>Create Request </button>
                                    &nbsp;
                                    <button id="btn_close" className={'btn btn-success'} onClick={this.CloseGrid.bind(this)} >Close</button>
                                </div>
                            </Field>
                        </Row>
                    </div>

                }


                 {this.state.IsFormReadOnly == true && this.state.Status != "Pending" &&
                    <div className={styles.ApprovalGrid}>
                        <GridForm>
                            <Row>
                                <Field span={3}>
                                <h1>{this.state.Status}</h1>
                                </Field>
                            </Row>
                            
                        </GridForm>
                       
                    </div>
                }



                <Modal show={this.state.SucessFullModal} >
                    <Modal.Body>
                        <div className="alert alert-success">
                            <strong>Success!</strong>
                        </div>
                    </Modal.Body>
                    <Modal.Footer>
                        <button type="button" onClick={this.CloseSucessFullModal.bind(this)} className="btn btn-default" data-dismiss="modal">Close</button>
                    </Modal.Footer>
                </Modal>
                <Modal show={this.state.ErrorModal} >
                    <Modal.Body>
                        <div className="alert alert-danger">
                            <strong>Success!</strong>
                        </div>
                    </Modal.Body>
                    <Modal.Footer>
                        <button type="button" onClick={this.CloseErrorModal.bind(this)} className="btn btn-default" data-dismiss="modal">Close</button>
                    </Modal.Footer>
                </Modal>
            </div>
        );
    }
    private _onChange(items: any[]) {
        if (items != null && items != undefined) {
            if (items.length > 0) {
                var Temp = items[0]._user.Description;
                items[0]._user.Id
                items[0]._user.Description
                this.setState({
                    selectedItems: items,
                    LineManager: Temp

                });
            }
        }
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
        return this._peopleList = [{
            imageUrl: './images/persona-female.png',
            imageInitials: 'PV',
            primaryText: 'Annie Lindqvist',
            secondaryText: 'Designer',
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
            const userRequestUrl: string = `https://arabtec.sharepoint.com/_api/SP.UI.ApplicationPages.ClientPeoplePickerWebServiceInterface.clientPeoplePickerSearchUser`;
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
                        const ensureUserUrl = `https://arabtec.sharepoint.com/_api/web/ensureUser`;
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