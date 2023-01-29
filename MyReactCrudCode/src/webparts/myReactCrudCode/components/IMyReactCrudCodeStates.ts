import { IDropdownOption } from "office-ui-fabric-react";
import { ICountryListItem } from "../models";

export interface IMyReactCrudCodeStates {
    listTitles: IDropdownOption[];
    countries: ICountryListItem[];
    status: string;
    // REACT Form CRUD States
    Items: any;
    ID: any;
    EmployeeName: any;
    EmployeeNameId: any;
    HireDate: any;
    JobDescription: string;
    HTML: any;
    // Microsoft Graph API State
    MSGraphHTML: any;
}
