import { IDropdownOption } from "office-ui-fabric-react";
import { ICountryListItem } from "../models";

export interface IMyReactCrudCodeStates {
    listTitles: IDropdownOption[];
    countries: ICountryListItem[];
    status: string;
}
