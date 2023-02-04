import {
  ButtonClickedCallback,
  ICountryListItem
} from '../../../models';

export interface IPnPHttpClientProps {
  spListItems: ICountryListItem[];
  onGetListItems?: ButtonClickedCallback;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  onAddListItem?: ButtonClickedCallback;
  onUpdateListItem?: ButtonClickedCallback;
  onDeleteListItem?: ButtonClickedCallback;
}
