import { WebPartContext } from "@microsoft/sp-webpart-base";
import { spfi, SPFI, SPFx } from "@pnp/sp";
import { LogLevel, PnPLogging } from "@pnp/logging";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/batching";
import { SPHttpClient, SPHttpClientResponse, ISPHttpClientOptions } from '@microsoft/sp-http';
import { IDropdownOption } from "office-ui-fabric-react";
import { ICountryListItem } from "../models";

// Imports to use in other File
// import { getContext } from './SPServices.ts';
// import { Config } from '../../Services/SPServices';


// Config Functions


let _context: WebPartContext = null;
let _sp: SPFI = null;
let _getJson: any = null;

export const getContext = (context?: WebPartContext): WebPartContext => {
    if (context != null) {
        _context = context;
        _sp = spfi().using(SPFx(_context)).using(PnPLogging(LogLevel.Warning));
    }
    return _context;
};
export const getSP = (): SPFI => {
    return _sp;
};
export const getJson = async(url: string): Promise<any> => {
    let response = await _context.spHttpClient.get(url, SPHttpClient.configurations.v1);
    _getJson = await response.json();
    return _getJson;
}


// Main Operation Class Of File


export class SPOperations{


    // Config Class Methods


    private static context: WebPartContext;
    private static sp: SPFI;

    public static Init(context: WebPartContext) {
        SPOperations.context = context;
        this.sp  = spfi().using(SPFx(context)).using(PnPLogging(LogLevel.Warning));
    }

    public static getContext(): WebPartContext {
        return SPOperations.context;
    }

    public static getSP(): SPFI {
        return this.sp;
    }

    public static async getJson(url: string): Promise<any> {
        let response = await SPOperations.context.spHttpClient.get(url, SPHttpClient.configurations.v1);
        let json = await response.json();
        return json;
    }

    // Config to set List Title
    private static listTitle: string = "Countries"
    public static setListTitle(listTitle: string){
        SPOperations.listTitle = listTitle;
    }


    // YouTube Videos Methods


    public GetAllList():Promise<IDropdownOption[]> {
        let restApiUrl: string = SPOperations.context.pageContext.web.absoluteUrl+"/_api/web/lists?select=Title";
        let listTitles: IDropdownOption[] = []
        return new Promise<IDropdownOption[]>((resolve,reject) => {
            SPOperations.context.spHttpClient
            .get(restApiUrl, SPHttpClient.configurations.v1)
            .then((respone: SPHttpClientResponse) => {
                respone.json().then(
                    (results: any)=>{
                        results.value.map((result:any) => {
                            listTitles.push({
                                key: result.Title, 
                                text: result.Title
                            });
                        });
                        resolve(listTitles);
                    },
                    (error: any):void => {
                        reject("error occured"+error);
                    }
                );
            });
        });
    }

    public createListItem(listTitle: string):Promise<string> {
        let restApiUrl: string = SPOperations.context.pageContext.web.absoluteUrl +
            "/_api/web/lists/getByTitle('" +
            listTitle +
            "')/items";
        const body: string = JSON.stringify({Title:"New Item Created"});
        const options: ISPHttpClientOptions = {
            headers: {
                Accept: "application/json;odata=nometadata",
                "content-type": "application/json;odata=nometadata",
                "odata-version": ""
            },
            body: body
        };
        return new Promise<string>((resolve,reject) => {
            SPOperations.context.spHttpClient
                .post(restApiUrl, SPHttpClient.configurations.v1,options)
                .then((respone: SPHttpClientResponse)=>{
                    respone.json().then(
                        (results: any)=>{
                            resolve("Item with Id "+results.ID+" created successfully");
                        },
                        (error: any)=>{
                            reject("Error occured"+error);
                        }
                    );
                });
        });
    }

    public deleteListItem(listTitle: string):Promise<string> {
        let restApiUrl: string = SPOperations.context.pageContext.web.absoluteUrl +
            "/_api/web/lists/getByTitle('" +
            listTitle +
            "')/items";
        const options: ISPHttpClientOptions = {
            headers: {
                Accept: "application/json;odata=nometadata",
                "content-type": "application/json;odata=nometadata",
                "odata-version": "",
                "IF-MATCH": "*",
                "X-HTTP-METHOD": "DELETE",
            }
        };
        return new Promise<string>((resolve,reject) => {
            this.getLatestItemId(listTitle)
                .then((itemId: number)=>{
                    SPOperations.context.spHttpClient
                    .post(restApiUrl+"("+itemId+")", SPHttpClient.configurations.v1,options)
                    .then(
                        (respone: SPHttpClientResponse)=>{
                            resolve("Item with Id "+itemId+" deleted sucessfully");
                        },
                        (error: any)=>{
                            reject("Error occured"+error);
                        }
                    );
                });

        });
    }

    public updateListItem(listTitle: string):Promise<string> {
        let restApiUrl: string = SPOperations.context.pageContext.web.absoluteUrl +
            "/_api/web/lists/getByTitle('" +
            listTitle +
            "')/items('13')";
        const body: string = JSON.stringify({Title:"Updated Item"});
        const options: ISPHttpClientOptions = {
            headers: {
                Accept: "application/json;odata=nometadata",
                "content-type": "application/json;odata=nometadata",
                "odata-version": "",
                "IF-MATCH": "*",
                "X-HTTP-METHOD": "MERGE",
            },
            body: body
        };
        return new Promise<string>((resolve,reject) => {
            SPOperations.context.spHttpClient
                .post(restApiUrl, SPHttpClient.configurations.v1,options)
                .then(
                    (respone: SPHttpClientResponse)=>{
                        resolve("Item updated sucessfully");
                    },
                    (error: any)=>{
                        reject("Error occured"+error);
                    }
                );
        });
    }

    public getLatestItemId(listTitle: string): Promise<number>{
        let restApiUrl: string = SPOperations.context.pageContext.web.absoluteUrl +
            "/_api/web/lists/getByTitle('" +
            listTitle +
            "')/items?$orderby=Id desc&$top=1&$select=Id";

        return new Promise<number>((resolve,reject) => {
            SPOperations.context.spHttpClient
                .get(restApiUrl, SPHttpClient.configurations.v1)
                .then((respone: SPHttpClientResponse)=>{
                    respone.json().then(
                        (results: any)=>{
                            resolve(results.value[0].Id);
                        },
                        (error: any)=>{
                            reject("Error occured"+error);
                        }
                    ); 
                });
        });
    }


    // Microsoft Learn Methods


    public _onGetListItems = async (): Promise<ICountryListItem[]> => {
        const response: ICountryListItem[] = await this._getListItems();
        return response;
    }
    
    public async _getListItems(): Promise<ICountryListItem[]> {
        const response = await SPOperations.context.spHttpClient.get(
            SPOperations.context.pageContext.web.absoluteUrl + `/_api/web/lists/getbytitle('`+SPOperations.listTitle+`')/items?$select=Id,Title`,
            SPHttpClient.configurations.v1);
        // `/_api/web/lists/getbytitle('Countries')/items?$select=Id,Title`

        if (!response.ok) {
            const responseText = await response.text();
            throw new Error(responseText);
        }

        const responseJson = await response.json();
        console.log(responseJson);
      
        return responseJson.value as ICountryListItem[];
    }
    
    public async _getItemEntityType(): Promise<string> {
        const endpoint: string = SPOperations.context.pageContext.web.absoluteUrl + 
            `/_api/web/lists/getbytitle('`+SPOperations.listTitle+`')/items?$select=Id,Title`;
        
        const response = await SPOperations.context.spHttpClient.get(
            endpoint,
            SPHttpClient.configurations.v1);
        
        if (!response.ok) {
            const responseText = await response.text();
            throw new Error(responseText);
        }
        
        const responseJson = await response.json();
        
        return responseJson.ListItemEntityTypeFullName;
    }

    public _onAddListItem = async (): Promise<ICountryListItem[]> => {
        const addResponse: SPHttpClientResponse = await this._addListItem();
        
        if (!addResponse.ok) {
            const responseText = await addResponse.text();
            throw new Error(responseText);
        }
        
        const getResponse: ICountryListItem[] = await this._getListItems();
        return getResponse;
    }

    public async _addListItem(): Promise<SPHttpClientResponse> {
        const itemEntityType = await this._getItemEntityType();
        
        const endpoint = SPOperations.context.pageContext.web.absoluteUrl + 
            `/_api/web/lists/getbytitle('`+SPOperations.listTitle+`')/items`;

        /* eslint-disable @typescript-eslint/no-explicit-any */
        const request: any = {};
        request.body = JSON.stringify({
            Title: new Date().toUTCString(),
            '@odata.type': itemEntityType
        });
        /* eslint-enable @typescript-eslint/no-explicit-any */
        
        return SPOperations.context.spHttpClient.post(
            endpoint,
            SPHttpClient.configurations.v1,
            request);
    }
    
    public _onUpdateListItem = async (): Promise<ICountryListItem[]> => {
        const updateResponse: SPHttpClientResponse = await this._updateListItem();
        
        if (!updateResponse.ok) {
            const responseText = await updateResponse.text();
            throw new Error(responseText);
        }
        
        const getResponse: ICountryListItem[] = await this._getListItems();
        return getResponse;
    }

    public async _updateListItem(): Promise<SPHttpClientResponse> {
        const getEndpoint: string = SPOperations.context.pageContext.web.absoluteUrl + 
            `/_api/web/lists/getbytitle('`+SPOperations.listTitle+`')/items?` +
            `$select=Id,Title&$orderby=ID desc&$top=1`;
        // $filter=Title eq 'United States'
        
        const getResponse = await SPOperations.context.spHttpClient.get(
            getEndpoint,
            SPHttpClient.configurations.v1);
        
        if (!getResponse.ok) {
            const responseText = await getResponse.text();
            throw new Error(responseText);
        }
        
        const responseJson = await getResponse.json();
        const listItem: ICountryListItem = responseJson.value[0];
        
        listItem.Title = 'USA';

        const postEndpoint: string = SPOperations.context.pageContext.web.absoluteUrl + 
            `/_api/web/lists/getbytitle('`+SPOperations.listTitle+`')/items(${listItem.Id})`;

        /* eslint-disable @typescript-eslint/no-explicit-any */
        const request: any = {};
        request.headers = {
            'X-HTTP-Method': 'MERGE',
            'IF-MATCH': (listItem as any)['@odata.etag']
        };
        /* eslint-enable @typescript-eslint/no-explicit-any */
        request.body = JSON.stringify(listItem);
        
        return SPOperations.context.spHttpClient.post(
            postEndpoint,
            SPHttpClient.configurations.v1,
            request);
    }
    
    public _onDeleteListItem = async (): Promise<ICountryListItem[]> => {
        const deleteResponse: SPHttpClientResponse = await this._deleteListItem();
        
        if (!deleteResponse.ok) {
            const responseText = await deleteResponse.text();
            throw new Error(responseText);
        }
        
        const getResponse: ICountryListItem[] = await this._getListItems();
        return getResponse;
    }

    public async _deleteListItem(): Promise<SPHttpClientResponse> {
        const getEndpoint = SPOperations.context.pageContext.web.absoluteUrl + 
            `/_api/web/lists/getbytitle('`+SPOperations.listTitle+`')/items?` +
            `$select=Id,Title&$orderby=ID desc&$top=1`;
        
        const getResponse = await SPOperations.context.spHttpClient.get(
            getEndpoint,
            SPHttpClient.configurations.v1);
        
        if (!getResponse.ok) {
            const responseText = await getResponse.text();
            throw new Error(responseText);
        }
        
        const responseJson = await getResponse.json();
        const listItem: ICountryListItem = responseJson.value[0];
        
        const postEndpoint = SPOperations.context.pageContext.web.absoluteUrl + 
            `/_api/web/lists/getbytitle('`+SPOperations.listTitle+`')/items(${listItem.Id})`;

        /* eslint-disable @typescript-eslint/no-explicit-any */
        const request: any = {};
        request.headers = {
            'X-HTTP-Method': 'DELETE',
            'IF-MATCH': '*'
        };
        /* eslint-enable @typescript-eslint/no-explicit-any */
        request.body = JSON.stringify(listItem);
        
        return SPOperations.context.spHttpClient.post(
            postEndpoint,
            SPHttpClient.configurations.v1,
            request);
    }

    
    // My Updated Microsoft Learn Methods

    
    public async _getItemEntityTypeShort(): Promise<string> {
        const endpoint: string = SPOperations.context.pageContext.web.absoluteUrl + 
            `/_api/web/lists/getbytitle('`+SPOperations.listTitle+`')/items?$select=Id,Title`;
        
        const response = await SPOperations.context.spHttpClient.get(
            endpoint,
            SPHttpClient.configurations.v1);
        
        if (!response.ok) {
            const responseText = await response.text();
            throw new Error(responseText);
        }
        
        const responseJson = await response.json();
        
        return responseJson.ListItemEntityTypeFullName;
    }

    public async _getLatestItemShort(): Promise<ICountryListItem> {
        const getEndpoint: string = SPOperations.context.pageContext.web.absoluteUrl + 
            `/_api/web/lists/getbytitle('`+SPOperations.listTitle+`')/items?` +
            `$select=Id,Title&$orderby=ID desc&$top=1`;
        // $filter=Title eq 'United States'
        
        const getResponse = await SPOperations.context.spHttpClient.get(
            getEndpoint,
            SPHttpClient.configurations.v1);
        
        if (!getResponse.ok) {
            const responseText = await getResponse.text();
            throw new Error(responseText);
        }
        
        const responseJson = await getResponse.json();
        const listItem: ICountryListItem = responseJson.value[0];

        return listItem;
    }

    public async _getListItemsShort(): Promise<ICountryListItem[]> {
        const response = await SPOperations.context.spHttpClient.get(
            SPOperations.context.pageContext.web.absoluteUrl + `/_api/web/lists/getbytitle('`+SPOperations.listTitle+`')/items?$select=Id,Title`,
            SPHttpClient.configurations.v1);
        // `/_api/web/lists/getbytitle('Countries')/items?$select=Id,Title`

        if (!response.ok) {
            const responseText = await response.text();
            throw new Error(responseText);
        }

        const responseJson = await response.json();
        console.log(responseJson);
      
        return responseJson.value as ICountryListItem[];
    }

    public async _addListItemShort(): Promise<ICountryListItem[]> {
        const itemEntityType = await this._getItemEntityTypeShort();
        
        const endpoint = SPOperations.context.pageContext.web.absoluteUrl + 
            `/_api/web/lists/getbytitle('`+SPOperations.listTitle+`')/items`;

        /* eslint-disable @typescript-eslint/no-explicit-any */
        const request: any = {};
        request.body = JSON.stringify({
            Title: new Date().toUTCString(),
            '@odata.type': itemEntityType
        });
        /* eslint-enable @typescript-eslint/no-explicit-any */
        
        const addResponse: SPHttpClientResponse = await SPOperations.context.spHttpClient.post(
            endpoint,
            SPHttpClient.configurations.v1,
            request);

        if (!addResponse.ok) {
            const responseText = await addResponse.text();
            throw new Error(responseText);
        }
        
        return await this._getListItems();
    }

    public async _updateListItemShort(): Promise<ICountryListItem[]> {
        const listItem: ICountryListItem = await this._getLatestItemShort();
        listItem.Title = 'USA';

        const postEndpoint: string = SPOperations.context.pageContext.web.absoluteUrl + 
            `/_api/web/lists/getbytitle('`+SPOperations.listTitle+`')/items(${listItem.Id})`;

        /* eslint-disable @typescript-eslint/no-explicit-any */
        const request: any = {};
        request.headers = {
            'X-HTTP-Method': 'MERGE',
            'IF-MATCH': (listItem as any)['@odata.etag']
        };
        /* eslint-enable @typescript-eslint/no-explicit-any */
        request.body = JSON.stringify(listItem);
        
        const updateResponse: SPHttpClientResponse = await SPOperations.context.spHttpClient.post(
            postEndpoint,
            SPHttpClient.configurations.v1,
            request);

        if (!updateResponse.ok) {
            const responseText = await updateResponse.text();
            throw new Error(responseText);
        }
        
        return await this._getListItems();
    }

    public async _deleteListItemShort(): Promise<ICountryListItem[]> {
        const listItem: ICountryListItem = await this._getLatestItemShort();
        
        const postEndpoint = SPOperations.context.pageContext.web.absoluteUrl + 
            `/_api/web/lists/getbytitle('`+SPOperations.listTitle+`')/items(${listItem.Id})`;

        /* eslint-disable @typescript-eslint/no-explicit-any */
        const request: any = {};
        request.headers = {
            'X-HTTP-Method': 'DELETE',
            'IF-MATCH': '*'
        };
        /* eslint-enable @typescript-eslint/no-explicit-any */
        request.body = JSON.stringify(listItem);
        
        const deleteResponse: SPHttpClientResponse = await SPOperations.context.spHttpClient.post(
            postEndpoint,
            SPHttpClient.configurations.v1,
            request);

        if (!deleteResponse.ok) {
            const responseText = await deleteResponse.text();
            throw new Error(responseText);
        }

        return await this._getListItems();
    }


    //  My Template for General CRUD Operations


    public async _getItemEntityTypeTemplate(listTitle: string): Promise<string> {
        const endpoint: string = SPOperations.context.pageContext.web.absoluteUrl + 
            `/_api/web/lists/getbytitle('`+listTitle+`')/items?$select=Id,Title`;
        
        const response = await SPOperations.context.spHttpClient.get(
            endpoint,
            SPHttpClient.configurations.v1);
        
        if (!response.ok) {
            const responseText = await response.text();
            throw new Error(responseText);
        }
        
        const responseJson = await response.json();
        
        return responseJson.ListItemEntityTypeFullName;
    }

    public async _getLatestItemTemplate(listTitle: string): Promise<ICountryListItem> {
        const getEndpoint: string = SPOperations.context.pageContext.web.absoluteUrl + 
            `/_api/web/lists/getbytitle('`+listTitle+`')/items?` +
            `$select=Id,Title&$orderby=ID desc&$top=1`;
        // $filter=Title eq 'United States'
        
        const getResponse = await SPOperations.context.spHttpClient.get(
            getEndpoint,
            SPHttpClient.configurations.v1);
        
        if (!getResponse.ok) {
            const responseText = await getResponse.text();
            throw new Error(responseText);
        }
        
        const responseJson = await getResponse.json();
        const listItem: ICountryListItem = responseJson.value[0];

        return listItem;
    }

    public async _getListItemsTemplate(listTitle: string): Promise<any> {
        const response = await SPOperations.context.spHttpClient.get(
            SPOperations.context.pageContext.web.absoluteUrl + `/_api/web/lists/getbytitle('`+listTitle+`')/items?$select=Id,Title`,
            SPHttpClient.configurations.v1);
        // `/_api/web/lists/getbytitle('Countries')/items?$select=Id,Title`

        if (!response.ok) {
            const responseText = await response.text();
            throw new Error(responseText);
        }

        const responseJson = await response.json();
        console.log(responseJson);
      
        //return responseJson.value as ICountryListItem[];
        return responseJson;
    }

    public async _addListItemTemplate(listTitle: string, Title:string): Promise<any> {
        const itemEntityType = await this._getItemEntityTypeTemplate(listTitle);
        
        const endpoint = SPOperations.context.pageContext.web.absoluteUrl + 
            `/_api/web/lists/getbytitle('`+listTitle+`')/items`;

        /* eslint-disable @typescript-eslint/no-explicit-any */
        const request: any = {};
        request.body = JSON.stringify({
            Title: Title,
            '@odata.type': itemEntityType
        });
        // request.body = body;
        /* eslint-enable @typescript-eslint/no-explicit-any */
        
        const addResponse: SPHttpClientResponse = await SPOperations.context.spHttpClient.post(
            endpoint,
            SPHttpClient.configurations.v1,
            request);

        if (!addResponse.ok) {
            const responseText = await addResponse.text();
            throw new Error(responseText);
        }
        
        //return await this._getListItems();
        return addResponse;
    }

    public async _updateListItemTemplate(listTitle: string, Title:string): Promise<any> {
        const listItem: ICountryListItem = await this._getLatestItemTemplate(listTitle);
        listItem.Title = Title;

        const postEndpoint: string = SPOperations.context.pageContext.web.absoluteUrl + 
            `/_api/web/lists/getbytitle('`+listTitle+`')/items(${listItem.Id})`;

        /* eslint-disable @typescript-eslint/no-explicit-any */
        const request: any = {};
        request.headers = {
            'X-HTTP-Method': 'MERGE',
            'IF-MATCH': (listItem as any)['@odata.etag']
        };
        /* eslint-enable @typescript-eslint/no-explicit-any */
        request.body = JSON.stringify(listItem);
        
        const updateResponse: SPHttpClientResponse = await SPOperations.context.spHttpClient.post(
            postEndpoint,
            SPHttpClient.configurations.v1,
            request);

        if (!updateResponse.ok) {
            const responseText = await updateResponse.text();
            throw new Error(responseText);
        }
        
        return await this._getListItems();
    }

    public async _deleteListItemTemplate(listTitle: string): Promise<ICountryListItem[]> {
        const listItem: ICountryListItem = await this._getLatestItemTemplate(listTitle);
        
        const postEndpoint = SPOperations.context.pageContext.web.absoluteUrl + 
            `/_api/web/lists/getbytitle('`+listTitle+`')/items(${listItem.Id})`;

        /* eslint-disable @typescript-eslint/no-explicit-any */
        const request: any = {};
        request.headers = {
            'X-HTTP-Method': 'DELETE',
            'IF-MATCH': '*'
        };
        /* eslint-enable @typescript-eslint/no-explicit-any */
        request.body = JSON.stringify(listItem);
        
        const deleteResponse: SPHttpClientResponse = await SPOperations.context.spHttpClient.post(
            postEndpoint,
            SPHttpClient.configurations.v1,
            request);

        if (!deleteResponse.ok) {
            const responseText = await deleteResponse.text();
            throw new Error(responseText);
        }

        return await this._getListItems();
    }

}