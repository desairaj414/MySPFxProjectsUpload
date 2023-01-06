import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPHttpClient, SPHttpClientResponse, ISPHttpClientOptions } from "@microsoft/sp-http";
import { IDropdownOption } from "office-ui-fabric-react";

export class SPOperations{
    /**
     * GetAllList
context:WebpartContext     */
    public GetAllList(context:WebPartContext):Promise<IDropdownOption[]> {
        let restApiUrl: string = context.pageContext.web.absoluteUrl+"/_api/web/lists?select=Title";
        var listTitles: IDropdownOption[] = []
        return new Promise<IDropdownOption[]>((resolve,reject) => {
            context.spHttpClient
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

    /**
     * createListItem
context: WebPartContext, listTitle: string     */
    public createListItem(context: WebPartContext, listTitle: string):Promise<string> {
        let restApiUrl: string = context.pageContext.web.absoluteUrl +
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
        return new Promise<string>(async (resolve,reject) => {
            context.spHttpClient
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

    public deleteListItem(context: WebPartContext, listTitle: string):Promise<string> {
        let restApiUrl: string = context.pageContext.web.absoluteUrl +
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
        return new Promise<string>(async (resolve,reject) => {
            this.getLatestItemId(context,listTitle)
                .then((itemId: number)=>{
                    context.spHttpClient
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

    public updateListItem(context: WebPartContext, listTitle: string):Promise<string> {
        let restApiUrl: string = context.pageContext.web.absoluteUrl +
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
        return new Promise<string>(async (resolve,reject) => {
            context.spHttpClient
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

    public getLatestItemId(context: WebPartContext, listTitle: string): Promise<number>{
        let restApiUrl: string = context.pageContext.web.absoluteUrl +
            "/_api/web/lists/getByTitle('" +
            listTitle +
            "')/items?$orderby=Id desc&$top=1&select=Id";

        return new Promise<number>(async (resolve,reject) => {
            context.spHttpClient
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

}