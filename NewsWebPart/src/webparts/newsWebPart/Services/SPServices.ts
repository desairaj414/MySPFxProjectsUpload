import { WebPartContext } from "@microsoft/sp-webpart-base";
import { spfi, SPFI, SPFx } from "@pnp/sp";
import { LogLevel, PnPLogging } from "@pnp/logging";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/batching";
import { SPHttpClient } from '@microsoft/sp-http';
//import { SPHttpClientResponse, ISPHttpClientOptions } from '@microsoft/sp-http';
import { INewsListItem } from "../models";


// Main Operation Class Of File


export class SPOperations{


    // Config Class Methods


    private static context: WebPartContext;
    private static sp: SPFI;
    context1: WebPartContext;
    sp1: SPFI;
    
    constructor(){
        this.context1 = SPOperations.context;
        this.sp1  = spfi().using(SPFx(this.context1)).using(PnPLogging(LogLevel.Warning));
    }

    public static Init(context: WebPartContext): void {
        SPOperations.context = context;
        this.sp  = spfi().using(SPFx(context)).using(PnPLogging(LogLevel.Warning));
    }

    public static getContext(): WebPartContext {
        return SPOperations.context;
    }

    public static getSP(): SPFI {
        return this.sp;
    }

    // public static async getJson(url: string): Promise<any> {
    //     const response = await SPOperations.context.spHttpClient.get(url, SPHttpClient.configurations.v1);
    //     const json = await response.json();
    //     return json;
    // }

    public async getNewsList():Promise<INewsListItem[]>{
        const today = new Date();
        const response = await SPOperations.context.spHttpClient.get(
            SPOperations.context.pageContext.web.absoluteUrl + 
            `/_api/web/lists/getbytitle('NewsWebPartList')/items?` +
            `$select=NewsTitle,PublishDate,NewsContent,NewsLink&` +
            `$filter=PublishDate le datetime'` + today.toISOString() + `'&` +
            `$orderby=PublishDate asc&`+
            `$top=4`,
            SPHttpClient.configurations.v1);

        if (!response.ok) {
            const responseText = await response.text();
            throw new Error(responseText);
        }

        const responseJson = await response.json();
        console.log(responseJson);
      
        return responseJson.value as INewsListItem[];
    }

}