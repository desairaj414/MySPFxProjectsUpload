import { IDropdownOption } from "office-ui-fabric-react";
import { SPFI } from "@pnp/sp";
import { getSP } from "../pnpjsConfig";

export class SPOperations{

    private sp: SPFI;

    constructor(){
        this.sp = getSP();
    }

    public getListTitles():Promise<IDropdownOption[]> {
        let listTitles:IDropdownOption[] = [];
        return new Promise<IDropdownOption[]>((resolve,reject) => {
            this.sp.web.lists.select('Title')().then(
                (results:any) => {
                    console.log(results);
                    results.map((result:any) => {
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
            )
        });
    }

    public createListItem(listTitle: string):Promise<string> {
        return new Promise<string>(async (resolve,reject) => {
            this.sp.web.lists.getByTitle(listTitle).items.add({Title:"New Pnp Item Created"}).then(
                (results: any)=>{
                    resolve("Item with Id "+results.data.ID+" created successfully");
                },
                (error: any)=>{
                    reject("Error occured"+error);
                }
            )
        });
    }

    public deleteListItem(listTitle: string):Promise<string> {
        return new Promise<string>(async (resolve,reject) => {
            this.getLatestItemId(listTitle).then((itemId: number)=>{
                this.sp.web.lists.getByTitle(listTitle).items.getById(itemId).delete().then(
                    (results: any)=>{
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
        return new Promise<string>(async (resolve,reject) => {
            this.getLatestItemId(listTitle).then((itemId: number)=>{
                this.sp.web.lists.getByTitle(listTitle).items.getById(itemId).update({Title:"Pnp Item Updated"}).then(
                    (results: any)=>{
                        resolve("Item with Id "+itemId+" updated sucessfully");
                    },
                    (error: any)=>{
                        reject("Error occured"+error);
                    }
                );
            });
        });
    }

    public getLatestItemId(listTitle: string): Promise<number>{
        return new Promise<number>(async (resolve,reject) => {
            this.sp.web.lists.getByTitle(listTitle).items.select("ID").orderBy("ID",false).top(1)().then(
                (results: any)=>{
                    resolve(results[0].Id);
                },
                (error: any)=>{
                    reject("Error occured"+error);
                }
            )
        });
    }

}