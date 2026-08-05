import {Web} from "@pnp/sp/presets/all";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { ListName } from "../Enum/ListNames";
import { ISharePointListColumns } from "../CommonMethods/ISharePointFormState";
export class CommonServiceApi{
    private web;
    constructor(siteurl:string){
        this.web=Web(siteurl);
    }

    public async addItems(formData:ISharePointListColumns):Promise<any>{
        try{
const list=this.web.lists.getByTitle(ListName.SharePointListName);
const items=await list.items.add({
    Title:formData.Name,
    EmailAddress:formData.Email,
    Age:parseInt(formData.Age),
    Salary:parseFloat(formData.Salary),
    Permission:formData.Permission,
    Address:formData.FullAddress,
    AdminId:formData.AdminId,
    ManagerId:{results:formData.ManagerId}
});
return items
        }
        catch(err){
console.log(err);
throw err;
        }
    }
}