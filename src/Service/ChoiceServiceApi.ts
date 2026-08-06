import { WebPartContext } from "@microsoft/sp-webpart-base";
import { ListName } from "../Enum/ListNames";

export default class GetChoiceApi{
    private context:WebPartContext;
    constructor(context:WebPartContext){
        this.context=context;
    }

    public async getChoiceValues(siteurl:string,fieldValue:string):Promise<any>{
        try{
            const response=await fetch(`${siteurl}/_api/web/lists/getbytitle('${ListName.SharePointListName}')/fields/?$filter=EntityPropertyName eq '${fieldValue}'`,

                {
                    method:'GET',
                    headers:{
                        'Accept':'application/json;odata=nometadata'
                    }
                }
            );
       if(!response.ok){
        throw new Error(`Error while fetching choice values ${response.text}-${response.status}`)
       } 
       const data=await response.json();
       const choice=data.value[0].Choices;
       return choice.map((item:any)=>({
        key:item,
        text:item
       }))   ; 
        }

        catch(err){
            console.error(err);
            return []
        }

    }

    public async getLookupValues():Promise<any>{
        try{
            const response=await fetch(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${ListName.Cities}')/items/?$select=Title,ID`,

                {
                    method:'GET',
                    headers:{
                       'Accept':'application/json;odata=nometadata'  
                    }
                }
            );
             if(!response.ok){
        throw new Error(`Error while fetching choice values ${response.text}-${response.status}`)
       } 
       const data=await response.json();
       return data.value.map((city:{Title:string,ID:string})=>({
        key:city.ID,
        text:city.Title
       }))
        }
        catch(err){
  console.error(err);
            return []
        }
    }
}