import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface ISpfxformProps {
 context:WebPartContext;
 siteurl:string;
 departmentOptions:any;
 genderOptions:any;
 cityOptions:any;
 skillsOptions:any;
}
