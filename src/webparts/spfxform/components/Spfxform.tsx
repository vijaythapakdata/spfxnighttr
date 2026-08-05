import * as React from 'react';
// import styles from './Spfxform.module.scss';
import type { ISpfxformProps } from './ISpfxformProps';
// import { escape } from '@microsoft/sp-lodash-subset';
import { CommonServiceApi } from '../../../Service/SharePointFormService';
import { ISharePointListColumns } from '../../../CommonMethods/ISharePointFormState';
import {sp} from "@pnp/sp/presets/all";
import { PrimaryButton, TextField, Toggle } from '@fluentui/react';
import {PeoplePicker,PrincipalType} from "@pnp/spfx-controls-react/lib/PeoplePicker"
import { handleMultiSelectedPeoplePicker, handleSingleSelectedPeoplePicker } from '../../../CommonMethods/PeoplePickerHandler';
const Spfxform:React.FC<ISpfxformProps>=(props)=>{
  const [formdata,setFormData]=React.useState<ISharePointListColumns>({
    Name:"",
    Email:"",
    Age:"",
    Salary:"",
    Permission:false,
    FullAddress:"",
    Admin:"",
    AdminId:"",
    Manager:[],
    ManagerId:[]
  });

React.useEffect(()=>{
  sp.setup({
    spfxContext:props.context as any
  })
},[]);


const createForm=async()=>{
  try{
    const _service=new CommonServiceApi(props.siteurl);
    const result=await _service.addItems(formdata);
    alert(`Item save successfully with id ${result.data.Id}`);
    setFormData({
      Name:"",
    Email:"",
    Age:"",
    Salary:"",
    Permission:false,
    FullAddress:"",
    Admin:"",
    AdminId:"",
    Manager:[],
    ManagerId:[]
    })

  }
  catch(err){
    console.log(err);
  }
}

const handleSubmit=React.useCallback((field:keyof ISharePointListColumns,value:string|number|boolean)=>{
  setFormData(prev=>({...prev,[field]:value}));
},[])
  return(
    <>
    
    <TextField
    label='Name'
    value={formdata.Name}
    onChange={(_,val)=>handleSubmit("Name",val||"")}
    />
      <TextField
    label='Email'
    value={formdata.Email}
    onChange={(_,val)=>handleSubmit("Email",val||"")}
    />
      <TextField
    label='Age'
    value={formdata.Age}
    onChange={(_,val)=>handleSubmit("Age",val||"")}
    />
      <TextField
    label='Salary'
    value={formdata.Salary}
    onChange={(_,val)=>handleSubmit("Salary",val||"")}
    prefix='$'
    suffix='USD'
    />
    <Toggle
    label="Permission"
    checked={formdata.Permission}
    onChange={(_,ch)=>handleSubmit("Permission",!!ch)}
    />
    {/* Single Selected */}
    <PeoplePicker
    context={props.context as any}
    titleText='Admin'
    personSelectionLimit={1}
    showtooltip={true}
    onChange={(items)=>handleSingleSelectedPeoplePicker(items,setFormData)}
    principalTypes={[PrincipalType.User]}
    ensureUser={true}
    defaultSelectedUsers={[formdata.Admin?formdata.Admin:""]}
    resolveDelay={1000}
    webAbsoluteUrl={props.siteurl}
    />
    {/* multiselected people picker */}
     <PeoplePicker
    context={props.context as any}
    titleText='Manager'
    personSelectionLimit={2}
    showtooltip={true}
    onChange={(items)=>handleMultiSelectedPeoplePicker(items,setFormData)}
    principalTypes={[PrincipalType.User]}
    ensureUser={true}
    defaultSelectedUsers={formdata.Manager}
    resolveDelay={1000}
    webAbsoluteUrl={props.siteurl}
    />
      <TextField
    label='Full Address'
    value={formdata.FullAddress}
    onChange={(_,val)=>handleSubmit("FullAddress",val||"")}
    multiline
    rows={3}
    />
    <br/>
    <PrimaryButton
    text='Save'
    onClick={createForm}
    />
    </>
  )
}
export default Spfxform;
