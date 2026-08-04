import * as React from 'react';
// import styles from './Spfxform.module.scss';
import type { ISpfxformProps } from './ISpfxformProps';
// import { escape } from '@microsoft/sp-lodash-subset';
import { CommonServiceApi } from '../../../Service/SharePointFormService';
import { ISharePointListColumns } from '../../../CommonMethods/ISharePointFormState';
import {sp} from "@pnp/sp/presets/all";
import { PrimaryButton, TextField, Toggle } from '@fluentui/react';
const Spfxform:React.FC<ISpfxformProps>=(props)=>{
  const [formdata,setFormData]=React.useState<ISharePointListColumns>({
    Name:"",
    Email:"",
    Age:"",
    Salary:"",
    Permission:false,
    FullAddress:""
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
    FullAddress:""
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
