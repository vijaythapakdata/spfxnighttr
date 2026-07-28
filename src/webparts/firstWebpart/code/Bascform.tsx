import * as React from 'react'
import { TextField,PrimaryButton,Dropdown,Stack, ChoiceGroup } from '@fluentui/react';
const BasicForm=()=>{
    return(
        <>
        <Stack tokens={{childrenGap:8}} >
<TextField
label='Name'
placeholder='enter your name here..'
/>
<TextField
label='Email'
placeholder='enter your email here..'
iconProps={{iconName:"mail"}}
/>

<TextField
label='Salary'
prefix='$'
suffix='USD'
/>
<TextField
type='password'
label='Password'
canRevealPassword={true}
/>
<TextField
type='file'
label='Upload file'
/>

<Dropdown
options={[
    {key:"IT",text:"IT"},
    {key:"HR",text:"HR"}
]}
label='Department'
placeholder='--select--'
/>
<ChoiceGroup
options={[
    {key:"Male",text:"Male"},
    {key:"Female",text:"Female"}
]}
label='Gender'
/>
<TextField
label='Address'
placeholder='enter your address here..'
multiline
rows={5}
/>

        </Stack>
        <br/>
<PrimaryButton
text='Save'
onClick={()=>alert("Saved successfully")}
iconProps={{iconName:"save"}}
/>
        </>
    )
}
export default BasicForm;