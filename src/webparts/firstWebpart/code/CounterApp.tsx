import { PrimaryButton, TextField } from '@fluentui/react';
import * as React from 'react';
import { useState } from 'react';

const CounterApp=()=>{
    const [count,setCount]=useState<number>(0);
    const[name,setName]=useState<string>("");

    return(
        <>
        <p>Count:{count}</p>
        <PrimaryButton
        text='Count'
        onClick={()=>setCount(count+1)}
        />
        <p>Name:{name}</p>
        <TextField
        label='Name'
        onChange={(_,name)=>setName(name||"")}
        />
        </>
    )
}
export default CounterApp;