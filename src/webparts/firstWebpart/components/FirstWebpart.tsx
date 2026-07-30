import * as React from 'react';
// import styles from './FirstWebpart.module.scss';
import type { IFirstWebpartProps } from './IFirstWebpartProps';
import BasicForm from '../code/Bascform';
import CounterApp from '../code/CounterApp';
// import { escape } from '@microsoft/sp-lodash-subset';

const  FirstWebpart:React.FC<IFirstWebpartProps>=(props)=>{
  return(
    <>
    <p>Hello world...</p>
    <hr/>
    <BasicForm/>
    <CounterApp/>
    </>
  )
}
export default  FirstWebpart;