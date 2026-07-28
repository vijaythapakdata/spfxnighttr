import * as React from 'react';
// import styles from './FirstWebpart.module.scss';
import type { IFirstWebpartProps } from './IFirstWebpartProps';
import BasicForm from '../code/Bascform';
// import { escape } from '@microsoft/sp-lodash-subset';

const  FirstWebpart:React.FC<IFirstWebpartProps>=(props)=>{
  return(
    <>
    <p>Hello world...</p>
    <hr/>
    <BasicForm/>
    </>
  )
}
export default  FirstWebpart;