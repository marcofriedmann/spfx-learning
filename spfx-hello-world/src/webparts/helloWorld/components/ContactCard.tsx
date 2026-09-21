import * as React from 'react';
import styles from './ContactCard.module.scss';
import type { IContactCardProps } from './IContactCardProps';
import { escape } from '@microsoft/sp-lodash-subset';
import welcomeDark from '../assets/welcome-dark.png';
import welcomeLight from '../assets/welcome-light.png';

const HelloWorld: React.FC<IContactCardProps> = (props) => {
  const [clickCount, setClickCount] = React.useState<number>(0);

  const _handleClick = (): void => {
    setClickCount((previousCount) => previousCount + 1);
  };

  const {
    title,
    name,
    role,
    email,
  } = props;

  return (
    <section className={`${styles.helloWorld}`}>
      
      <div>
        
        <p>Name: {escape(name)}. </p>
        <p>Titel: {escape(title)}. </p>
        <p>Rolle: {escape(role)}. </p>
        <p>E-Mail: {escape(email)}. </p>
        
      </div>
    </section>
  );
};

export default HelloWorld;
