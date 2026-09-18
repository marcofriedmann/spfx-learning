import * as React from 'react';
import styles from './HelloWorld.module.scss';
import type { IHelloWorldProps } from './IHelloWorldProps';
import { escape } from '@microsoft/sp-lodash-subset';
import welcomeDark from '../assets/welcome-dark.png';
import welcomeLight from '../assets/welcome-light.png';

const HelloWorld: React.FC<IHelloWorldProps> = (props) => {
  const [clickCount, setClickCount] = React.useState<number>(0);

  const _handleClick = (): void => {
    setClickCount((previousCount) => previousCount + 1);
  };

  const {
    description,
    isDarkTheme,
    environmentMessage,
    userDisplayName,
    siteTitle,
    title,
  } = props;

  return (
    <section className={`${styles.helloWorld}`}>
      <div className={styles.welcome}>
        <img
          alt=""
          src={isDarkTheme ? welcomeDark : welcomeLight}
          className={styles.welcomeImage}
        />
        <h2>Well done Marco, {escape(userDisplayName)}!</h2>
        <div>{environmentMessage}</div>
        <div>
          Web part property value: <strong>{escape(description)}</strong>
        </div>
      </div>
      <p>Button clicked {clickCount} times. </p>
      <button onClick={_handleClick}>Click me</button>
      <div>
        <h3>Mein erstes SPFx WebPart 🚀</h3>
        <p>Diese Site heißt {escape(siteTitle)}. </p>
        <p>Der Titel des Webparts ist {escape(title)}. </p>
        <p>
          The SharePoint Framework (SPFx) is a extensibility model for Microsoft
          Viva, Microsoft Teams and SharePoint. It&#39;s the easiest way to
          extend Microsoft 365 with automatic Single Sign On, automatic hosting
          and industry standard tooling.
        </p>
        <h4>Learn more about SPFx development:</h4>
        <ul className={styles.links}>
          <li>
            <a href="https://aka.ms/spfx" target="_blank" rel="noreferrer">
              SharePoint Framework Overview
            </a>
          </li>
          <li>
            <a
              href="https://aka.ms/spfx-yeoman-graph"
              target="_blank"
              rel="noreferrer"
            >
              Use Microsoft Graph in your solution
            </a>
          </li>
          <li>
            <a
              href="https://aka.ms/spfx-yeoman-teams"
              target="_blank"
              rel="noreferrer"
            >
              Build for Microsoft Teams using SharePoint Framework
            </a>
          </li>
          <li>
            <a
              href="https://aka.ms/spfx-yeoman-viva"
              target="_blank"
              rel="noreferrer"
            >
              Build for Microsoft Viva Connections using SharePoint Framework
            </a>
          </li>
          <li>
            <a
              href="https://aka.ms/spfx-yeoman-store"
              target="_blank"
              rel="noreferrer"
            >
              Publish SharePoint Framework applications to the marketplace
            </a>
          </li>
          <li>
            <a
              href="https://aka.ms/spfx-yeoman-api"
              target="_blank"
              rel="noreferrer"
            >
              SharePoint Framework API reference
            </a>
          </li>
          <li>
            <a href="https://aka.ms/m365pnp" target="_blank" rel="noreferrer">
              Microsoft 365 Developer Community
            </a>
          </li>
        </ul>
      </div>
    </section>
  );
};

export default HelloWorld;
