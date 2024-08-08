import * as React from 'react';
import { useState, useEffect } from 'react';
import styles from './ReactHelloWorld.module.scss';
import type { IReactHelloWorldProps } from './IReactHelloWorldProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

const ReactHelloWorld: React.FunctionComponent<IReactHelloWorldProps> = (props:IReactHelloWorldProps) => {

  const {
    description,
    environmentMessage,
    userDisplayName,
    siteURL,
    spHttpClient
  } = props;
  
  const [count, setCount] = useState(0);
  const [evenOdd, setEvenOdd] = useState<string>('');
  const [siteLists, setSiteLists] = useState<string[]>([]);

  const incrementCount = () => {
    console.log("Increment button clicked");
    setCount(count + 1);
  };

/*  
  const [yourData, setYourData] = useState<any[]>([]);

  const loadListData = async (): Promise<void> => {
    try {
     const apiUrl: string = `${props.context.pageContext.web.absoluteUrl}/_api/web/lists/getByTitle('Important Links')/items`;
     const response = await props.context.spHttpClient.get(apiUrl, SPHttpClient.configurations.v1);
   
     if (response.ok) {
       const data = await response.json();
       setYourData(data);
     } else {
       console.log('Failed to fetch data from SharePoint. Error:'
       ,response.statusText);
       }
     } catch (error) {
        console.log('Error loading data from SharePoint:', error);
       }
   };
*/

  useEffect(() => {
    console.log("componentDidMount called");

    /* eslint-disable @typescript-eslint/no-floating-promises */
    (async () => {
      // line wrapping added for readability
      const endpoint: string = `${siteURL}/_api/web/lists?$select=Title&$filter=Hidden eq false&$orderby=Title&$top=10`;
      const rawResponse: SPHttpClientResponse = await spHttpClient.get(endpoint, SPHttpClient.configurations.v1);
      setSiteLists(
        (await rawResponse.json()).value.map((list: { Title: string }) => {
          return list.Title;
        })
      );
    })();
  }, []);

  useEffect(() => {
    setEvenOdd((count % 2 === 0) ? 'even' : 'odd');
  }, [count]);

  //useEffect(() => {
    //loadListData();
  //}, []);

  useEffect(() => {
    console.log("componentDidUpdate called");
  });

  useEffect(() => {
    return () => {
        console.log("componentWillUnmount called");
    }
  },[count]);
 
  return (
    <section className={`${styles.reactHelloWorld}`}>
      <div className={styles.welcome}>
        <h2>Well done, {escape(userDisplayName)}!</h2>
        <div>{environmentMessage}</div>
        <div>Web part property value: <strong>{escape(description)}</strong></div>
      </div>
      <div>Counter: <strong>{count}</strong> is <strong>{evenOdd}</strong></div>
      <button onClick={incrementCount}>Increment</button>
      <ul>
         {
           siteLists.map((list: string) => (
             <li key={list}>{list}</li>
           ))
         }
      </ul>
      <div>
        <h3>Welcome to SharePoint Framework!</h3>
        <p>
          The SharePoint Framework (SPFx) is a extensibility model for Microsoft Viva, Microsoft Teams and SharePoint. It&#39;s the easiest way to extend Microsoft 365 with automatic Single Sign On, automatic hosting and industry standard tooling.
        </p>
        <h4>Learn more about SPFx development:</h4>
        <ul className={styles.links}>
          <li><a href="https://aka.ms/spfx" target="_blank" rel="noreferrer">SharePoint Framework Overview</a></li>
          <li><a href="https://aka.ms/spfx-yeoman-graph" target="_blank" rel="noreferrer">Use Microsoft Graph in your solution</a></li>
          <li><a href="https://aka.ms/spfx-yeoman-teams" target="_blank" rel="noreferrer">Build for Microsoft Teams using SharePoint Framework</a></li>
          <li><a href="https://aka.ms/spfx-yeoman-viva" target="_blank" rel="noreferrer">Build for Microsoft Viva Connections using SharePoint Framework</a></li>
          <li><a href="https://aka.ms/spfx-yeoman-store" target="_blank" rel="noreferrer">Publish SharePoint Framework applications to the marketplace</a></li>
          <li><a href="https://aka.ms/spfx-yeoman-api" target="_blank" rel="noreferrer">SharePoint Framework API reference</a></li>
          <li><a href="https://aka.ms/m365pnp" target="_blank" rel="noreferrer">Microsoft 365 Developer Community</a></li>
        </ul>
      </div>
    </section>
  );
}

export default ReactHelloWorld;