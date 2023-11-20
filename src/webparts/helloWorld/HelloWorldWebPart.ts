import { BaseClientSideWebPart, } from '@microsoft/sp-webpart-base';
//import styles from './HelloWorldWebPart.module.scss';
import {
    type IPropertyPaneConfiguration,
   
    PropertyPaneTextField,
    PropertyPaneDropdown,
    //PropertyPaneDropdownOptionType,
    IPropertyPaneDropdownOption
    
  } from '@microsoft/sp-property-pane';

import 'bootstrap/dist/css/bootstrap.css';

import 'bootstrap/dist/js/bootstrap.bundle';

import { 
  SPHttpClient,
  SPHttpClientResponse
 } from '@microsoft/sp-http';

import { Version } from '@microsoft/sp-core-library';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { escape } from '@microsoft/sp-lodash-subset';

export interface ISPLists {
  value: ISPList[];
}

export interface ISPList {
  Title: string;
  Id: string;
  Name: string;
  Surname: string;
  Image: string;
}

export interface IHelloWorldWebPartProps {

  description: string;
  selectedList: string;


}

export default class HelloWorldWebPart extends BaseClientSideWebPart<IHelloWorldWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  private lists: IPropertyPaneDropdownOption[] = [];

 // private lists:PropertyPaneDropdownOptionType = [];

 private _selectedList: string = '';


  public render(): void {
    this.domElement.innerHTML = `
      <section class="helloWorld ${!!this.context.sdks.microsoftTeams ? 'teams' : ''}">
        <div class="welcome">
          <img alt="" src="${this._isDarkTheme ? require('./assets/welcome-dark.png') : require('./assets/welcome-light.png')}" class="welcomeImage" />
          <h2>Well done, ${escape(this.context.pageContext.user.displayName)}!</h2>
          <div>${this._environmentMessage}</div>
        </div>
        <div>
          <h3>Welcome to SharePoint Framework!</h3>
          <div>Web part description: <strong>${escape(this.properties.description)}</strong></div>
          <div>Loading from: <strong>${escape(this.context.pageContext.web.title)}</strong></div>
          <div>something now happening again:-</div>
          <div>this is running now.</div>

          <div>this is running again:-</div>


          <div>
            <h3>Select a list:</h3>
            <div>
              <select id="listSelector" onchange="fetchListItems(this.value)">
                <option value="">-- Select a list --</option>
              </select>
            </div>
            <div class="" id="itemList"></div>

            

            <div id="carouselExample" class="carousel slide">
 
  <div class="carousel-inner">
    
    
    
    
  </div>
  <button class="carousel-control-prev mt-10" type="button" data-bs-target="#carouselExample" data-bs-slide="prev">
    <span class="carousel-control-prev-icon" aria-hidden="true"></span>
    <span class="visually-hidden">Previous</span>
  </button>
  <button class="carousel-control-next" type="button" data-bs-target="#carouselExample" data-bs-slide="next">
    <span class="carousel-control-next-icon" aria-hidden="true"></span>
    <span class="visually-hidden">Next</span>
  </button>
  
</div>
          

          </div>
        </div>
      </section>
    `;

    this._renderListAsync();
    this._loadLists();
  
   
  }

  
   

    // this._renderListAsync();
    // this._loadLists();
  

  protected onInit(): Promise<void> {
     
    this._selectedList = this.properties.selectedList || '';

    return this._getEnvironmentMessage().then(message => {
      this._environmentMessage = message;
    });
  }

  private _getListData(listName: string): Promise<ISPLists> {
    return this.context.spHttpClient.get(
      `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${listName}')/Items`,
      SPHttpClient.configurations.v1
    )
      .then((response: SPHttpClientResponse) => {
        return response.json();
      });
  }

  private _renderList(items: ISPList[]): void {
    const itemListElement = this.domElement.querySelector('.carousel-inner');

    if (itemListElement) {
      itemListElement.innerHTML = '';

      items.forEach((item: ISPList) => {
        const data = JSON.parse(item.Image);
        const imgurl = data.serverUrl + data.serverRelativeUrl;

        // <img src="${imgurl}" alt="Image" /><br/>
        itemListElement.innerHTML += `

        <div class="carousel-item active">

        <img src="${imgurl}"  class="d-block w-100"  alt="Image">

       </div>
        
        `;
      });
    }
  }

  private _loadLists(): void {
    this.context.spHttpClient.get(
      `${this.context.pageContext.web.absoluteUrl}/_api/web/lists?$filter=Hidden eq false`,
      SPHttpClient.configurations.v1
    )
      .then((response: SPHttpClientResponse) => {
        return response.json();
      })
      .then((data) => {
        const listSelector = this.domElement.querySelector('#listSelector') as HTMLSelectElement;
        if (listSelector) {
          data.value.forEach((list:ISPList) => {
            listSelector.innerHTML += `<option value="${list.Title}">${list.Title}</option>`;
          });
        }
      }).catch(()=>{
        console.log("this is not working now");
      });
  }

  private _renderListAsync():void{
    const listSelector = this.domElement.querySelector('#listSelector') as HTMLSelectElement;
    if (listSelector) {
      listSelector.addEventListener('change', (event: Event) => {
        const selectedList = listSelector.value;
        if (selectedList) {
          this._getListData(selectedList).then((response) => {
            this._renderList(response.value);
          }).catch(()=>{
            console.log("this is not working:");
          });
         }// else {
        //   this.domElement.querySelector('#itemList').innerHTML = '';
        // }
      });
    }
  }

  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) {
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
        .then(context => {
          let environmentMessage: string = '';
          switch (context.app.host.name) {
            case 'Office':
              environmentMessage = this.context.isServedFromLocalhost ? 'Running in Office locally' : 'Running in Office';
              break;
            case 'Teams':
            case 'TeamsModern':
              environmentMessage = this.context.isServedFromLocalhost ? 'Running in Teams locally' : 'Running in Teams';
              break;
            default:
              environmentMessage = 'Unknown environment';
          }
          return environmentMessage;
        });
    }
    return Promise.resolve(this.context.isServedFromLocalhost ? 'Running in SharePoint locally' : 'Running in SharePoint');
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }
    this._isDarkTheme = !!currentTheme.isInverted;
    const { semanticColors } = currentTheme;
    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }
  }

  protected get dataVersion(): Version {

    return Version.parse('1.0');

  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: 'Property Pane'
          },
          groups: [
            {
              groupName: 'Group 1',
              groupFields: [
                PropertyPaneTextField('description', {
                  label: 'Description'
                }),
                PropertyPaneDropdown('selectedList', {
                  label: 'Select a list',
                 
                  options: this.lists,
                  //onRender: this._onRenderPropertyPaneDropdown.bind(this),
                  // onChanged: this.onPropertyPaneFieldChanged.bind(this),
                  // selectedKey: this._selectedList,
                  
                  
                })
              ]
            }
          ]
        }
      ]
    };
  }


  // protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: string, newValue: string): void {
  //   if (propertyPath === 'selectedList' && newValue) {
  //     this._selectedList = newValue;
  //     this.render(); // Trigger a re-render when the selected list changes
  //   }
  // }

 






 }







// this is previous working if above code is not work then use it:-




// import { Version } from '@microsoft/sp-core-library';

// import {
//   type IPropertyPaneConfiguration,
//   PropertyPaneTextField
// } from '@microsoft/sp-property-pane';

// import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
// import type { IReadonlyTheme } from '@microsoft/sp-component-base';
// import { escape } from '@microsoft/sp-lodash-subset';

// import styles from './HelloWorldWebPart.module.scss';

// //import bootstrap from "bootstrap"

// import 'bootstrap/dist/css/bootstrap.css';

// import 'bootstrap/dist/js/bootstrap.bundle';





// import * as strings from 'HelloWorldWebPartStrings';

// export interface IHelloWorldWebPartProps {
//   description: string;
// }

// export default class HelloWorldWebPart extends BaseClientSideWebPart<IHelloWorldWebPartProps> {

//   private _isDarkTheme: boolean = false;
//   private _environmentMessage: string = '';

//   public render(): void {
//     this.domElement.innerHTML = `
//     <section class="${styles.helloWorld} ${!!this.context.sdks.microsoftTeams ? styles.teams : ''}">
//       <div class="${styles.welcome}">
//         <img alt="" src="${this._isDarkTheme ? require('./assets/welcome-dark.png') : require('./assets/welcome-light.png')}" class="${styles.welcomeImage}" />
//         <h2>Well done, ${escape(this.context.pageContext.user.displayName)}!</h2>
//         <div>${this._environmentMessage}</div>
//         <div>Web part property value: <strong>${escape(this.properties.description)}</strong></div>
//       </div>
//       <div>
//         <h3>Welcome to SharePoint Framework!</h3>
//         <p>
//         The SharePoint Framework (SPFx) is a extensibility model for Microsoft Viva, Microsoft Teams and SharePoint. It's the easiest way to extend Microsoft 365 with automatic Single Sign On, automatic hosting and industry standard tooling.
//         </p>
//         <h4>Learn more about SPFx development:</h4>
//           <ul class="${styles.links}">
//             <li><a href="https://aka.ms/spfx" target="_blank">SharePoint Framework Overview</a></li>
//             <li><a href="https://aka.ms/spfx-yeoman-graph" target="_blank">Use Microsoft Graph in your solution</a></li>
//             <li><a href="https://aka.ms/spfx-yeoman-teams" target="_blank">Build for Microsoft Teams using SharePoint Framework</a></li>
//             <li><a href="https://aka.ms/spfx-yeoman-viva" target="_blank">Build for Microsoft Viva Connections using SharePoint Framework</a></li>
//             <li><a href="https://aka.ms/spfx-yeoman-store" target="_blank">Publish SharePoint Framework applications to the marketplace</a></li>
//             <li><a href="https://aka.ms/spfx-yeoman-api" target="_blank">SharePoint Framework API reference</a></li>
//             <li><a href="https://aka.ms/m365pnp" target="_blank">Microsoft 365 Developer Community</a></li>
//           </ul>
//       </div>


//       <div id="carouselExample" class="carousel slide">
//   <div class="carousel-inner">
//     <div class="carousel-item active">
//       <img src="https://img.freepik.com/free-photo/wide-angle-shot-single-tree-growing-clouded-sky-during-sunset-surrounded-by-grass_181624-22807.jpg" class="d-block w-100" alt="first">
//     </div>
//     <div class="carousel-item">
//       <img src="https://images.ctfassets.net/hrltx12pl8hq/28ECAQiPJZ78hxatLTa7Ts/2f695d869736ae3b0de3e56ceaca3958/free-nature-images.jpg?fit=fill&w=1200&h=630" class="d-block w-100" alt="second">
//     </div>
//     <div class="carousel-item">
//       <img src="https://t4.ftcdn.net/jpg/05/47/97/81/360_F_547978128_vqEEUYBr1vcAwfRAqReZXTYtyawpgLcC.jpg" class="d-block w-100" alt="third">
//     </div>
//     <div class="carousel-item">
//       <img src="https://images.ctfassets.net/hrltx12pl8hq/28ECAQiPJZ78hxatLTa7Ts/2f695d869736ae3b0de3e56ceaca3958/free-nature-images.jpg?fit=fill&w=1200&h=630" class="d-block w-100" alt="fourth">
//     </div>
//   </div>
//   <button class="carousel-control-prev" type="button" data-bs-target="#carouselExample" data-bs-slide="prev">
//     <span class="carousel-control-prev-icon" aria-hidden="true"></span>
//     <span class="visually-hidden">Previous</span>
//   </button>
//   <button class="carousel-control-next" type="button" data-bs-target="#carouselExample" data-bs-slide="next">
//     <span class="carousel-control-next-icon" aria-hidden="true"></span>
//     <span class="visually-hidden">Next</span>
//   </button>
// </div>
//     </section>`;
//   }

//   protected onInit(): Promise<void> {
//     return this._getEnvironmentMessage().then(message => {
//       this._environmentMessage = message;
//     });
//   }



//   private _getEnvironmentMessage(): Promise<string> {
//     if (!!this.context.sdks.microsoftTeams) { // running in Teams, office.com or Outlook
//       return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
//         .then(context => {
//           let environmentMessage: string = '';
//           switch (context.app.host.name) {
//             case 'Office': // running in Office
//               environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
//               break;
//             case 'Outlook': // running in Outlook
//               environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
//               break;
//             case 'Teams': // running in Teams
//             case 'TeamsModern':
//               environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
//               break;
//             default:
//               environmentMessage = strings.UnknownEnvironment;
//           }

//           return environmentMessage;
//         });
//     }

//     return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
//   }

//   protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
//     if (!currentTheme) {
//       return;
//     }

//     this._isDarkTheme = !!currentTheme.isInverted;
//     const {
//       semanticColors
//     } = currentTheme;

//     if (semanticColors) {
//       this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
//       this.domElement.style.setProperty('--link', semanticColors.link || null);
//       this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
//     }

//   }

//   protected get dataVersion(): Version {
//     return Version.parse('1.0');
//   }

//   protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
//     return {
//       pages: [
//         {
//           header: {
//             description: strings.PropertyPaneDescription
//           },
//           groups: [
//             {
//               groupName: strings.BasicGroupName,
//               groupFields: [
//                 PropertyPaneTextField('description', {
//                   label: strings.DescriptionFieldLabel
//                 })
//               ]
//             }
//           ]
//         }
//       ]
//     };
//   }
// }
