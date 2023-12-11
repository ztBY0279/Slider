import { BaseClientSideWebPart, } from '@microsoft/sp-webpart-base';
import styles from './HelloWorldWebPart.module.scss';
import {
    type IPropertyPaneConfiguration,
   
    PropertyPaneTextField,
    PropertyPaneDropdown,
    //PropertyPaneDropdownOptionType,
    IPropertyPaneDropdownOption,
  
   
    
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
  description:string;
  caption:string;
  link:string;
}

export interface IHelloWorldWebPartProps {

  description: string;
  selectedList: string;
  caption: string;
  Description: string;
  link: string;
 // description1:string;
 // caption1: string;   // New property for image caption
 // link1: string;      // New property for custom navigation URL

 // for custom width and height:-

imageWidth: number; // New property for controlling image width
imageHeight: number; // New property for controlling image height
transitionEffect:string;
visualEffect: string;


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
              <!--    <select id="listSelector" onchange="fetchListItems(this.value)"> -->
              <select id="listSelector" >
                <option value="" id = "somenew">-- Select a list  --</option>
              </select>
            </div>
            <div class="" id="itemList"></div>

            

            <div id="carouselExample" class="carousel slide fade zoom ${this.properties.visualEffect} ${styles.fade} ${styles.zoom} ${styles.slide}">
 
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
        <p>this is working now:</p>
        <p>this is again working now:</p>
      </section>
    `;

// fade in ,zoom and slide effect code :- 

// Update the transition effect class based on user preferences or dynamic settings
const carouselElement = this.domElement.querySelector('#carouselExample');
if (carouselElement) {
  // Remove existing transition classes before adding the new one
  carouselElement.classList.remove('fade', 'slide', 'zoom');
  carouselElement.classList.add(this.properties.transitionEffect);
}
  

    this._renderListAsync();
    this._loadLists();

    const listSelector = this.domElement.querySelector('#listSelector') as HTMLSelectElement;
  if (listSelector) {
    listSelector.value = this.properties.selectedList || '';
  }
  
   
  }

  
  


  protected onInit(): Promise<void> {
    this._selectedList = this.properties.selectedList || '';
  
    return Promise.all([
      this._getEnvironmentMessage().then(message => {
        this._environmentMessage = message;
      }),
      this._loadLists()
    ]).then(() => {
      this._renderListAsync();
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

    console.log("items in _renderList() method:",items);

    if(items.length === 0){
      alert("list does not contain any data:");
    }

    if(items === null || items === undefined){
      console.log(items);
      alert("the list is empty:");
    }
    const itemListElement = this.domElement.querySelector('.carousel-inner');

    if (itemListElement) {
      itemListElement.innerHTML = '';

      items.forEach((item: ISPList,index:number) => {

        console.log("the value of item.image is give back: ",item.Image);

        if(item.Image === undefined || item.Image === null){
          alert("image column is not exist: ");
        }

        console.log("item in foreach of items",item);

        if(item === null || item === undefined){
          alert("the list is empty:");
        }
        const data = JSON.parse(item.Image);
        const imgurl = data.serverUrl + data.serverRelativeUrl;
        console.log(imgurl);

        // caption and description are added here:-

        if(imgurl === undefined){
           
          alert("the list is typecally empty it does not contain Images:");

        }

        const caption = item.caption; // Change this to the property you want as a caption
      const description = item.description; // Change this to the property you want as a description
         
     const Link = item.link;
      // active class:-
     // const activeClass = index === 0 ? 'active' : '';
     console.log("the value of Link is: ",Link);
   

        // <img src="${imgurl}" alt="Image" /><br/>
        itemListElement.innerHTML += `

        <div class="carousel-item active ${this.properties.visualEffect}">

        <img src="${imgurl}"  class="d-block w-100" style="width: ${this.properties.imageWidth}px; height: ${this.properties.imageHeight}px;" alt="Image">

        <div class="carousel-caption d-none d-md-block">
            <h5 style="color:black;" class = "${styles.caption} ">this is caption value : ${caption}</h5>
            <p style="color:black;" class = "${styles.description}">this is description value: ${description}</p>
            <a href="${Link}" class="btn btn-primary" target="_blank">Learn More</a>
        </div>

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

        this.lists = data.value.map((list: ISPList) => {
          return { key: list.Title, text: list.Title };
        });
        this.context.propertyPane.refresh();
         


      }).catch(()=>{
        console.log("this is not working now");
      });
  }



  private _renderListAsync():void{
    const listSelector = this.domElement.querySelector('#listSelector') as HTMLSelectElement;
    if (listSelector) {
      listSelector.value = this._selectedList;

      listSelector.addEventListener('change', (event: Event) => {

        console.log("this._renderListAsync method is called now:");
        console.log("list selector value is :",listSelector.value);

        const selectedList = listSelector.value;

        console.log("the length of selectedList: ",selectedList.length);

        console.log("the selected list itself now: ",selectedList);

        console.log("the value of selectedlist is :",selectedList);


        if (selectedList) {
          this._selectedList = selectedList;
          this._getListData(selectedList).then((response) => {
            this._renderList(response.value);
          }).catch(()=>{
            console.log("this is not working:");
          });
           
          this.properties.selectedList = selectedList;
           this.context.propertyPane.refresh();

         }// else {
        //   this.domElement.querySelector('#itemList').innerHTML = '';
        // }
      });
    }
  }



  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: string, newValue: string): void {
    if (propertyPath === 'selectedList' && newValue) {
      this._selectedList = newValue;
  
      // Update the main screen dropdown when PropertyPane dropdown changes
      const listSelector = this.domElement.querySelector('#listSelector') as HTMLSelectElement;
      if (listSelector) {
        listSelector.value = newValue;
      }
     const somenewOption = this.domElement.querySelector("#somenew") as HTMLOptionElement;
     console.log(somenewOption);
     console.log(somenewOption.innerHTML);
     
     console.log("the new value is :",newValue);
     somenewOption.innerHTML = newValue;
     somenewOption.innerText = newValue;
     console.log("somenewoption innerthml after new value become:",somenewOption.innerHTML);
     console.log("the innerText property:- ",somenewOption.innerText);

    //  somenewOption.innerText = newValue;
      // Trigger the rendering of the list asynchronously

     // this._loadLists();
      this._getListData(newValue).then((response) => {
        this._renderList(response.value);
      }).catch(() => {
        console.log("Error fetching list items");
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
                  options:this.lists
                }),

                PropertyPaneTextField('caption', {
                  label: 'Caption',
                  value:"some text1"
                }),
                PropertyPaneTextField('Description', {
                  label: 'Description',
                  value:"some text2"
                }),
                PropertyPaneTextField('link', {
                  label: 'Link',
                  value:this.properties.Description
                }),

               

                


              ]
            },

            // for image height and width 
            {
              groupName: 'Group 2',
              groupFields: [
                // ... (existing fields)
          
                PropertyPaneTextField('imageWidth', {
                  label: 'Image Width (in pixels)',
                  value: '300', // Set a default value or retrieve from a property
                }),
                PropertyPaneTextField('imageHeight', {
                  label: 'Image Height (in pixels)',
                  value: '200', // Set a default value or retrieve from a property
                }),

                PropertyPaneDropdown('transitionEffect', {
                  label: 'Transition Effect',
                  options: [
                    { key: 'fade', text: 'Fade' },
                    { key: 'slide', text: 'Slide' },
                    { key: 'zoom', text: 'Zoom' },
                  ],
                }),

                PropertyPaneDropdown('visualEffect', {
                  label: 'Visual Effect',
                  options: [
                    { key: 'none', text: 'None' },
                    { key: 'filter1', text: 'Filter 1' },
                    { key: 'filter2', text: 'Filter 2' },
                    // Add more options as needed
                  ],
                })
                

              ],
            },

          ]
        }
      ]
    };
  }





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
