import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneDropdown,
  IPropertyPaneDropdownOption,
  PropertyPaneToggle,
  PropertyPaneSlider
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart, WebPartContext } from '@microsoft/sp-webpart-base';




import MyTeams from './components/MyTeams';
import { IMyTeamsProps } from './components/IMyTeamsProps';
import { IMyTeamsState} from './components/MyTeams';
import { ServiceProvider } from '../../shared/services/ServiceProvider';
import { ThemeSettingName, TooltipHostBase } from 'office-ui-fabric-react';

// Properties used in the sharepoint property pane

export interface IMyTeamsWebPartProps {
  description: string;
  teamDisplayName: string;
  channelDisplayName: string;
  readOnly:boolean;
  numberOfPosts:number;
}






export default class MyTeamsWebPart extends BaseClientSideWebPart<IMyTeamsWebPartProps> {

  // Local variables

  private serviceProvider;
  private teams : IPropertyPaneDropdownOption[];
  private channels : IPropertyPaneDropdownOption[];
  private teamsDropdownDisabled: boolean = true;
  private channelsDropdownDisabled: boolean = true;

  public render(): void {

    const element: React.ReactElement<IMyTeamsProps> = React.createElement(
      MyTeams,
      {
        // Setting local variables with the property values
        context: this.context,
        teamDisplayName: this.properties.teamDisplayName,
        channelDisplayName: this.properties.channelDisplayName,
        readOnly:this.properties.readOnly,
        numberOfPosts: this.properties.numberOfPosts,
        description: this.properties.description
       }
    );

    this.serviceProvider = new ServiceProvider(this.context);

    ReactDom.render(element, this.domElement);


  }

  // Initialize properties in the webpart property pane

  public onInit<T>(): Promise<T> {

    this.teams = [];
    this.channels = [];
    this.serviceProvider = new ServiceProvider(this.context);

    return new Promise<T>(
      (resolve: (args: T) => void, reject: (error: Error) => void) => {

        // get initial properties for teams and channels
        if(this.properties.teamDisplayName == ''){
        this.serviceProvider.getProperties(0).then(result => result.forEach((resultProperties,index) => {
          // Get all teams
          if(index == 0 && this.properties.teamDisplayName == '')
          {
            this.properties.teamDisplayName = resultProperties[0].id;
            resultProperties.forEach(team => {
              this.teams.push(<IPropertyPaneDropdownOption>{
                text: team.displayName,
                key: team.id
              });
            });
          }
          // Get all channels
          else if(index == 1 && this.properties.channelDisplayName == '')
          {
            this.properties.channelDisplayName = resultProperties[0].id;
            resultProperties.forEach(channel => {
              this.channels.push(<IPropertyPaneDropdownOption>{
                text: channel.displayName,
                key: channel.id
              });
            });
          }



        resolve(undefined);
     }));

    }
    else
    {
      this.serviceProvider.getProperties(this.properties.teamDisplayName).then(result => result.forEach((resultProperties,index) => {
        // Get all teams
        if(index == 0 )
        {
           resultProperties.forEach(team => {
            this.teams.push(<IPropertyPaneDropdownOption>{
              text: team.displayName,
              key: team.id
            });
          });
        }
        // Get all the channels for the selected team
        else if(index == 1)
        {

          resultProperties.forEach(channel => {
            this.channels.push(<IPropertyPaneDropdownOption>{
              text: channel.displayName,
              key: channel.id
            });
          });
        }



      resolve(undefined);
   }));

    }

    });

  }







  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }



  // get channels for a team
  private getChannels<T>(teamId): Promise<T> {

    this.channels = [];
    this.serviceProvider = new ServiceProvider(this.context);

    return new Promise<T>((resolve: (args: T) => void, reject: (error: Error) => void) => {


      this.serviceProvider.getChannels(teamId).then(result => result.forEach(channel => {
        this.channels.push(<IPropertyPaneDropdownOption>{
          text: channel.displayName,
          key: channel.id

        });
        if(this.properties.channelDisplayName == '')
        {
          this.properties.channelDisplayName = this.channels[0].key.toString();

        }
      }));
      resolve(undefined);
    });

  }


  // Set properties if any field in the property pane changes

  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): void {

    if(propertyPath==='teamDisplayName' && newValue)
    {
      this.properties.teamDisplayName = newValue;
      this.getChannels(newValue);
      this.properties.channelDisplayName = this.channels[0].key.toString();
      this.render();

    }

    if(propertyPath==='channelDisplayName' && newValue)
    {
     this.properties.channelDisplayName = newValue;
     this.render();
    }

    if(propertyPath==='readOnly' && newValue)
    {
     this.properties.readOnly = newValue;
       this.render();
    }

    if(propertyPath==='numberOfPosts' && newValue)
    {
     this.properties.numberOfPosts = newValue;
       this.render();
    }

  }

  // Configure property pane

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
           description: '',

          },
          groups: [
            {

                groupName: 'Web Part Title',
              groupFields: [
                PropertyPaneTextField('description', {
                  label: 'Title'
                })]},
                {

                  groupName: 'Configuration',
                  groupFields: [
                PropertyPaneDropdown('teamDisplayName', {
                  label: 'Team',
                  options: this.teams,
                  selectedKey: this.teams[0].key

                }),

                PropertyPaneDropdown('channelDisplayName', {
                  label: 'Channel',
                  options: this.channels,
                  selectedKey: this.channels[0].key,


                }),


                PropertyPaneSlider('numberOfPosts',{
                  label: 'Max number of posts to show',
                  min:1,
                  max:10,
                  value:5,
                  showValue:true,
                  step:1
                    }),
                    PropertyPaneToggle('readOnly',{

                      label: 'Read Only?'

                    } )

              ]


            }
          ]
        }
      ]
    };
  }
}
