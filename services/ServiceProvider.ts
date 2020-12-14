import { WebPartContext } from "@microsoft/sp-webpart-base";
import { MSGraphClient } from "@microsoft/sp-http";

export class ServiceProvider {
  public _graphClient: MSGraphClient;
  private spcontext: WebPartContext;
  public constructor(spcontext: WebPartContext) {
    this.spcontext = spcontext;
  }

  // This method calls the Graph api to get teams of the signed in user in sharepoint tenant

  public getTeams = async (): Promise<[]> => {

    this._graphClient = await this.spcontext.msGraphClientFactory.getClient(); //TODO
    let myTeams: [] = [];
    try {
      const teamsResponse = await this._graphClient.api('me/joinedTeams').version('v1.0').get();
      myTeams = teamsResponse.value as [];
    } catch (error) {
      console.log('Unable to get teams', error);
    }
    return myTeams;
  }

  // This method calls the graph api to get all the channels of a selected team

  public getChannels = async (teamID): Promise<[]> => {

    this._graphClient = await this.spcontext.msGraphClientFactory.getClient(); //TODO

    let channels: [] = [];
    try {
      const channelResponse = await this._graphClient.api('teams/' + teamID + '/channels').version('v1.0').get();
      channels = channelResponse.value as [];
    } catch (error) {
      console.log('unable to get channels', error);
    }

    return channels;

  }

  // This method calls the graph api to get all the messages for the selected team and a selected channel

  public getChannelMessages = async (teamID, channelId): Promise<[]> => {

    this._graphClient = await this.spcontext.msGraphClientFactory.getClient(); //TODO
    let messages: [] = [];
    var tenLatestMessages;
    try {
      const messagesResponse = await this._graphClient.api('teams/' + teamID + '/channels/' + channelId + "/messages/" ).version('beta').get();
      messages = messagesResponse.value as [];
      // Gets the ten latest mesages if the messages are more than 10
      if(messages.length>10)
      {
        tenLatestMessages =  messages.slice(0,10).reverse();
      }
      else
      {
        tenLatestMessages = messages.reverse();
      }
    } catch (error) {
      console.log('unable to get channel messages', error);
    }
    return tenLatestMessages;
  }

  //  This method calls the graph api to get replies for a particular message in a channel

  public getChannelMessageReplies = async (teamID, channelId, messageId): Promise<[]> => {

    this._graphClient = await this.spcontext.msGraphClientFactory.getClient(); //TODO

    let replies: [] = [];
    try {
      const replyResponse = await this._graphClient.api('teams/' + teamID + '/channels/' + channelId + "/messages/" + messageId+"/replies" ).version('beta').get();
      replies = replyResponse.value as [];
    } catch (error) {
      console.log('unable to get message replies', error);
    }
    return replies;
  }

  // This method calls the graph api to send message to a team's channel

  public sendMessage = async (teamId, channelId, message): Promise<[]> => {

    this._graphClient = await this.spcontext.msGraphClientFactory.getClient();

    try {

      var content = {
        "body": {
          "content": message
        }
      };
      const messageResponse =   await this._graphClient.api('/teams/' + teamId + '/channels/' + channelId + "/messages/")
        .version("beta").post(content);

      return messageResponse;

    } catch (error) {
      console.log('Unable to send message', error);
      return null;
    }

  }

  // This method calls the graph api to send a reply to a channel's message

  public sendReply = async (teamId, channelId, messageId, reply): Promise<[]> => {
    this._graphClient = await this.spcontext.msGraphClientFactory.getClient();
    try {

      var content = {
        "body": {
          "content": reply
        }
      };
      const replyResponse = await this._graphClient.api('/teams/' + teamId + '/channels/' + channelId + "/messages/"+messageId+"/replies")
        .version("v1.0").post(content);
      return replyResponse;

    } catch (error) {
      console.log('Unable to send reply', error);
      return null;
    }

  }
  // This method get the teams and channels to set in the properties of a property pane of a sharepoint webpart

  public getProperties = async (teamID): Promise<any[]> => {

    let myValues: any[] = [];
    let channels:any[] = [];
    let teams:any[] = [];
    teams = (await this.getTeams());
    myValues.push(teams);

    // Get channels for initial load when no team is selected
    if(teamID == 0)
    {
      channels = (await this.getChannels(teams[0].id));
    }
    // Get channels if user has selected any team from the property pane dropdown
    else
    {
      channels = (await this.getChannels(teamID));
    }
    myValues.push(channels);
    return myValues;
  }

  // This method calls the graph api to get all the users in the sharepoint tenant

  public getUsers= async():  Promise<any[]> =>
  {
    this._graphClient = await this.spcontext.msGraphClientFactory.getClient(); //TODO


    try {
      const userResponse = await this._graphClient.api('users').version('v1.0').get();

      return userResponse.value as [];
    } catch (error) {
      console.log('unable to get users', error);
    }
  }

  // This method calls the graph api to get photo of a particular user

  public  getPhoto = async (userId): Promise<any> =>
  {
    this._graphClient = await  this.spcontext.msGraphClientFactory.getClient(); //TODO
   try {
      const photoResponse = await  this._graphClient.api(['users',userId, 'photo/$value'].join('/')).responseType('blob').version('v1.0').get();
      return photoResponse;
       }
  catch (error) {

           console.log('unable to get photo', error);
      }

  }




}
