import  * as React from 'react';
import styles from './MyTeams.module.scss';
import { IMyTeamsProps } from './IMyTeamsProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { ServiceProvider } from '../../../shared/services/ServiceProvider';
import { IMyTeamsWebPartProps } from '../MyTeamsWebPart';
import { Message } from '@microsoft/microsoft-graph-types';
import { transitionKeysAreEqual } from 'office-ui-fabric-react/lib/utilities/keytips/IKeytipTransitionKey';
import Avatar from 'react-avatar';




// States in the project
export interface IMyTeamsState {

  channelMessages: any;
  channelMessageReplies: any[];
  toggleText: string;
  values: any[];
  opened: boolean;
  readMore: boolean;

}
// Constant for the maximum number of characters in a message before show more link appears
const MAX_LENGTH = 250;
export default class MyTeams extends React.Component<IMyTeamsProps, IMyTeamsState> {

  // Local variables
  private serviceProvider;
  private messageTextRef;
  private replyTextRef;
  private hideIds = [];
  private count = 0;
  private userPhotos = [];
  private openReplyIds = [];
  private replies = [];
  private openReplyTextboxes = [];
  private readMoreMessageIds = [];
  private readMoreReplyIds = [];



  public constructor(props: IMyTeamsProps, state: IMyTeamsState) {
    super(props);

    this.serviceProvider = new ServiceProvider(this.props.context);

    // initializing states
    this.state = {

      channelMessages: [],
      channelMessageReplies: [],
      toggleText:"Collapse all",
      values: [],
      opened: false,
      readMore:false

    };

  }

  protected get disableReactivePropertyChanges(): boolean {
    return true;
  }

  // React component
  public render(): React.ReactElement<IMyTeamsProps> {

    return (

  <React.Fragment>
    {/* Show description property as heading */}
    <div className={styles.heading}>
     {this.props.description}</div>

      {/* Show all channel messages */}
        {this.state.channelMessages.map(
          (message: any, index: number) => (
            index >= (this.state.channelMessages.length - this.props.numberOfPosts) && message.from.user!=null &&

           <React.Fragment>
            <div className={styles.mainDiv}>
              {/* Show user photo, if user photo is not set use initials */}
              <div className={styles.avatar} >
                {this.userPhotos.map((photo:any) => (
                    photo.userId == message.from.user.id &&
                     <React.Fragment>
                        <Avatar src={photo.photoUrl} className={styles.profileImage} size="70px"  name={message.from.user.displayName}></Avatar>
                     </React.Fragment>
                ))}
              </div>

             <div className={styles.message}>
             <div className={styles["message-body"]}>
             <div className={styles["sender-container"]}>
             <div className={styles["message-sender"]}>
             <label ><b>{message.from.user.displayName}</b></label></div>
             <div className={styles["message-timestamp"]}>     <label >{(message.createdDateTime).split("T")[0].split("-").join("/")}</label>&nbsp;&nbsp;
             <label >{(message.createdDateTime).split("T")[1].split(".")[0]}</label></div>
             </div>
             <br></br><br></br>
             <div>
              {message.body.content.length > MAX_LENGTH ?
              (
              <div>
              {
              !(this.readMoreMessageIds.indexOf(message.id)!=-1) &&
              message.body.content.substring(0, MAX_LENGTH)}
              { !(this.readMoreMessageIds.indexOf(message.id)!=-1) &&
              <React.Fragment>
              ...<a className={styles["reply-link"]} href="#" onClick={() => this.showFullMessage(message.id)}>See more</a>
              </React.Fragment>}
              {(this.readMoreMessageIds.indexOf(message.id)!=-1) &&
              (
              <div>
              <label>{message.body.content}</label><br></br>
              <a className={styles["reply-link"]} href="#" onClick={() => this.showFullMessage(message.id)}>See Less</a>
              </div>
              )
              }
              </div>

             ) :

         <label>{message.body.content}</label>


      }
    </div>


              </div>
              {/* show/hide reply link for replies */}
              {this.state.channelMessageReplies.map(


                (reply: any) => (

                  reply.replyToId == message.id && this.count!=1 &&

                  <React.Fragment>
                {this.setCounter(true)}
                <div className={styles.collapseGrid}>
                <a href="#" className={styles.collapse} onClick={() => this.Collapse(message.id)}> Show/Hide Replies</a>
                </div>
              </React.Fragment>
                  ))}

              {/* Show replies for a message */}

            {this.state.channelMessageReplies.map(
                (reply: any) => (

           reply.replyToId == message.id && !(this.hideIds.indexOf(reply.replyToId)!=-1) &&

            <React.Fragment>
            <div className={styles["reply-grid-container"]} >
            {/* Get user photo for each reply */}
            <div className={styles.avatar} >
            {this.userPhotos.map((photo:any) => (
              photo.userId == reply.from.user.id &&
              <React.Fragment>
             <Avatar src={photo.photoUrl} className={styles.replyProfileImage} size="50px"  name={reply.from.user.displayName}></Avatar>
              </React.Fragment>
            ))}

             </div>
              <div className={styles["reply-message-body"]} id={reply.replyToId}  >

            <div className={styles["sender-container"]}>
            <div className={styles["message-sender"]}> <label ><b>{reply.from.user.displayName}</b></label></div>

            <div className={styles["message-timestamp"]}>    <label >{(reply.createdDateTime).split("T")[0].split("-").join("/")}</label>&nbsp;&nbsp;
              <label >{(reply.createdDateTime).split("T")[1].split(".")[0]}</label></div>
              </div>
            <br></br><br></br>

               <div>
               {reply.body.content.length > MAX_LENGTH ?
        (
          <div>
            {
            !(this.readMoreReplyIds.indexOf(reply.id)!=-1) &&

            reply.body.content.substring(0, MAX_LENGTH)}
            { !(this.readMoreReplyIds.indexOf(reply.id)!=-1) &&
            <React.Fragment>
            ...<a className={styles["reply-link"]} href="#" onClick={() => this.showFullReply(reply.id)}>See more</a>
            </React.Fragment>}
          {(this.readMoreReplyIds.indexOf(reply.id)!=-1) &&
             (
              <div>
              <label>{reply.body.content}</label><br></br>
              <a className={styles["reply-link"]} href="#" onClick={() => this.showFullReply(reply.id)}>See Less</a>
              </div>
             )

          }
          </div>

        ) :
        <label>{reply.body.content}</label>
      }
    </div>


               </div>
               </div>
             </React.Fragment>
                ))

                }

         {this.setCounter(false)}

         {this.props.channelDisplayName &&
          <React.Fragment>
          {!(this.openReplyIds.indexOf(message.id)!=-1) && !this.props.readOnly &&
          <React.Fragment>

         <div className={styles.replyDiv}  >
         <a href="#" className={styles["reply-link"]}  onClick={() => this.showReplyInput(message.id)}><img className={styles["reply-icon"]}  src={require('../../../icons/reply.png')}></img> Reply</a>
         </div>

          </React.Fragment>
      }
       {(this.openReplyIds.indexOf(message.id)!=-1) && !this.props.readOnly &&
       <React.Fragment>
            <div className={styles.newMessage} key={index}>
                <input className={styles.replyTextbox} placeholder='&nbsp;&nbsp;&nbsp;     Send a reply'   onChange={this.handleChange.bind(this, index)}  type="text" id="message" name="reply" hidden={this.props.readOnly}  autoFocus={true} onBlur={() => this.onFocusOut(message.id,index)}  />
            <input type="image" className={styles.photo} src={require('../../../icons/arrow-icon.png')} onMouseOver={e => (e.currentTarget.src = require('../../../icons/sent.png'))} onMouseOut={e => (e.currentTarget.src = require('../../../icons/arrow-icon.png'))}  onClick={() => this.sendReply(message.id,index)}  hidden={this.props.readOnly} title="Send" />
         </div>
         </React.Fragment>
       }
          </React.Fragment>
        }

         </div>
         </div>
            </React.Fragment>

          )
        )
        }

        {this.props.channelDisplayName &&
        <React.Fragment>


          <div className={styles.newMessage}>
          <input className={styles.textbox} placeholder='&nbsp;&nbsp;&nbsp;&nbsp;      Start a new Conversation'  ref={(elm) => { this.messageTextRef = elm; }} type="text" id="message" name="message" hidden={this.props.readOnly} />
          <input type="image" className={styles.photo} src={require('../../../icons/arrow-icon.png')} onMouseOver={e => (e.currentTarget.src = require('../../../icons/sent.png'))} onMouseOut={e => (e.currentTarget.src = require('../../../icons/arrow-icon.png'))} onClick={() => this.sendMesssage()} hidden={this.props.readOnly} title="Send" />
       </div>
        </React.Fragment>
      }

      </React.Fragment>

);
  }

 // Initial loading of messages and replies
  public componentDidMount(){

   this.getUserPhoto();

   this.getMessages();

  }

  // Message update on channel change
  public async componentDidUpdate(prevProps,prevState) {

    if ( prevProps.channelDisplayName!= this.props.channelDisplayName  )
     {
      console.log('Channel changed ');

      this.getMessages();
     }
}


// Get all the messages and their replies
private getMessages(){

  this.serviceProvider.
      getChannelMessages(this.props.teamDisplayName, this.props.channelDisplayName)
      .then(
        (result: any[]): void => {
          console.log(result);
          result.sort(this.sortByDate);
          this.setState({ channelMessages: result });

          let resultArr = [];
           result.forEach(function(message){

          this.serviceProvider.
            getChannelMessageReplies(this.props.teamDisplayName, this.props.channelDisplayName,message.id)
              .then(
                (replies: any[]): void => {
                  console.log(replies);
                  replies.forEach( (replyResult)=>{
                    resultArr.push(replyResult);

                  });
                  resultArr.sort(this.sortByDate);
                  this.setState({channelMessageReplies:resultArr});

                }
              )
              .catch(error => {
                console.log(error);
              });

          }.bind(this)); // end foreach

        }
      );

     }


// Send a new message
  private  sendMesssage()  {
    var _this = this;
    this.serviceProvider.
      sendMessage(this.props.teamDisplayName, this.props.channelDisplayName, this.messageTextRef.value)
      .then(
        (result: any[]): void => {

          console.log(result);
         _this.getMessages();
        _this.messageTextRef.value = '';
        _this.messageTextRef.focus();
        }
      )
      .catch(error => {
        console.log(error);
      });

  }

  // Show reply textbox
  private showReplyInput(messageId)
  {

    if(this.openReplyIds.indexOf(messageId)!=-1)
    {
      this.openReplyIds.splice(this.openReplyIds.indexOf(messageId),1);
      this.setState({opened : false});
    }
    else
    {
      this.openReplyIds.push(messageId);
      this.setState({opened : true});
    }

  }

  // Hide reply textbox when focus is out
  private onFocusOut(messageId,index)
  {

    this.replies[index] = this.state.values[index];
    if(typeof this.state.values[index] === 'undefined' || this.state.values[index] == '' || this.openReplyTextboxes[index].value == '' )
    {
     this.showReplyInput(messageId);
    }

  }

  private handleChange(i, event) {
    this.replyTextRef = event.target;
   this.openReplyTextboxes[i] = this.replyTextRef;
    let values = [this.state.values];

      values[i] = event.target.value;


    this.setState({ values });
 }

 // Send a reply to a message
  private sendReply(messageId,index) {
    var _this = this;
    var _index = index;


    this.serviceProvider.
      sendReply(this.props.teamDisplayName, this.props.channelDisplayName,messageId, this.replies[index])
      .then(
        (result: any[]): void => {

          _this.getMessages();


         _this.openReplyTextboxes[index].value = '';
         _this.openReplyTextboxes[index].focus();


        }
      )
      .catch(error => {
        console.log(error);
      });

  }

  private showFullMessage(messageId){

    if(this.readMoreMessageIds.indexOf(messageId)!=-1)
    {
      this.readMoreMessageIds.splice(this.readMoreMessageIds.indexOf(messageId),1);
      this.setState({readMore : false});
    }
    else
    {
      this.readMoreMessageIds.push(messageId);
      this.setState({readMore : true});
    }
      }


      private showFullReply(replyId){

        if(this.readMoreReplyIds.indexOf(replyId)!=-1)
        {
          this.readMoreReplyIds.splice(this.readMoreReplyIds.indexOf(replyId),1);
          this.setState({readMore : false});
        }
        else
        {
          this.readMoreReplyIds.push(replyId);
          this.setState({readMore : true});
        }
       }

 // Collapse all the replies for a message
  private Collapse(messageId){

// Show a meesage if its already collapsed
if(this.hideIds.indexOf(messageId)!=-1)
{
  this.hideIds.splice(this.hideIds.indexOf(messageId),1);
  this.setState({toggleText : "Collapse all"});
}

// Collapse messages otherwise
else
{
  this.hideIds.push(messageId);
  this.setState({toggleText : "Show Replies"});
}
  }

  private setCounter(reply)
  {
    if(reply)
    {
      this.count = 1;
    }
    else
    {
      this.count = 0;
    }

  }

  // Sort messages by created date
  private sortByDate( a, b ) {
    if ( a.createdDateTime < b.createdDateTime ){
      return -1;
    }
    if ( a.createdDateTime > b.createdDateTime ){
      return 1;
    }
    return 0;
  }

// get all users and their profile pictures
 private getUserPhoto()
 {
  this.serviceProvider.getUsers()
  .then( (result: any[]): void => {
    console.log(result);
    result.forEach(function(user){
      this.getPhoto(user.id);
    }.bind(this));
  });
 }

 // Gets profile picture of a user by id
 private  getPhoto(userId)
  {
    var _this = this;
    this.serviceProvider.getPhoto(userId)
     .then(  (blob: any): void => {

      if(blob === undefined)
      {

        _this.userPhotos.push({
          userId: userId,
          photoUrl: null
        });
      } // end if
      if(blob!== undefined)
      {
      var base64data = null;
      var reader = new FileReader();
      reader.readAsDataURL(blob);
      reader.onloadend  = () =>  {
       base64data = reader.result;

       _this.userPhotos.push({
        userId: userId,
        photoUrl: base64data.toString()
      }); // end push

    };// end reader

  } // end if


    })
    .catch(error => {
      console.log(error);
    });


}

}
