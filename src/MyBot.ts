import { TeamsActivityHandler, TurnContext, Activity, TeamsInfo, MessagingExtensionAction, MessagingExtensionActionResponse, BotFrameworkAdapter, Attachment, AttachmentLayout, MessagingExtensionResultType } from "botbuilder";
import { BotCommandHandlerManager } from "./handlers/bot-command-handler";
import fs from 'fs';
import { notifyAllCardBuilder, validateIfAppInstalledWithJIT, SampleCardWithButton } from "./handlers/adaptive-card-factory";

const url = 'https://whole-piglet-eminent.ngrok-free.app';

export class TeamsBot extends TeamsActivityHandler {

  private conversationReferences: {};
  private adapter: BotFrameworkAdapter;

  constructor(adapter: BotFrameworkAdapter) {
    super();
    this.adapter = adapter;
    this.conversationReferences = JSON.parse(fs.readFileSync('conversationReferences.db').toString());

    // test cases https://loop.microsoft.com/p/eyJ1IjoiaHR0cHM6Ly9taWNyb3NvZnQuc2hhcmVwb2ludC1kZi5jb20vY29udGVudHN0b3JhZ2UvQ1NQXzQ0OGJlMDcxLTY2YzUtNDUzMy05M2U3LTg0M2FlMmMxY2U4ZT9uYXY9Y3owbE1rWmpiMjUwWlc1MGMzUnZjbUZuWlNVeVJrTlRVRjgwTkRoaVpUQTNNUzAyTm1NMUxUUTFNek10T1RObE55MDRORE5oWlRKak1XTmxPR1VtWkQxaUpUSXhZMlZEVEZKTlZtMU5NRmRVTlRSUk5qUnpTRTlxY25WSFVsazBTVmRJZEVkcExXNTZSbXhHTmxWTFVsRnBSRzQzT1dGaFZsSTJhbWRKVEcxbVUzUlJhU1ptUFRBeFFWaEhSMUZWUjFCRFdFOVZUVTFRUkVaQ1FqSmFXVWRGV1RaQlUwdFlTakltWXowbE1rWW1ZVDFNYjI5d1FYQndKbkE5SlRRd1pteDFhV1I0SlRKR2JHOXZjQzF3WVdkbExXTnZiblJoYVc1bGNpWjRQU1UzUWlVeU1uY2xNaklsTTBFbE1qSlVNRkpVVlVoNGRHRlhUbmxpTTA1MldtNVJkV015YUdoamJWWjNZakpzZFdSRE1XdGFhVFZxWWpJeE9GbHBSbXBhVlU1TlZXc3hWMkpWTUhkV01WRXhUa1pGTWs1SVRrbFVNbkI1WkZWa1UxZFVVa3BXTUdnd1VqSnJkR0p1Y0VkaVJWa3lWbFYwVTFWWGJFVmlhbU0xV1ZkR1YxVnFXbkZhTUd4TllsZGFWR1JHUm5CbVJFRjRVVlpvU0ZJeFJsWlJlbFpFVm14c1NGRlZPVTVWYTNCaFVsVndRMVJWYkV0VU1FWmFWVlZvVEU0eFJTVXpSQ1V5TWlVeVF5VXlNbWtsTWpJbE0wRWxNakl4TURVMVlqbGtNQzAwWm1JNUxUUTJaVFl0WVRFMVl5MHdNelV4TVdKa1lUSXpaV01sTWpJbE4wUT0ifQ%3D%3D?ct=1718171190850&&LOF=1

    this.handleOnConversationUpdate = this.handleOnConversationUpdate.bind(this);
    this.onConversationUpdate(this.handleOnConversationUpdate);

    this.handleOnMemberAdded = this.handleOnMemberAdded.bind(this);
    this.onMembersAdded(this.handleOnMemberAdded);

    this.handleOnMemberRemoved = this.handleOnMemberRemoved.bind(this);
    this.onMembersRemoved(this.handleOnMemberRemoved);

    this.handleOnMessage = this.handleOnMessage.bind(this);
    this.onMessage(this.handleOnMessage);
    this.onMessageDelete(this.getDefaultEventHandler('MessageDelete'));
    this.onMessageUpdate(this.getDefaultEventHandler('MessageUpdate'));

    this.onInstallationUpdate(this.getDefaultEventHandler('InstallationUpdate'));

    this.onReactionsAdded(this.getDefaultEventHandler('ReactionsAdded'));
    this.onReactionsRemoved(this.getDefaultEventHandler('ReactionsRemoved'));
    this.onMessageReaction(this.getDefaultEventHandler('MessageReaction'));


    this.onTeamsChannelCreated(this.getDefaultEventHandler('TeamsChannelCreated'));
    this.onTeamsChannelDeleted(this.getDefaultEventHandler('TeamsChannelDeleted'));
    this.onTeamsChannelRenamed(this.getDefaultEventHandler('TeamsChannelRenamed'));
    this.onTeamsTeamArchived(this.getDefaultEventHandler('TeamsTeamArchived'));
    this.onTeamsTeamDeleted(this.getDefaultEventHandler('TeamsTeamDeleted'));
    this.onTeamsTeamHardDeleted(this.getDefaultEventHandler('TeamsTeamHardDeleted'));
    this.onTeamsTeamRestored(this.getDefaultEventHandler('TeamsTeamRestored'));
    this.onTeamsTeamUnarchived(this.getDefaultEventHandler('TeamsTeamUnarchived'));

  }

  public async handleTeamsMessagingExtensionCardButtonClicked(context: TurnContext, obj: any): Promise<void> {
    await context.sendActivity(`handleTeamsMessagingExtensionCardButtonClicked: ${JSON.stringify(obj)}`);
  }

  public async handleTeamsMessagingExtensionSubmitAction(context: TurnContext, action: MessagingExtensionAction): Promise<MessagingExtensionActionResponse> {
    const command = action.commandId;
    switch (command) {
      case 'notify_all':
        const times = action.data.times;
        this.notifyAll();
        const card = notifyAllCardBuilder(times, Object.keys(this.getAllConversations()).length);
        return this.getComposeExtensionCardResponse(card);
      case 'jit':
        return validateIfAppInstalledWithJIT(context);
      case 'sample':
        return SampleCardWithButton(context);
    }
  }

  private getComposeExtensionCardResponse(attachment: Attachment) {
    return {
      composeExtension: {
        type: 'result' as MessagingExtensionResultType,
        attachmentLayout: 'list' as AttachmentLayout,
        attachments: [attachment]
      }
    };
  }

  public async notifyAll() {
    for (const conversationReference of Object.values(this.getAllConversations())) {
      await this.adapter.continueConversation(conversationReference, async (context) => {
        try {
          await context.sendActivity('proactive hello');
        } catch (error) {
          if (error.code === 'BotNotInConversationRoster' || error.code === 'ConversationNotFound') {
            console.log('BotNotInConversationRoster or ConversationNotFound error');
          } else {
            console.error(error);
          }
        }

      });
    }
  }

  public getAllConversations() {
    return this.conversationReferences;
  }

  private getDefaultEventHandler(name: string) {
    return async (context: TurnContext, next: () => Promise<void>) => {
      await context.sendActivity(`Event: ${name} triggered. Context: ${JSON.stringify(context)}`);
      await next();
    };
  }

  private async handleOnMessage(context: TurnContext, next: () => Promise<void>) {
    this.addConversationReference(context.activity);
    await BotCommandHandlerManager.handleCommand(context.activity.text, context);
    await next();
  }

  private async handleOnConversationUpdate(context: TurnContext, next: () => Promise<void>) {
    this.addConversationReference(context.activity);
    await context.sendActivity(`onConversationUpdate event detected: ${context.activity.channelData.team.id}`);
    await next();
  }

  private async handleOnMemberAdded(context: TurnContext, next: () => Promise<void>) {
    const membersAdded = context.activity.membersAdded;
    for (let cnt = 0; cnt < membersAdded.length; cnt++) {
      const member = await TeamsInfo.getMember(context, membersAdded[cnt].id)
      await context.sendActivity(`Welcome ${member.name}\n@T3 Bot help to see the list of commands`);
    }
    await next();
  }

  private async handleOnMemberRemoved(context: TurnContext, next: () => Promise<void>) {
    const membersRemoved = context.activity.membersRemoved;
    for (let cnt = 0; cnt < membersRemoved.length; cnt++) {
      await context.sendActivity(`Sorry to see you go Unknown(${membersRemoved[cnt].id})`);
    }
    await next();
  }

  private addConversationReference(activity: Partial<Activity>) {
    const conversationReference = TurnContext.getConversationReference(activity);
    this.conversationReferences[conversationReference.conversation.id] = conversationReference;
    fs.writeFileSync('conversationReferences.db', JSON.stringify(this.conversationReferences, null, 2));
  }

}