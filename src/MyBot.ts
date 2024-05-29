import { TeamsActivityHandler, TurnContext, Activity, TeamsInfo } from "botbuilder";
import { BotCommandHandlerManager } from "./handlers/bot-command-handler";
import fs from 'fs';

const url = 'https://whole-piglet-eminent.ngrok-free.app';

export class TeamsBot extends TeamsActivityHandler {

  private conversationReferences: {};

  constructor() {
    super();
    this.conversationReferences = JSON.parse(fs.readFileSync('conversationReferences.db').toString());

    this.handleOnConversationUpdate = this.handleOnConversationUpdate.bind(this);
    this.onConversationUpdate(this.handleOnConversationUpdate);

    this.handleOnMemberAdded = this.handleOnMemberAdded.bind(this);
    this.onMembersAdded(this.handleOnMemberAdded);

    this.handleOnMemberRemoved = this.handleOnMemberRemoved.bind(this);
    this.onMembersRemoved(this.handleOnMemberRemoved);

    this.handleOnMessage = this.handleOnMessage.bind(this);
    this.onMessage(this.handleOnMessage);
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