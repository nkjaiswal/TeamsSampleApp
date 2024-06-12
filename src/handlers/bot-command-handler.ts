import { TurnContext, TeamsInfo, CardFactory, Attachment } from "botbuilder";


interface BotCommandHandler {
    validateCommand(command: string): boolean;
    handleCommand(command: string, context: TurnContext): Promise<void | Attachment>;
}

class HelpBotCommandHandler implements BotCommandHandler {

    private helpReply = `
    Here are list of task you can do:
    - help: show this help message
    - echo: echo back the message you sent
    - team-profile: get the team profile
    - team-members: get the team members
    - member <memberId>: get the member details
    - meeting-info optional<meetingId>: get the meeting info
    visit https://whole-piglet-eminent.ngrok-free.app/api/notify to proactively message everyone who has previously messaged this bot.
    `;

    validateCommand(command: string): boolean {
        return command?.startsWith('help');
    }

    async handleCommand(command: string, context: TurnContext): Promise<void> {
        await context.sendActivity(this.helpReply);
    }
}

class EchoBotCommandHandler implements BotCommandHandler {
    validateCommand(command: string): boolean {
        return command?.startsWith('echo');
    }

    async handleCommand(command: string, context: TurnContext): Promise<void> {
        await context.sendActivity(`You said: ${command}`);
    }
}

class GetTeamProfileBotCommandHandler implements BotCommandHandler {
    validateCommand(command: string): boolean {
        return command?.startsWith('team-profile');
    }

    async handleCommand(command: string, context: TurnContext): Promise<void> {
        const teamId = context.activity.conversation.id;
        await context.sendActivity(`Team ID: ${teamId}`);
        const teamDetailsPromise = TeamsInfo.getTeamDetails(context);
        const teamChannelPromise = TeamsInfo.getTeamChannels(context);
        const [teamDetails, teamChannel] = await Promise.all([teamDetailsPromise, teamChannelPromise]);

        await context.sendActivity(`getTeamDetails(teamID): " ${JSON.stringify(teamDetails)}"`);
        await context.sendActivity(`getTeamChannels(): " ${JSON.stringify(teamChannel)}"`);

        const teamsChannelId = context.activity.channelData.channel.id;
        const type = context.activity.channelData.channel.type;
        await context.sendActivity(`Channel ID: ${teamsChannelId},\n Channel Type: ${type}`);

        const channelDetailsPromise = TeamsInfo.getTeamDetails(context, teamsChannelId);
        const channelChannelPromise = TeamsInfo.getTeamChannels(context, teamsChannelId);
        const [channelDetails, channelChannel] = await Promise.all([channelDetailsPromise, channelChannelPromise]);
        await context.sendActivity(`getTeamDetails(channelID): " ${JSON.stringify(channelDetails)}"`);
        await context.sendActivity(`getTeamChannels(channelID): " ${JSON.stringify(channelChannel)}"`);

    }
}

class GetTeamMembersBotCommandHandler implements BotCommandHandler {
    validateCommand(command: string): boolean {
        return command?.startsWith('team-members');
    }

    async handleCommand(command: string, context: TurnContext): Promise<void> {
        const members = await TeamsInfo.getPagedMembers(context, 100);
        await context.sendActivity(`getPagedMembers(teamID): ${JSON.stringify(members)}`);

        const teamMember = await TeamsInfo.getPagedTeamMembers(context, context.activity.conversation.id, 100);
        await context.sendActivity(`getPagedTeamMembers(teamID): ${JSON.stringify(teamMember)}`);

        const teamChannelId = context.activity.channelData.channel.id;
        const channelMembers = await TeamsInfo.getPagedTeamMembers(context, teamChannelId, 100);
        await context.sendActivity(`getPagedTeamMembers(channelID): ${JSON.stringify(channelMembers)}`);
    }
}

class GetMember implements BotCommandHandler {
    validateCommand(command: string): boolean {
        return command?.startsWith('member');
    }

    async handleCommand(command: string, context: TurnContext): Promise<void> {
        const memberId = command.split(' ')[1];
        console.log('memberId', memberId);
        const member = await TeamsInfo.getMember(context, memberId);
        await context.sendActivity(`getMember(memberId): ${JSON.stringify(member)}`);
    }
}

class GetMeetingInfo implements BotCommandHandler {
    validateCommand(command: string): boolean {
        return command?.startsWith('meeting-info');
    }

    async handleCommand(command: string, context: TurnContext): Promise<void> {
        const meetingID = command.split(' ')[1];
        const meeting = await TeamsInfo.getMeetingInfo(context, meetingID);
        await context.sendActivity(`getMeetingInfo(meetingID): ${JSON.stringify(meeting)}`);
        const participant = await TeamsInfo.getMeetingParticipant(context, meetingID);
        await context.sendActivity(`getMeetingParticipant(meetingID): ${JSON.stringify(participant)}`);
    }
}

const handlers = [
    new HelpBotCommandHandler(),
    new EchoBotCommandHandler(),
    new GetTeamProfileBotCommandHandler(),
    new GetTeamMembersBotCommandHandler(),
    new GetMember(),
    new GetMeetingInfo()
];

export class BotCommandHandlerManager {
    static async handleCommand(text: string, context: TurnContext) {
        const command = text.replace('<at>T3 Bot01</at> ', '').trim().toLowerCase();
        for (const handler of handlers) {
            if (handler.validateCommand(command)) {
                return await handler.handleCommand(command, context);
            }
        }
        await context.sendActivity(`I'm sorry, I don't recognize the command "${command}". Type "help" to see the list of commands I support.`);
    }
}