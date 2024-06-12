import { TurnContext, CardFactory, TeamsInfo, AttachmentLayout, MessagingExtensionResultType, TaskModuleContinueResponse } from "botbuilder";


export const notifyAllCardBuilder = (notificationCount: number, conversationCount: number) => {
    return CardFactory.heroCard('Hero Card Title', `Total Notifications: ${notificationCount}, Total Conversations: ${conversationCount}`);
}

export const SampleCardWithButton = (context: TurnContext) => {
    const executionId = context.activity.value.data.execution_id;
    const card = CardFactory.heroCard(
        'Sample Card',
        'This is a sample card with a button',
        [],
        [
            {
                type: 'invoke',
                title: 'Click Me',
                value: { key: 'sample_card_button_click_me', executionId }
            }
        ]
    );
    return {
        composeExtension: {
            type: 'result' as MessagingExtensionResultType,
            attachmentLayout: 'list' as AttachmentLayout,
            attachments: [card]
        }
    };
}

export const validateIfAppInstalledWithJIT = async (context: TurnContext) => {
    try {
        const member = await getSingleMember(context);
        return {
            task: {
                type: 'continue',
                value: {
                    card: GetAdaptiveCardAttachment(),
                    height: 400,
                    title: `Hello ${member}`,
                    width: 300
                }
            } as TaskModuleContinueResponse
        };
    } catch (error) {
        if (error.code === 'BotNotInConversationRoster') {
            const installCard = CardFactory.adaptiveCard({
                actions: [
                    {
                        type: 'Action.Submit',
                        title: 'Continue',
                        data: { msteams: { justInTimeInstall: true } }
                    }
                ],
                body: [
                    {
                        text: 'Looks like you have not used Action Messaging Extension app in this team/chat. Please click **Continue** to add this app.',
                        type: 'TextBlock',
                        wrap: true
                    }
                ],
                type: 'AdaptiveCard',
                version: '1.0'
            });
            return {
                composeExtension: {
                    type: 'result' as MessagingExtensionResultType,
                    attachmentLayout: 'list' as AttachmentLayout,
                    attachments: [installCard]
                }
            };
        }
    }

}

const GetAdaptiveCardAttachment = () => {
    return CardFactory.adaptiveCard({
        actions: [{ type: 'Action.Submit', title: 'Close' }],
        body: [
            {
                text: 'This app is installed in this conversation. You can now use it to do some great stuff!!!',
                type: 'TextBlock',
                isSubtle: false,
                wrap: true
            }
        ],
        type: 'AdaptiveCard',
        version: '1.0'
    });
}

const getSingleMember = async (context: TurnContext) => {
    try {
        const member = await TeamsInfo.getMember(
            context,
            context.activity.from.id
        );
        return member.name;
    } catch (e) {
        if (e.code === 'MemberNotFoundInConversation') {
            context.sendActivity('Member Not Found');
            return e.code;
        }
        throw e;
    }
}