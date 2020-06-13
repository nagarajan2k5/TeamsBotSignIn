import { BotDeclaration, MessageExtensionDeclaration, PreventIframe } from "express-msteams-host";
import * as debug from "debug";
import * as botDialogs from "botbuilder-dialogs";
import * as builder from "botbuilder";
import { StatePropertyAccessor, CardFactory, TurnContext, MemoryStorage, ConversationState, ActivityTypes, TeamsActivityHandler } from "botbuilder";
import HelpDialog from "./dialogs/HelpDialog";
import WelcomeCard from "./dialogs/WelcomeDialog";
import SignInDialog from "./dialogs/SignInDialog";

// Initialize debug logging module
const log = debug("msteams");
const connectionName: string = process.env.OAuthConnectionName || '';
const MAIN_WATERFALL_DIALOG = "MainWaterfallDialog";
const OAUTH_PROMPT = "OAuthPrompt";

/**
 * Implementation for UserSignIn
 */
@BotDeclaration(
    "/api/messages",
    new MemoryStorage(),
    process.env.MICROSOFT_APP_ID,
    process.env.MICROSOFT_APP_PASSWORD)
@PreventIframe("/userSignInBot/userSignIn.html")
export class UserSignIn extends TeamsActivityHandler {
    private readonly conversationState: ConversationState;
    private readonly dialogs: botDialogs.DialogSet;
    private dialogState: StatePropertyAccessor<botDialogs.DialogState>;

    /**
     * The constructor
     * @param conversationState
     */
    public constructor(conversationState: ConversationState) {
        log("method: constructor");
        super();

        this.conversationState = conversationState;
        this.dialogState = conversationState.createProperty("dialogState");
        this.dialogs = new botDialogs.DialogSet(this.dialogState);
        this.dialogs.add(new HelpDialog("help"));
        this.dialogs.add(new SignInDialog("signin1"));

        this.dialogs.add(new botDialogs.OAuthPrompt(OAUTH_PROMPT, {
            connectionName,
            text: "Please sign in so I can show you your profile.",
            title: "Sign in",
            timeout: 300000
        }));

        this.dialogs.add(new botDialogs.WaterfallDialog(MAIN_WATERFALL_DIALOG, [
            this.checkTokenExists.bind(this),
            this.getTokenStep.bind(this),
            this.showUserToken.bind(this),
            this.showMeetingCard.bind(this)

        ]));

        // this.onTurn( async (context: TurnContext): Promise<void> => {
        // });

        // Set up the Activity processing
        this.onMessage(async (context: TurnContext): Promise<void> => {
            log("handler: onMessage");

            //Adding typing indicator to bot
            await this.sendTypingIndicatorAsync(context);

            // TODO: add your own bot logic in here
            switch (context.activity.type) {
                case ActivityTypes.Message:
                    let text = TurnContext.removeRecipientMention(context.activity);
                    text = text.toLowerCase();

                    if (text.startsWith("hello")) {
                        await context.sendActivity("Oh, hello to you as well!");
                        return;
                    } else if (text.startsWith("help")) {
                        const dc = await this.dialogs.createContext(context);
                        await dc.beginDialog("help");
                    } else if (text.startsWith("sign in")) {
                        const dc = await this.dialogs.createContext(context);
                        await dc.beginDialog(MAIN_WATERFALL_DIALOG);
                    }
                    else {
                        await context.sendActivity(`I\'m terribly sorry, but my master hasn\'t trained me to do anything yet...`);
                    }
                    break;
                default:
                    break;
            }

            // Save state changes
            return this.conversationState.saveChanges(context);
        });

        this.onConversationUpdate(async (context: TurnContext): Promise<void> => {
            log("handler: onConversationUpdate");
            if (context.activity.membersAdded && context.activity.membersAdded.length !== 0) {
                for (const idx in context.activity.membersAdded) {
                    if (context.activity.membersAdded[idx].id === context.activity.recipient.id) {
                        const welcomeCard = CardFactory.adaptiveCard(WelcomeCard);
                        await context.sendActivity({ attachments: [welcomeCard] });
                    }
                }
            }
        });

        this.onMessageReaction(async (context: TurnContext): Promise<void> => {
            log("handler: onMessageReaction");
            const added = context.activity.reactionsAdded;
            if (added && added[0]) {
                await context.sendActivity({
                    textFormat: "xml",
                    text: `That was an interesting reaction (<b>${added[0].type}</b>)`
                });
            }
        });

        this.onTokenResponseEvent(async (context: TurnContext): Promise<void> => {
            log("handler: onTokenResponseEvent");
        });
    }

    protected async handleTeamsSigninVerifyState(context: builder.TurnContext, query: builder.SigninStateVerificationQuery): Promise<void> {
        log("handler: handleTeamsSigninVerifyState");
        await context.sendActivity("Well!, token received");
        const dc = await this.dialogs.createContext(context);
        await dc.continueDialog();

    }

    public async run(context: builder.TurnContext) {
        log("handler: run");
        await super.run(context);

        // Save any state changes. The load happened during the execution of the Dialog.
        await this.conversationState.saveChanges(context, false);
    }



    private getTokenStep(stepContext: botDialogs.WaterfallStepContext) {
        log("step: getTokenStep");
        return stepContext.beginDialog(OAUTH_PROMPT);
        //stepContext.continueDialog();
    }

    protected async showUserToken(stepContext: botDialogs.WaterfallStepContext) {
        log("step: showUserToken");
        const tokenResponse: builder.TokenResponse = stepContext.result;
        if (tokenResponse) {
            await stepContext.context.sendActivity("Token response received");
        }
        else {
            await stepContext.context.sendActivity("Token response NULL");
        }
        await stepContext.next();
    }

    protected async showMeetingCard(stepContext: botDialogs.WaterfallStepContext) {
        log("step: showMeetingCard");
        await stepContext.context.sendActivity("send a card");
        await stepContext.endDialog();
    }

    private checkTokenExists(stepContext: botDialogs.WaterfallStepContext) {
        log("step: checkTokenExists")
        //Logic yet to add
        return stepContext.continueDialog();
    }

    private async sendTypingIndicatorAsync(context: builder.TurnContext) {
        await context.sendActivity({ type: ActivityTypes.Typing });
    }
}
