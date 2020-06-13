import { Dialog, DialogContext, DialogTurnResult, OAuthPrompt } from "botbuilder-dialogs";

export default class SignInDialog extends Dialog {
    constructor(dialogId: string) {
        super(dialogId);
        
    }

    public async beginDialog(context: DialogContext, options?: any): Promise<DialogTurnResult> {
        context.context.sendActivity(`Sign dialog begin`);

        return await context.endDialog();
    }
}
