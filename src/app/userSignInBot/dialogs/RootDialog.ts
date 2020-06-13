import * as builder from "botbuilder";
import * as dialogs from "botbuilder-dialogs";

export class RootDialog extends dialogs.ComponentDialog {
    constructor(dialogId: string) {
        super(dialogId);
        
    }
    public async run(context: builder.TurnContext, accessor: builder.StatePropertyAccessor<dialogs.DialogState>) {
        const dialogSet = new dialogs.DialogSet(accessor);
        dialogSet.add(this);

        const dialogContext = await dialogSet.createContext(context);
        const results = await dialogContext.continueDialog();
        if (results.status === dialogs.DialogTurnStatus.empty) {
            await dialogContext.beginDialog(this.id);
        }
    }
}