import { StatePropertyAccessor, TurnContext } from 'botbuilder';
import {
    ComponentDialog,
    DialogSet,
    DialogState,
    DialogTurnResult,
    DialogTurnStatus,
    WaterfallDialog,
    WaterfallStepContext
} from 'botbuilder-dialogs';

import { SiteDetails } from './siteDetails';
import { SiteDialog } from './siteDialog';

const SITE_DIALOG = 'siteDialog';
const MAIN_WATERFALL_DIALOG = 'waterfallDialog';

export class MainDialog extends ComponentDialog {
    
    constructor(id: string) {
        super(id);

        this.addDialog(new SiteDialog(SITE_DIALOG))
            .addDialog(new WaterfallDialog(MAIN_WATERFALL_DIALOG, [
                this.initialStep.bind(this),
                this.finalStep.bind(this)
            ]));

        this.initialDialogId = MAIN_WATERFALL_DIALOG;
    }

    /**
     * The run method handles the incoming activity (in the form of a DialogContext) and passes it through the dialog system.
     * If no dialog is active, it will start the default dialog.
     * @param {TurnContext} context
     */
    public async run(context: TurnContext, accessor: StatePropertyAccessor<DialogState>) {
        const dialogSet = new DialogSet(accessor);
        dialogSet.add(this);

        const dialogContext = await dialogSet.createContext(context);
        const results = await dialogContext.continueDialog();
        if (results.status === DialogTurnStatus.empty) {
            await dialogContext.beginDialog(this.id);
        }
    }

    /**
     * Initial step in the waterfall. This will kick of the site dialog
     */
    private async initialStep(stepContext: WaterfallStepContext): Promise<DialogTurnResult> {
        const siteDetails = new SiteDetails();
        return await stepContext.beginDialog('siteDialog', siteDetails);
    }

    /**
     * This is the final step in the main waterfall dialog.
     */
    private async finalStep(stepContext: WaterfallStepContext): Promise<DialogTurnResult> {
        if (stepContext.result) {
            const result = stepContext.result as SiteDetails;
            const msg = `I have created a ${ JSON.stringify(result) }`;
            await stepContext.context.sendActivity(msg);
        }
        return await stepContext.endDialog();

    }
}
