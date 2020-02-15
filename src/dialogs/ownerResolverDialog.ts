import {
  ComponentDialog,
  DialogTurnResult,
  OAuthPrompt,
  PromptValidatorContext,
  TextPrompt,
  WaterfallDialog,
  WaterfallStepContext
} from 'botbuilder-dialogs';
import { GraphHelper } from '../helpers/graphHelper';

const TEXT_PROMPT = 'textPrompt';
const WATERFALL_DIALOG = 'waterfallDialog';
const OAUTH_PROMPT = 'OAuthPrompt';

export class OwnerResolverDialog extends ComponentDialog {
  private static tokenResponse: any;
  
  private static async ownerPromptValidator(promptContext: PromptValidatorContext<string>): Promise<boolean> {
    if (promptContext.recognized.succeeded) {
      
      const owner: string = promptContext.recognized.value;
      if (!OwnerResolverDialog.validateEmail(owner)) {
        promptContext.context.sendActivity('Malformatted email adress.');
        return false;
      }
      
      if (!await GraphHelper.userExists(owner, OwnerResolverDialog.tokenResponse))  {
        promptContext.context.sendActivity('User doesn\'t exist.');
        return false;
      }

      return true;

    } else {
      return false;
    }
  }

  private static validateEmail(email: string): boolean {
    const re = /^((([a-z]|\d|[!#\$%&'\*\+\-\/=\?\^_`{\|}~]|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])+(\.([a-z]|\d|[!#\$%&'\*\+\-\/=\?\^_`{\|}~]|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])+)*)|((\x22)((((\x20|\x09)*(\x0d\x0a))?(\x20|\x09)+)?(([\x01-\x08\x0b\x0c\x0e-\x1f\x7f]|\x21|[\x23-\x5b]|[\x5d-\x7e]|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(\\([\x01-\x09\x0b\x0c\x0d-\x7f]|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF]))))*(((\x20|\x09)*(\x0d\x0a))?(\x20|\x09)+)?(\x22)))@((([a-z]|\d|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(([a-z]|\d|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])([a-z]|\d|-|\.|_|~|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])*([a-z]|\d|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])))\.)+(([a-z]|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(([a-z]|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])([a-z]|\d|-|\.|_|~|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])*([a-z]|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF]))){2,6}$/i;
    return re.test(email);
  }

  constructor(id: string) {
    super(id || 'ownerResolverDialog');
    
    this
        .addDialog(new TextPrompt(TEXT_PROMPT, OwnerResolverDialog.ownerPromptValidator.bind(this)))
        .addDialog(new OAuthPrompt(OAUTH_PROMPT, {
          connectionName: process.env.connectionName,
          text: 'Please Sign In',
          timeout: 300000,
          title: 'Sign In'
        }))
        .addDialog(new WaterfallDialog(WATERFALL_DIALOG, [
          this.promptStep.bind(this),
          this.initialStep.bind(this),
          this.finalStep.bind(this)
        ]));

    this.initialDialogId = WATERFALL_DIALOG;

  }

  /**
   * Prompt step in the waterfall. 
   */
  private async promptStep(stepContext: WaterfallStepContext): Promise<DialogTurnResult> {
      return await stepContext.beginDialog(OAUTH_PROMPT);
  }

  private async initialStep(stepContext: WaterfallStepContext): Promise<DialogTurnResult> {
    const tokenResponse = stepContext.result;
    if (tokenResponse && tokenResponse.token) {
      
      OwnerResolverDialog.tokenResponse = tokenResponse;

      const siteDetails = (stepContext.options as any).siteDetails;
      const promptMsg = 'Provide an owner email';

      if (!siteDetails.owner) {
        return await stepContext.prompt(TEXT_PROMPT, {
          prompt: promptMsg
        });
      } else {
        return await stepContext.next(siteDetails.owner);
      }
    }
    await stepContext.context.sendActivity('Login was not successful please try again.');
    return await stepContext.endDialog();
  }

  private async finalStep(stepContext: WaterfallStepContext): Promise<DialogTurnResult> {    
    const owner = stepContext.result;
    return await stepContext.endDialog(owner);
  }
}
