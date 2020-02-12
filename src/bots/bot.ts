import {
  ActivityHandler,
  BotState,
  ConversationState,
  StatePropertyAccessor,
  UserState
} from 'botbuilder';
import { Dialog, DialogState } from 'botbuilder-dialogs';
import { MainDialog } from '../dialogs/mainDialog';

export class SimonBot extends ActivityHandler {
  private conversationState: BotState;
  private userState: BotState;
  private dialog: Dialog;
  private dialogState: StatePropertyAccessor<DialogState>;
  /**
   *
   * @param {ConversationState} conversationState
   * @param {UserState} userState
   * @param {Dialog} dialog
   */
  constructor(
    conversationState: BotState,
    userState: BotState,
    dialog: Dialog
  ) {
    super();
    if (!conversationState) {
        throw new Error('[SimonBot]: Missing parameter. conversationState is required');
    }
    if (!userState) {
        throw new Error('[SimonBot]: Missing parameter. userState is required');
    }
    if (!dialog) {
        throw new Error('[SimonBot]: Missing parameter. dialog is required');
    }
    this.conversationState = conversationState as ConversationState;
    this.userState = userState as UserState;
    this.dialog = dialog;
    this.dialogState = this.conversationState.createProperty<DialogState>('DialogState');

    this.onMessage(async (context, next) => {
      
      // Run the Dialog with the new message Activity.
      await (this.dialog as MainDialog).run(context, this.dialogState);

      // By calling next() you ensure that the next BotHandler is run.
      await next();
    });

    this.onDialog(async (context, next) => {
        // Save any state changes. The load happened during the execution of the Dialog.
        await this.conversationState.saveChanges(context, false);
        await this.userState.saveChanges(context, false);

        // By calling next() you ensure that the next BotHandler is run.
        await next();
    });

    this.onMembersAdded(async (context, next) => {
      const membersAdded = context.activity.membersAdded;
      for (const member of membersAdded) {
        if (member.id !== context.activity.recipient.id) {
          const welcome = `Welcome to Simon Bot ${ member.name }. This Bot is a work in progress. At this time we have some dialogs working. Type anything to get started.`;
          await context.sendActivity(welcome);
        }
      }
      // By calling next() you ensure that the next BotHandler is run.
      await next();
    });
  }
}
