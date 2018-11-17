import { override } from '@microsoft/decorators';
import {
  BaseListViewCommandSet,
  Command,
  IListViewCommandSetListViewUpdatedParameters,
  IListViewCommandSetExecuteEventParameters
} from '@microsoft/sp-listview-extensibility';
import { Dialog } from '@microsoft/sp-dialog';
import { updateTodoItem } from '../../webparts/todo/api';

export default class TodoCommandSet extends BaseListViewCommandSet<{}> {

  @override
  public onListViewUpdated(event: IListViewCommandSetListViewUpdatedParameters): void {
    const command: Command = this.tryGetCommand('COMPLETE_TODO_ITEM_COMMAND');
    if (command) {
      command.visible = (
        event.selectedRows.length === 1 &&
        event.selectedRows[0].getValueByName('Completed') === 'No'
      );
    }
  }

  @override
  public onExecute(event: IListViewCommandSetExecuteEventParameters): void {
    switch (event.itemId) {
      case 'COMPLETE_TODO_ITEM_COMMAND':
        const todoItemId: number = Number(event.selectedRows[0].getValueByName('ID'));
        const title: number = event.selectedRows[0].getValueByName('Title');

        Dialog.alert(`Completing the todo item: ${title}`);
        updateTodoItem(this.context.serviceScope, todoItemId, /* completed */ true)
          .then(() => {
            // There is no way to refresh the list view yet. Refresh the page as a workaround.
            // See https://github.com/SharePoint/sp-dev-docs/issues/1449
            window.location.reload();
          });
        break;
      default:
        throw new Error('Unknown command');
    }
  }
}
