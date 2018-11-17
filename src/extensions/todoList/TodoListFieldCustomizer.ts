import { override } from '@microsoft/decorators';
import {
  BaseFieldCustomizer,
  IFieldCustomizerCellEventParameters
} from '@microsoft/sp-listview-extensibility';

const DAYS: number = 24 * 60 * 60 * 1000;

export default class TodoListFieldCustomizer extends BaseFieldCustomizer<{}> {
  @override
  public onRenderCell(event: IFieldCustomizerCellEventParameters): void {
    const completed: boolean = event.fieldValue === 'Yes';
    const createdDate: Date = new Date(event.listItem.getValueByName('CreatedDate'));
    const completedDate: null | Date = event.listItem.getValueByName('CompletedDate') ? new Date(event.listItem.getValueByName('CompletedDate')) : null;
    console.log(createdDate, completedDate);

    event.domElement.textContent = completed ? '\u2713' : '\u2717';
    if (event.domElement.parentElement && event.domElement.parentElement.parentElement) {
      const rowElement: HTMLElement = event.domElement.parentElement.parentElement;
      if (completed) {
        rowElement.style.backgroundColor = '#bdbdbd';
      } else if (Date.now() - createdDate.getTime() > 7 * DAYS) {
        rowElement.style.backgroundColor = '#fbaa85';
      }
    }
  }
}
