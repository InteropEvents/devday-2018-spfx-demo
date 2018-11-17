import { ServiceScope } from '@microsoft/sp-core-library';
import * as lodash from '@microsoft/sp-lodash-subset';
import * as React from 'react';

import styles from './TodoList.module.scss';
import { TodoItem } from '../todoItem';
import { getTodoItems, ITodoItem, updateTodoItem } from '../api';

export interface ITodoListProps {
  serviceScope: ServiceScope;
}

export interface ITodoListState {
  todoItems: undefined | ITodoItem[];
}

export class TodoList extends React.PureComponent<ITodoListProps, ITodoListState> {
  public constructor(props: ITodoListProps) {
    super(props);

    this.updateTodoItem = this.updateTodoItem.bind(this);

    this.state = {
      todoItems: undefined,
    };
  }

  public componentDidMount(): void {
    getTodoItems(this.props.serviceScope)
      .then((todoItems) => this.setState({ todoItems }))
      .catch((error) => console.error(error));
  }

  public render(): JSX.Element {
    return (
      <div>
        <h2>Here is the to-do list for my team.</h2>
        {this.renderContent()}
      </div>
    );
  }

  public renderContent(): JSX.Element {
    return this.state.todoItems
      ? (
        <ul className={styles.list}>
          {this.state.todoItems.map((item) => (
            <TodoItem
              key={item.Id}
              serviceScope={this.props.serviceScope}
              className={styles.item}
              item={item}
              updateTodoItem={this.updateTodoItem}
            />
          ))}
        </ul>
      )
      : (
        <div>Loading...</div>
      );
  }

  private async updateTodoItem(todoItemId: number, completed: boolean): Promise<void> {
    const newTodoItem = await updateTodoItem(this.props.serviceScope, todoItemId, completed);

    this.setState((state) => {
      if (state.todoItems) {
        const todoItems: ITodoItem[] = [...state.todoItems];
        const targetIndex = lodash.findIndex(todoItems, (item) => item.Id === todoItemId);
        todoItems.splice(targetIndex, 1, newTodoItem);
        return { todoItems };
      } else {
        return {};
      }
    });
  }
}
