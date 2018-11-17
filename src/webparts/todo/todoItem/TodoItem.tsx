import { ServiceScope } from '@microsoft/sp-core-library';
import * as React from 'react';

import { ITodoItem } from '../api';

import styles from './TodoItem.module.scss';
import { PageContext } from '@microsoft/sp-page-context';

export interface ITodoItemProps {
  serviceScope: ServiceScope;
  className: string;
  item: ITodoItem;
  updateTodoItem(todoItemId: number, completed: boolean): Promise<void>;
}

export interface ITodoItemState {
  isRequesting: boolean;
}

export class TodoItem extends React.PureComponent<ITodoItemProps, ITodoItemState> {
  public constructor(props: ITodoItemProps) {
    super(props);

    this.completeTodoItem = this.completeTodoItem.bind(this);

    this.state = {
      isRequesting: false,
    };
  }

  public render(): JSX.Element {
    const pageContext: PageContext = this.props.serviceScope.consume(PageContext.serviceKey);
    const ownerPhotoUrl: string = `${pageContext.web.absoluteUrl}/_layouts/15/UserPhoto.aspx?AccountName=${this.props.item.Owner.UserName}`;

    const activeStyle: string = this.props.item.Completed ? '' : styles.active;

    return (
      <li className={`${styles.todoItem} ${this.props.className} ${activeStyle}`}>
        <img
          className={styles.userPhoto}
          src={ownerPhotoUrl}
        />
        <div className={styles.content}>
          <p className={styles.title}>
            {this.props.item.Title}
          </p>
          <div className={styles.createdBy}>
            Created by {this.props.item.Owner.Title}
          </div>
        </div>
        {this.renderButton()}
      </li>
    );
  }

  private renderButton(): false | JSX.Element {
    return this.props.item.Completed
      ? false
      : (
        <button
          className={styles.button}
          onClick={this.completeTodoItem}
          disabled={this.state.isRequesting}
        >
          {this.state.isRequesting ? 'Updating...' : 'Complete it'}
        </button>
      );
  }

  private async completeTodoItem(): Promise<void> {
    this.setState({ isRequesting: true });
    await this.props.updateTodoItem(this.props.item.Id, /* completed */ true);
    this.setState({ isRequesting: false });
  }
}
