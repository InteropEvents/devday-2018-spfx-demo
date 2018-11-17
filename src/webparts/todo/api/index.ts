import { ServiceScope } from '@microsoft/sp-core-library';
import { PageContext } from '@microsoft/sp-page-context';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

export interface ITodoOwner {
  Id: number;
  Title: string;
  UserName: string;
}

export interface ITodoItem {
  Id: number;
  Title: string;
  Completed: boolean;
  CreatedDate: Date;
  CompletedDate: null | Date;
  Owner: ITodoOwner;
}

const todoODataQuery: string = '$select=Id,Title,Completed,CreatedDate,CompletedDate,Owner/Id,Owner/Title,Owner/UserName&$expand=Owner';

export async function getTodoItems(serviceScope: ServiceScope): Promise<ITodoItem[]> {
  const pageContext: PageContext = serviceScope.consume(PageContext.serviceKey);
  const httpClient: SPHttpClient = serviceScope.consume(SPHttpClient.serviceKey);

  // Check the API document here: https://docs.microsoft.com/en-us/sharepoint/dev/sp-add-ins/working-with-lists-and-list-items-with-rest
  const endpoint: string = `${pageContext.web.absoluteUrl}/_api/web/lists/GetByTitle('Todo')/items?${todoODataQuery}`;
  const response: SPHttpClientResponse = await httpClient.get(endpoint, SPHttpClient.configurations.v1);
  const body: { value: ITodoItem[] } = await response.json();

  return body.value;
}

export async function updateTodoItem(serviceScope: ServiceScope, todoItemId: number, completed: boolean): Promise<ITodoItem> {
  const pageContext: PageContext = serviceScope.consume(PageContext.serviceKey);
  const httpClient: SPHttpClient = serviceScope.consume(SPHttpClient.serviceKey);

  const endpoint: string = `${pageContext.web.absoluteUrl}/_api/web/lists/GetByTitle('Todo')/items('${todoItemId}')?${todoODataQuery}`;
  await httpClient.post(endpoint, SPHttpClient.configurations.v1, {
    headers: {
      'IF-MATCH': '*',
      'X-HTTP-Method': 'MERGE',
      'OData-version': '3.0',
      'Accept': 'application/json;odata=verbose',
      'Content-Type': 'application/json;odata=verbose',
    },
    body: JSON.stringify({
      __metadata: {
        type: 'SP.Data.TodoListItem'
      },
      Completed: completed,
      CompletedDate: completed ? new Date().toISOString() : null,
    })
  });

  const response: SPHttpClientResponse = await httpClient.get(endpoint, SPHttpClient.configurations.v1);
  const body: ITodoItem = await response.json();

  return body;
}
